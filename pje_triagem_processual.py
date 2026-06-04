# pje_trf3_filtra_policia_v4_resiliente.py
# - Reinício preventivo do Chrome a cada N consultas
# - Backoff forte + reset hard quando travar (loading infinito/timeout)
# - Retomada automática por arquivo posicao_atual.txt (não perde progresso)
# - Mantém filtro pela LISTAGEM (tabela) amarrado ao CNJ

import re
import time
import random
import unicodedata
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    StaleElementReferenceException,
    TimeoutException,
)


# =========================================================
# 1) CONFIGURAÇÕES
# =========================================================

URL = "https://pje1g.trf5.jus.br/pjeconsulta/ConsultaPublica/listView.seam?cid=1973847"

ARQ_ENTRADA = "resposta_DATAJUD_TRF5_HC_FATO.xlsx"
COL_NUM = "numero_processo"

ARQ_SAIDA = "trf5_filtragem.xlsx"

MAX_TENTATIVAS_POR_PROCESSO = 3

CHECKPOINT_CADA = 10
ARQ_CHECKPOINT = "checkpoint_trf5.xlsx"

# Ritmo (mais conservador para volume alto)
SLEEP_MIN = 4.0
SLEEP_MAX = 9.0

PAUSA_LONGA_CADA = 15
PAUSA_LONGA_MIN = 45
PAUSA_LONGA_MAX = 120

# Resiliência para 3880 consultas
RESTART_CADA = 140
RESTART_PAUSA_MIN = 60
RESTART_PAUSA_MAX = 120

BACKOFF_TRAVA_MIN = 90
BACKOFF_TRAVA_MAX = 180

ARQ_POSICAO = "posicao_atual.txt"


# =========================================================
# 2) SELETORES (confirmados por você)
# =========================================================

INPUT_ID = "fPP:numProcesso-inputNumeroProcessoDecoration:numProcesso-inputNumeroProcesso"
BTN_ID = "fPP:searchProcessos"

SEL_RESULT_CELL = (By.CSS_SELECTOR, "td[id^='fPP:processosTable:']")

# (pode existir em outros TRFs; no TRF5 frequentemente não aparece)
SEL_MSG_NENHUM_PROCESSO = (By.CSS_SELECTOR, "dt.alert.alert-danger span.rich-messages-label")

# ✅ TRF5: contador "N resultados encontrados" (quando zero, NÃO vem número)
SEL_RESULTADOS_COUNT = (By.CSS_SELECTOR, "div#fPP\\:processosTable\\:j_id229")


# =========================================================
# 3) REGRAS DE DETECÇÃO
# =========================================================

KEYWORDS_POLICIA = [
    "POLICIA",
    "DELEGACIA", "DELEGADO", "DELEGADA",
    "POLICIA FEDERAL",
    "POLICIA CIVIL", "COMANDANTE",
    "POLICIA MILITAR",
    "PRF", "POLICIA RODOVIARIA", "RODOVIARIA FEDERAL",
    "SECRETARIA DE SEGURANCA", " AGENCIA NACIONAL DE VIGILANCIA SANITARIA"
]

MSG_PATTERNS = [
    r"sua\s+pesquisa\s+n[aã]o\s+encontrou\s+nenhum\s+processo\s+dispon[ií]vel"
]

BLOCK_PATTERNS = [
    r"access\s+denied",
    r"acesso\s+negado",
    r"temporariamente\s+indispon[ií]vel",
    r"muitas\s+requisi[cç][oõ]es",
    r"too\s+many\s+requests",
    r"forbidden",
    r"service\s+unavailable",
    r"erro\s+500",
    r"erro\s+503",
]


# =========================================================
# 4) HELPERS
# =========================================================

def cnj_format(num20: str) -> str:
    n = re.sub(r"\D", "", str(num20)).zfill(20)
    return f"{n[:7]}-{n[7:9]}.{n[9:13]}.{n[13]}.{n[14:16]}.{n[16:20]}"

def normaliza(txt: str) -> str:
    txt = (txt or "")
    txt = txt.replace("\xa0", " ")
    txt = unicodedata.normalize("NFKD", txt)
    txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
    txt = re.sub(r"\s+", " ", txt).strip().upper()
    return txt

def tem_policia(texto: str) -> bool:
    t = normaliza(texto)
    return any(k in t for k in KEYWORDS_POLICIA)

def body_text(driver) -> str:
    try:
        return driver.find_element(By.TAG_NAME, "body").text
    except Exception:
        return ""

def pagina_tem_mensagem_nenhum_processo(driver) -> bool:
    els = driver.find_elements(*SEL_MSG_NENHUM_PROCESSO)
    if els:
        txt = (els[0].text or "").strip().lower()
        return "sua pesquisa não encontrou nenhum processo disponível".lower() in txt

    txt2 = body_text(driver).lower()
    return re.search(MSG_PATTERNS[0], txt2) is not None

def detectou_bloqueio(driver) -> str | None:
    txt = body_text(driver).lower()
    input_exists = len(driver.find_elements(By.ID, INPUT_ID)) > 0

    for pat in BLOCK_PATTERNS:
        if re.search(pat, txt):
            return pat

    if not input_exists:
        return "input_missing"

    return None

def recuperar_acesso(driver, wait_plan=(30, 60, 120, 240, 300)) -> bool:
    for sec in wait_plan:
        motivo = detectou_bloqueio(driver)
        if motivo is None:
            return True

        print(f"🛑 PJe parece indisponível/bloqueado ({motivo}). Esperando {sec}s e recarregando...")
        time.sleep(sec)

        try:
            driver.get(URL)
            WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.ID, INPUT_ID)))
        except Exception as e:
            print("…recarregar falhou:", repr(e))

    return detectou_bloqueio(driver) is None

def normalizar_20_digitos(raw) -> str:
    s = str(raw).strip()
    dig = re.sub(r"\D", "", s)
    return dig.zfill(20)

def salvar_posicao(i: int):
    try:
        with open(ARQ_POSICAO, "w", encoding="utf-8") as f:
            f.write(str(i))
    except Exception:
        pass

def ler_posicao() -> int:
    try:
        with open(ARQ_POSICAO, "r", encoding="utf-8") as f:
            v = int(f.read().strip())
            return max(1, v)
    except Exception:
        return 1


# -----------------------------
# Helpers: loading/overlay
# -----------------------------

def wait_page_idle(driver, timeout=25):
    wait = WebDriverWait(driver, timeout)

    try:
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    except Exception:
        pass

    try:
        wait.until(lambda d: d.execute_script("return (window.jQuery ? jQuery.active : 0)") == 0)
    except Exception:
        pass

    overlay_css = [
        ".rf-pp-shade", ".rf-pp-cntr", ".rf-pp-shdw",
        ".rich-mpnl-mask", ".rich-mpnl-panel",
        ".blockUI", ".blockUI.blockOverlay",
        ".loading", ".spinner", ".ajaxStatus",
        ".rich-mp-container", ".rich-mpnl-content",
    ]
    for css in overlay_css:
        try:
            wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, css)))
        except Exception:
            pass

def safe_click(driver, by, locator, timeout=25):
    wait_page_idle(driver, timeout=timeout)
    wait = WebDriverWait(driver, timeout)

    el = wait.until(EC.presence_of_element_located((by, locator)))
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    except Exception:
        pass
    time.sleep(0.1)

    try:
        el = wait.until(EC.element_to_be_clickable((by, locator)))
        el.click()
        return
    except ElementClickInterceptedException:
        wait_page_idle(driver, timeout=timeout)
        el = driver.find_element(by, locator)
        driver.execute_script("arguments[0].click();", el)
        return

def safe_type_masked_digits(driver, input_id, digits20, timeout=25, per_char_sleep=0.08) -> str:
    wait_page_idle(driver, timeout=timeout)
    wait = WebDriverWait(driver, timeout)

    el = wait.until(EC.presence_of_element_located((By.ID, input_id)))
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    except Exception:
        pass
    time.sleep(0.1)

    try:
        el.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].focus();", el)

    el.send_keys(Keys.CONTROL, "a")
    el.send_keys(Keys.BACKSPACE)
    time.sleep(0.15)

    for ch in digits20:
        el.send_keys(ch)
        time.sleep(per_char_sleep)

    time.sleep(0.25)
    wait_page_idle(driver, timeout=timeout)

    return (el.get_attribute("value") or "").strip()


# -----------------------------
# ✅ TRF5: ler quantidade de resultados pelo contador
# -----------------------------

def ler_qtd_resultados_trf5(driver, timeout=25) -> int:
    """
    TRF5: quando não encontra nada, aparece só "resultados encontrados" (sem número).
    Quando encontra, costuma aparecer "N resultados encontrados".
    Retorna:
      - >=0: quantidade
      - -1: não conseguiu ler o elemento
    """
    end = time.time() + timeout
    last_txt = None
    stable_hits = 0

    # tenta esperar estabilizar (evita ler estado antigo)
    while time.time() < end:
        els = driver.find_elements(*SEL_RESULTADOS_COUNT)
        if els:
            txt = (els[0].text or "").strip().lower()

            if txt == last_txt and txt != "":
                stable_hits += 1
                if stable_hits >= 3:  # ~0.6s estável
                    break
            else:
                stable_hits = 0
                last_txt = txt
        time.sleep(0.2)

    els = driver.find_elements(*SEL_RESULTADOS_COUNT)
    if not els:
        return -1

    txt = (els[0].text or "").strip().lower()
    m = re.search(r"\b(\d+)\b", txt)
    return int(m.group(1)) if m else 0


def esperar_resultado_ou_mensagem(driver, timeout=25) -> str:
    """
    Retorna: ok | inacessivel | blocked | timeout
    - ok: achou celula de resultado
    - inacessivel: TRF5 silencioso (qtd==0) ou mensagem (quando existir)
    - blocked: padrões de bloqueio/indisponibilidade
    - timeout: não houve nenhum sinal claro
    """
    end = time.time() + timeout
    while time.time() < end:
        if detectou_bloqueio(driver) is not None:
            return "blocked"

        # 1) resultado (tabela/células)
        if driver.find_elements(*SEL_RESULT_CELL):
            return "ok"

        # 2) mensagem (quando existir)
        if pagina_tem_mensagem_nenhum_processo(driver):
            return "inacessivel"

        # 3) ✅ contador TRF5
        qtd = ler_qtd_resultados_trf5(driver, timeout=2)
        if qtd == 0:
            return "inacessivel"
        if qtd > 0:
            # se já sabe que tem >0, dá um respiro pra tabela renderizar
            time.sleep(0.3)

        time.sleep(0.5)

    return "timeout"


def get_cell_text(cell) -> str:
    return (cell.get_attribute("innerText")
            or cell.get_attribute("textContent")
            or cell.text
            or "")

def extrair_texto_da_listagem_por_cnj(driver, cnj_esperado: str, timeout=40) -> str:
    end = time.time() + timeout
    cnj_up = cnj_esperado.upper()

    while time.time() < end:
        if detectou_bloqueio(driver) is not None:
            return ""

        if pagina_tem_mensagem_nenhum_processo(driver):
            return ""

        # ✅ TRF5: se contador diz zero, não adianta procurar CNJ
        qtd = ler_qtd_resultados_trf5(driver, timeout=2)
        if qtd == 0:
            return ""

        cells = driver.find_elements(*SEL_RESULT_CELL)
        for c in cells[:60]:
            t = get_cell_text(c)
            if cnj_up in (t or "").upper():
                return t

        time.sleep(0.4)

    return ""

def salvar_excel(candidatos_policia, inacessiveis, timeouts, erros, caminho):
    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        pd.DataFrame(candidatos_policia).to_excel(writer, sheet_name="candidatos_policia", index=False)
        pd.DataFrame(inacessiveis).to_excel(writer, sheet_name="inacessiveis", index=False)
        pd.DataFrame(timeouts).to_excel(writer, sheet_name="timeouts", index=False)
        pd.DataFrame(erros).to_excel(writer, sheet_name="erros", index=False)

def build_options():
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    return options

def start_driver():
    driver = webdriver.Chrome(options=build_options())
    wait = WebDriverWait(driver, 25)
    driver.get(URL)
    wait.until(EC.presence_of_element_located((By.ID, INPUT_ID)))
    wait_page_idle(driver, timeout=25)
    return driver, wait

def reset_hard(driver, motivo="trava"):
    print(f"🧯 Reset hard ({motivo}): fechando navegador + backoff…")
    try:
        driver.quit()
    except Exception:
        pass
    time.sleep(random.uniform(BACKOFF_TRAVA_MIN, BACKOFF_TRAVA_MAX))
    return start_driver()


# =========================================================
# 5) MAIN
# =========================================================

def main():
    df = pd.read_excel(ARQ_ENTRADA, dtype={COL_NUM: str})
    if COL_NUM not in df.columns:
        raise ValueError(f"Coluna '{COL_NUM}' não existe. Colunas: {list(df.columns)}")

    numeros_raw = df[COL_NUM].tolist()
    total = len(numeros_raw)

    candidatos_policia = []
    inacessiveis = []
    timeouts = []
    erros = []

    start_i = ler_posicao()
    print(f"▶️ Retomada: começando do item {start_i}/{total} (arquivo {ARQ_POSICAO})")

    driver, wait = start_driver()

    try:
        for i, raw in enumerate(numeros_raw, start=1):
            if i < start_i:
                continue

            salvar_posicao(i)

            # Reinício preventivo
            if i > 1 and RESTART_CADA and i % RESTART_CADA == 0:
                print(f"🔄 Reiniciando navegador preventivamente no item {i}...")
                try:
                    driver.quit()
                except Exception:
                    pass
                time.sleep(random.uniform(RESTART_PAUSA_MIN, RESTART_PAUSA_MAX))
                driver, wait = start_driver()

            tent = 0
            num20 = normalizar_20_digitos(raw)
            if len(num20) != 20:
                erros.append({"numero_processo": str(raw), "erro": f"numero_len_{len(num20)}"})
                continue

            cnj = cnj_format(num20)

            while tent < MAX_TENTATIVAS_POR_PROCESSO:
                tent += 1
                try:
                    print(f"\n[{i}/{total}] {raw} -> {num20} ({cnj}) (tentativa {tent}/{MAX_TENTATIVAS_POR_PROCESSO})")

                    wait_page_idle(driver, timeout=25)

                    if detectou_bloqueio(driver) is not None:
                        if not recuperar_acesso(driver):
                            erros.append({"numero_processo": str(raw), "erro": "bloqueio_persistente"})
                            print("❌ Bloqueio persistente. Pulando.")
                            break
                        wait_page_idle(driver, timeout=25)

                    valor = safe_type_masked_digits(driver, INPUT_ID, num20, timeout=25, per_char_sleep=0.08)

                    if len(valor) < 12:
                        print(f"⚠️ Preenchimento truncou: '{valor}'. Tentando novamente...")
                        continue

                    # best effort: tentar invalidar resultado antigo
                    try:
                        WebDriverWait(driver, 3).until(EC.invisibility_of_element_located(SEL_RESULT_CELL))
                    except Exception:
                        pass

                    safe_click(driver, By.ID, BTN_ID, timeout=25)

                    status = esperar_resultado_ou_mensagem(driver, timeout=25)

                    if status == "blocked":
                        print("🛑 Bloqueio detectado. Tentando recuperar…")
                        if recuperar_acesso(driver):
                            driver.get(URL)
                            wait.until(EC.presence_of_element_located((By.ID, INPUT_ID)))
                            continue
                        erros.append({"numero_processo": str(raw), "erro": "bloqueio_persistente_pos_pesquisa"})
                        driver, wait = reset_hard(driver, motivo="bloqueio_persistente")
                        break

                    if status == "inacessivel":
                        inacessiveis.append({"numero_processo": str(raw), "motivo": "zero_resultados_trf5_ou_indisponivel"})
                        print("🔒 Zero resultados (TRF5 silencioso) ou nenhum processo disponível.")
                        break

                    if status == "timeout":
                        print("⏳ Timeout (sem resultado nem mensagem/contador). Reset hard e repetindo…")
                        timeouts.append({"numero_processo": str(raw), "motivo": "timeout_sem_resultado"})
                        driver, wait = reset_hard(driver, motivo="timeout_sem_resultado")
                        continue

                    texto = extrair_texto_da_listagem_por_cnj(driver, cnj_esperado=cnj, timeout=40)

                    if not texto:
                        print("⏳ Não achei o CNJ na listagem (provável travamento/resultado antigo). Reset hard e repetindo…")
                        timeouts.append({"numero_processo": str(raw), "motivo": "cnj_nao_apareceu_na_listagem"})
                        driver, wait = reset_hard(driver, motivo="cnj_nao_apareceu")
                        continue

                    if tem_policia(texto):
                        candidatos_policia.append({"numero_processo": str(raw), "trecho": texto})
                        print("✅ Candidato (polícia/delegado detectado).")
                    else:
                        print("- Sem polícia/delegado no resultado.")

                    break  # sucesso neste processo

                except ElementClickInterceptedException:
                    print("🧱 Clique interceptado (loading/overlay). Reset hard e repetindo…")
                    driver, wait = reset_hard(driver, motivo="click_intercepted")
                    continue

                except (StaleElementReferenceException,):
                    print("🔁 DOM mudou (stale). Repetindo tentativa…")
                    continue

                except Exception as e:
                    if detectou_bloqueio(driver) is not None and recuperar_acesso(driver):
                        continue
                    erros.append({"numero_processo": str(raw), "erro": repr(e)})
                    print("❌ Erro:", repr(e))
                    driver, wait = reset_hard(driver, motivo="erro_generico")
                    break

            time.sleep(random.uniform(SLEEP_MIN, SLEEP_MAX))
            if PAUSA_LONGA_CADA and i % PAUSA_LONGA_CADA == 0:
                time.sleep(random.uniform(PAUSA_LONGA_MIN, PAUSA_LONGA_MAX))

            if CHECKPOINT_CADA and i % CHECKPOINT_CADA == 0:
                salvar_excel(candidatos_policia, inacessiveis, timeouts, erros, ARQ_CHECKPOINT)
                print(f"💾 Checkpoint salvo: {ARQ_CHECKPOINT}")

    finally:
        try:
            driver.quit()
        except Exception:
            pass

    salvar_excel(candidatos_policia, inacessiveis, timeouts, erros, ARQ_SAIDA)
    print(f"\n✅ Finalizado. Saída: {ARQ_SAIDA}")
    print(f"🧾 Posição final registrada em {ARQ_POSICAO} (você pode apagar para recomeçar do 1).")


if __name__ == "__main__":
    main()
