import os, sys, re, time, random, threading, pathlib, unicodedata, logging
from datetime import datetime, timedelta
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image
from playwright.sync_api import sync_playwright
from twocaptcha import TwoCaptcha
import fdb

# --------------------------------------------------------------------------------------
# Carrega fbclient.dll em dev ou empacotado
# --------------------------------------------------------------------------------------
def _carregar_fbclient():
    candidatos, add_dirs = [], set()
    env = [os.getenv("FIREBIRD_CLIENT_PATH", ""), os.getenv("FBCLIENT_PATH", "")]
    candidatos += [p for p in env if p]
    add_dirs.update([os.path.dirname(p) for p in candidatos if p])

    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
        exe_dir = os.path.dirname(sys.executable)
        candidatos += [
            os.path.join(base, "fbclient.dll"),
            os.path.join(exe_dir, "fbclient.dll"),
            os.path.join(base, "bin", "fbclient.dll"),
        ]
        add_dirs.update([base, exe_dir, os.path.join(base, "plugins"), os.path.join(base, "bin")])
    else:
        comuns = [
            r"C:\Program Files\Firebird\Firebird_5_0\bin\fbclient.dll",
            r"C:\Program Files\Firebird\Firebird_4_0\bin\fbclient.dll",
            r"C:\Program Files (x86)\Firebird\Firebird_5_0\bin\fbclient.dll",
            r"C:\Program Files (x86)\Firebird\Firebird_4_0\bin\fbclient.dll",
        ]
        this_dir = pathlib.Path(__file__).resolve().parent
        cwd = pathlib.Path.cwd()
        locais = [
            str(this_dir / "fbclient.dll"),
            str(cwd / "fbclient.dll"),
            str(this_dir / "bin" / "fbclient.dll"),
            str(cwd / "bin" / "fbclient.dll"),
        ]
        candidatos += comuns + locais
        add_dirs.update([os.path.dirname(p) for p in comuns] + [str(this_dir), str(cwd), str(this_dir / "bin"), str(cwd / "bin")])

    candidatos.append("fbclient.dll")
    for d in list(add_dirs):
        try:
            if d and os.path.isdir(d):
                os.add_dll_directory(d)
        except Exception:
            pass

    for p in candidatos:
        try:
            if os.path.isabs(p):
                if os.path.exists(p):
                    fdb.load_api(p); return p
            else:
                fdb.load_api(p); return p
        except Exception:
            continue

    raise RuntimeError(
        "fbclient.dll não encontrada. Instale o Firebird Client (mesma arquitetura do Python) "
        "ou defina FIREBIRD_CLIENT_PATH com o caminho completo."
    )

# --------------------------------------------------------------------------------------
# Estado Parte II
# --------------------------------------------------------------------------------------
DF_CNPJ_PARTE2 = None
CAMINHO_CNPJ_PARTE2 = ""
MES_PASTA = ""
RAIZ_PASTAS2 = ""

def selecionar_planilha_cnpj_parte2():
    global DF_CNPJ_PARTE2, CAMINHO_CNPJ_PARTE2, RAIZ_PASTAS2, MES_PASTA
    caminho = filedialog.askopenfilename(title="Selecione a planilha CNPJ.xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])
    if not caminho:
        return
    try:
        df = pd.read_excel(caminho)
        if "CNPJ" not in df.columns or "Nome" not in df.columns:
            messagebox.showerror("Erro", "A planilha precisa ter as colunas 'CNPJ' e 'Nome'.")
            return
        DF_CNPJ_PARTE2 = df.copy()
        CAMINHO_CNPJ_PARTE2 = caminho
        RAIZ_PASTAS2 = os.path.dirname(caminho)
        _, ultimo = primeiro_e_ultimo_dia_mes_anterior()
        MES_PASTA = ultimo.strftime("%m.%Y")

        quadro_log2.configure(state="normal")
        quadro_log2.insert("end", f"Planilha CNPJ carregada: {os.path.basename(caminho)}\n")
        quadro_log2.insert("end", f"Raiz base: {RAIZ_PASTAS2}\n")
        quadro_log2.insert("end", f"Pasta alvo: {MES_PASTA}\n")
        quadro_log2.configure(state="disabled")
        quadro_log2.see("end")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao ler a planilha: {e}")

def limpar_nome(s):
    s = unicodedata.normalize('NFKD', str(s)).encode('ASCII', 'ignore').decode('ASCII')
    s = re.sub(r'[^\w\s-]', '', s).strip()
    return re.sub(r'[-\s]+', '_', s)

def limpar_nome_para_chave(nome: str) -> str:
    nome = unicodedata.normalize('NFKD', str(nome)).encode('ASCII', 'ignore').decode('ASCII')
    nome = re.sub(r'[^\w\s.-]', '', nome).strip()
    nome = re.sub(r'[\s-]+', '_', nome)
    return nome.upper()

def normalizar_nome_pasta_variantes(nome: str):
    base = limpar_nome(nome)
    return list(dict.fromkeys([base, base.upper(), base.lower(), limpar_nome_para_chave(nome)]))

def _normkey_nome(s: str) -> str:
    return re.sub(r'[^A-Z0-9]', '', limpar_nome_para_chave(s).upper())

def listar_pastas_por_planilha(raiz: str, df_cnpj: pd.DataFrame, quadro_log_widget=None) -> list[str]:
    encontrados, faltantes = [], []
    subdirs = [d for d in os.listdir(raiz) if os.path.isdir(os.path.join(raiz, d))]
    mapa_norm = {_normkey_nome(d): d for d in subdirs}
    subs_lower = {d.lower(): d for d in subdirs}

    def tem_planilha(p):
        for base in ("nfedestinatario.xlsx", "nfe.xlsx", "nfce.xlsx"):
            if os.path.exists(os.path.join(p, base)):
                return True
        return False

    for _, r in df_cnpj.iterrows():
        nome = str(r["Nome"])
        pasta = None
        key = _normkey_nome(nome)
        if key in mapa_norm:
            pasta = os.path.join(raiz, mapa_norm[key])
        if not pasta:
            for v in normalizar_nome_pasta_variantes(nome):
                cand = os.path.join(raiz, v)
                if os.path.isdir(cand): pasta = cand; break
                if v.lower() in subs_lower: pasta = os.path.join(raiz, subs_lower[v.lower()]); break
        if pasta:
            (encontrados if tem_planilha(pasta) else faltantes).append(pasta if tem_planilha(pasta) else (pasta, "sem planilhas"))
        else:
            faltantes.append((nome, "pasta não encontrada"))

    if quadro_log_widget is not None:
        quadro_log_widget.configure(state="normal")
        quadro_log_widget.insert("end", f"Pastas encontradas: {len(encontrados)}\n")
        for item, motivo in faltantes[:30]:
            quadro_log_widget.insert("end", f"  Não usada: {item} ({motivo})\n")
        quadro_log_widget.configure(state="disabled")
        quadro_log_widget.see("end")

    return encontrados

# --------------------------------------------------------------------------------------
# Log
# --------------------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler("cte_automacao.log"), logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# --------------------------------------------------------------------------------------
# Parte III – CTe
# --------------------------------------------------------------------------------------
def mes_anterior_formatado():
    hoje = datetime.now()
    primeiro = hoje.replace(day=1)
    ultimo = primeiro - timedelta(days=1)
    return ultimo.strftime("%m/%Y")

def emitir_cte(context, quadro_log3, barra_progresso3):
    try:
        login, senha = open('credenciais.txt', 'r').readline().strip().split(',')
        periodo = mes_anterior_formatado()
        pasta_destino = (datetime.now().replace(day=1) - timedelta(days=1)).strftime("%m.%Y")
        os.makedirs(pasta_destino, exist_ok=True)

        cn = pd.read_excel('CNPJ.xlsx')
        quadro_log3.insert("end", f"Planilha: {len(cn)} CNPJs\n"); quadro_log3.see("end")

        page = context.new_page()
        page.goto("https://sat.sef.sc.gov.br/tax.NET/Login.aspx?ReturnUrl=%2ftax.NET%2f")
        page.wait_for_load_state("networkidle")
        simulate_human_delay(2, 5)
        page.locator('#Body_pnlMain_tbxUsername').fill(login)
        simulate_human_delay(1, 3)
        page.locator('#Body_pnlMain_tbxUserPassword').fill(senha)
        simulate_human_delay(1, 3)
        page.locator('#Body_pnlMain_btnLogin > span').click()
        simulate_human_delay(2, 4)

        page.wait_for_selector('#Body_Main_ctl09_ctl07_rptAppList_ctl03_1 > span')
        page.locator('#s2id_Body_ApplicationMasterHeader_ApplicationMasterSearchAppsInput_txtSearchApp_hid_single_txtSearchApp_value > a').click()
        page.fill('#select2-drop > div > input', 'CTe - consulta')
        with context.expect_page() as new_page_info:
            page.locator('#select2-drop > ul > li > div').click()
        new_page = new_page_info.value
        new_page.wait_for_load_state("networkidle")
        time.sleep(2)

        new_page.get_by_role("link", name="Consulta CTe por emitente").click()
        new_page.locator("#select2-drop").get_by_text("Consulta CTe por tomador").click()
        new_page.get_by_placeholder("mm/aaaa").fill(periodo)

        for _, row in cn.iterrows():
            cnpj = str(row['CNPJ']).zfill(14)
            nome = limpar_nome(row['Nome'])
            quadro_log3.insert("end", f"Processando: {nome} ({cnpj})\n"); quadro_log3.see("end")
            try:
                time.sleep(1)
                new_page.locator('#s2id_Body_Main_Main_sepBusca_ctbContribuinte_hid_single_ctbContribuinte_value > a > span.select2-chosen').click()
                new_page.locator("#select2-drop").get_by_role("textbox").fill(cnpj)
                new_page.wait_for_selector('#select2-drop > ul > li > div')
                new_page.locator('#select2-drop > ul > li > div').click()

                new_page.locator('#Body_Main_Main_sepBusca_btnPesquisar > span').click()
                time.sleep(2)
                new_page.locator("#Body_Main_Main_grpCTe_btn0").click()

                pasta = os.path.join(pasta_destino, nome)
                os.makedirs(pasta, exist_ok=True)
                with new_page.expect_download() as dlinfo:
                    new_page.locator('#btn0').first.click()
                dlinfo.value.save_as(os.path.join(pasta, 'CTE.xlsx'))
                quadro_log3.insert("end", f"OK: {nome}\n"); quadro_log3.see("end")
            except Exception as e:
                quadro_log3.insert("end", f"Erro com {nome}: {e}\n"); quadro_log3.see("end")
                try:
                    if new_page.locator("#__SatMessageBox").is_visible():
                        new_page.locator("#__SatMessageBox > button").click()
                except:
                    pass
                continue
    except Exception as e:
        quadro_log3.insert("end", f"Erro geral: {e}\n"); quadro_log3.see("end")

def rodar_emissao_cte_thread():
    threading.Thread(target=executar_emissao_cte, daemon=True).start()

def executar_emissao_cte():
    quadro_log3.configure(state="normal")
    quadro_log3.insert("end", "Iniciando emissão de CTEs...\n"); quadro_log3.see("end")
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False, channel='chrome')
            context = browser.new_context()
            emitir_cte(context, quadro_log3, barra_progresso3)
        quadro_log3.insert("end", "Emissão finalizada.\n")
    except Exception as e:
        quadro_log3.insert("end", f"Erro na emissão: {e}\n")
    quadro_log3.configure(state="disabled")

# --------------------------------------------------------------------------------------
# Auxiliares gerais
# --------------------------------------------------------------------------------------
def simulate_human_delay(a=1, b=3):
    time.sleep(random.uniform(a, b))

def primeiro_e_ultimo_dia_mes_anterior():
    hoje = datetime.now()
    primeiro_atual = hoje.replace(day=1)
    ultimo_ant = primeiro_atual - timedelta(days=1)
    primeiro_ant = ultimo_ant.replace(day=1)
    return primeiro_ant, ultimo_ant

def setar_datas_nfe(page, ini, fim):
    sel_ini = "#Body_Main_Main_sepConsultaNfpe_datDataInicial"
    sel_fim = "#Body_Main_Main_sepConsultaNfpe_datDataFinal"

    def preencher(sel, val):
        try:
            el = page.locator(sel)
            el.click()
            page.keyboard.press("Control+A"); page.keyboard.press("Delete")
            el.fill(val); time.sleep(0.15)
            try:
                lido = el.input_value()
            except:
                lido = page.evaluate("(s)=>document.querySelector(s)?.value||''", sel)
            return lido == val
        except:
            return False

    ok1, ok2 = preencher(sel_ini, ini), preencher(sel_fim, fim)
    if not (ok1 and ok2):
        page.evaluate(
            """(iSel, fSel, i, f) => {
                const fire = el => { el.dispatchEvent(new Event('input',{bubbles:true}));
                                     el.dispatchEvent(new Event('change',{bubbles:true}));
                                     el.dispatchEvent(new Event('blur',{bubbles:true})); };
                const iel = document.querySelector(iSel), fel = document.querySelector(fSel);
                if (iel) { iel.value = i; fire(iel); }
                if (fel) { fel.value = f; fire(fel); }
            }""",
            sel_ini, sel_fim, ini, fim
        )
        time.sleep(0.15)

# --------------------------------------------------------------------------------------
# Parte I – SAT (NF-e/NFC-e)
# --------------------------------------------------------------------------------------
def login_sistema(context, login, senha):
    page = context.new_page()
    page.goto("https://sat.sef.sc.gov.br/tax.NET/Login.aspx?ReturnUrl=%2ftax.NET%2f")
    page.wait_for_load_state("networkidle")
    simulate_human_delay(2, 4)
    page.locator('#Body_pnlMain_tbxUsername').fill(login)
    page.locator('#Body_pnlMain_tbxUserPassword').fill(senha)
    page.locator('#Body_pnlMain_btnLogin > span').click()
    simulate_human_delay(2, 4)
    page.wait_for_selector('#Body_Main_ctl09_ctl07_rptAppList_ctl03_1 > span')
    page.locator('#s2id_Body_ApplicationMasterHeader_ApplicationMasterSearchAppsInput_txtSearchApp_hid_single_txtSearchApp_value > a > span.select2-chosen > span').click()
    page.fill('#select2-drop > div > input', 'NFe / NFCe - Consulta')
    with context.expect_page() as new_page_info:
        page.locator('#select2-drop > ul > li > div').click()
    return page, new_page_info.value

def resolver_captcha(new_page):
    new_page.get_by_role("img").screenshot(path="captcha.png")
    api_key = os.getenv('APIKEY_2CAPTCHA', '07353710a01aaea98dea3888aeca8a47')
    result = TwoCaptcha(api_key).normal("captcha.png")
    new_page.fill('#Body_Main_Main_sepConsultaNfpe_ctl21 > div > input', result['code'])
    new_page.locator('#Body_Main_Main_sepConsultaNfpe_btnBuscar > span').click()
    time.sleep(3)

def processar_cnpj(new_page, cnpj, nome_empresa, pasta_destino, primeiro, ultimo, quadro_log, barra_progresso, total, atual):
    cnpj = str(cnpj).zfill(14)
    pasta = os.path.join(pasta_destino, nome_empresa)
    os.makedirs(pasta, exist_ok=True)

    ini, fim = primeiro.strftime("%d/%m/%Y"), ultimo.strftime("%d/%m/%Y")

    def captcha_visivel():
        try:
            return new_page.locator('#Body_Main_Main_sepConsultaNfpe_ctl17 > img').is_visible()
        except:
            return False

    def buscar():
        if captcha_visivel(): resolver_captcha(new_page)
        else:
            try:
                new_page.locator('#Body_Main_Main_sepConsultaNfpe_btnBuscar > span').click()
                time.sleep(0.8)
            except: pass

    def exportar(destino):
        buscar()
        if captcha_visivel(): resolver_captcha(new_page)
        with new_page.expect_download(timeout=60000) as dlinfo:
            new_page.locator('#Body_Main_Main_sepConsultaNfpe_btnExportar > span').click()
        dlinfo.value.save_as(destino)

    def filtrar_cols(caminho):
        cols = ["NomeEmitente","TipoDeOperacaoEntradaOuSaida","Situacao","ChaveAcesso","DataEmissao",
                "NumeroDocumento","ValorTotalNota","ValorTotalISS","ValorTotalICMS","ValorBaseCalculoICMS"]
        try:
            df = pd.read_excel(caminho)
            keep = [c for c in cols if c in df.columns]
            if keep: df[keep].to_excel(caminho, index=False)
        except Exception as e:
            quadro_log.insert("end", f"Erro ao filtrar {os.path.basename(caminho)}: {e}\n")

    try:
        for sel in ('#Body_Main_Main_sepConsultaNfpe_ctl10_idnEmitente_MaskedField',
                    '#Body_Main_Main_sepConsultaNfpe_ctl11_idnDestinatario_MaskedField'):
            try: new_page.fill(sel, '')
            except: pass

        # Emitente (NF-e)
        new_page.fill('#Body_Main_Main_sepConsultaNfpe_ctl10_idnEmitente_MaskedField', cnpj)
        setar_datas_nfe(new_page, ini, fim)
        arq = os.path.join(pasta, 'nfe.xlsx'); exportar(arq); filtrar_cols(arq)

        # Destinatário (NF-e)
        new_page.locator('#Body_Main_Main_sepConsultaNfpe_ctl10_idnEmitente_MaskedField').select_text(); new_page.locator('#Body_Main_Main_sepConsultaNfpe_ctl10_idnEmitente_MaskedField').press('Delete')
        new_page.locator('#Body_Main_Main_sepConsultaNfpe_ctl11_idnDestinatario_MaskedField').fill(cnpj)
        setar_datas_nfe(new_page, ini, fim)
        exportar(os.path.join(pasta, 'nfedestinatario.xlsx'))
        new_page.locator('#Body_Main_Main_sepConsultaNfpe_ctl11_idnDestinatario_MaskedField').select_text(); new_page.locator('#Body_Main_Main_sepConsultaNfpe_ctl11_idnDestinatario_MaskedField').press('Delete')

        # NFC-e (emitente e destinatário)
        new_page.get_by_role("link", name="NF-e").click()
        new_page.locator("#select2-drop").get_by_text("NFC-e").click()

        new_page.fill('#Body_Main_Main_sepConsultaNfpe_ctl10_idnEmitente_MaskedField', cnpj)
        setar_datas_nfe(new_page, ini, fim)
        exportar(os.path.join(pasta, 'nfce.xlsx'))
        new_page.locator('#Body_Main_Main_sepConsultaNfpe_ctl10_idnEmitente_MaskedField').select_text(); new_page.locator('#Body_Main_Main_sepConsultaNfpe_ctl10_idnEmitente_MaskedField').press('Delete')

        new_page.locator('#Body_Main_Main_sepConsultaNfpe_ctl11_idnDestinatario_MaskedField').fill(cnpj)
        setar_datas_nfe(new_page, ini, fim)
        exportar(os.path.join(pasta, 'nfcedestinatario.xlsx'))

        quadro_log.insert("end", f"{cnpj} finalizado.\n")
        new_page.get_by_role("link", name="NFC-e").click()
        new_page.locator("#select2-drop").get_by_text("NF-e").click()
    except Exception as e:
        quadro_log.insert("end", f"Erro com {cnpj}: {e}\n")

    quadro_log.see("end")
    barra_progresso.set(atual / total)

def iniciar_automacao(df, quadro_log, barra_progresso):
    quadro_log.configure(state="normal")
    quadro_log.insert("end", "Iniciando automação...\n"); quadro_log.see("end")

    primeiro, ultimo = primeiro_e_ultimo_dia_mes_anterior()
    pasta_destino = ultimo.strftime("%m.%Y")
    os.makedirs(pasta_destino, exist_ok=True)

    with open("credenciais.txt", "r") as f:
        login, senha = f.readline().strip().split(",")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel='chrome')
        context = browser.new_context()
        page, new_page = login_sistema(context, login, senha)

        total = len(df)
        for i, row in enumerate(df.itertuples(), 1):
            nome = limpar_nome(row.Nome)
            cnpj = str(row.CNPJ).zfill(14)
            processar_cnpj(new_page, cnpj, nome, pasta_destino, primeiro, ultimo, quadro_log, barra_progresso, total, i)
            time.sleep(2)

        new_page.close(); page.close(); browser.close()
    quadro_log.insert("end", "Automação finalizada.\n")
    quadro_log.configure(state="disabled")

# --------------------------------------------------------------------------------------
# Parte II – Conciliação
# --------------------------------------------------------------------------------------
NOME_NF = "nfedestinatario.xlsx"
ARQ_CNPJ_RAIZ = "CNPJ.xlsx"
TOLERANCIA_REAIS = 0.00
EXCEL_ENGINE = None

COLUNAS_FINAIS = [
    "NomeEmitente","TipoDeOperacaoEntradaOuSaida","Situacao","ChaveAcesso",
    "DataEmissao","NumeroDocumento","ValorTotalNota","ValorTotalISS",
    "ValorTotalICMS","ValorBaseCalculoICMS","Verificação"
]

def _escolher_engine_excel():
    global EXCEL_ENGINE
    if EXCEL_ENGINE: return EXCEL_ENGINE
    try:
        import xlsxwriter; EXCEL_ENGINE = "xlsxwriter"
    except Exception:
        try:
            import openpyxl; EXCEL_ENGINE = "openpyxl"
        except Exception:
            EXCEL_ENGINE = None
    return EXCEL_ENGINE

def salvar_dataframe(df, out_path, sheet_name="Dados"):
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    eng = _escolher_engine_excel()
    if eng is None:
        csv_path = os.path.splitext(out_path)[0] + ".csv"
        df.to_csv(csv_path, index=False, sep=";", encoding="utf-8")
        return csv_path
    with pd.ExcelWriter(out_path, engine=eng) as w:
        df.to_excel(w, index=False, sheet_name=sheet_name)
    return out_path

def so_digitos(x) -> str:
    return re.sub(r"\D", "", str(x or ""))

def parse_money(v):
    import numpy as np
    if pd.isna(v): return np.nan
    s = str(v).strip().replace(" ", "")
    try: return float(s)
    except: pass
    if "," in s and "." in s:
        s = s.replace(".", "") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
        s = s.replace(",", ".")
    elif "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try: return float(s)
    except: return np.nan

def marcar_lancado_series(series_val_nf, set_db_centavos, tol_centavos=0):
    out = []
    for v in series_val_nf:
        if pd.isna(v): out.append("Não Lançado"); continue
        cents = int(round(float(v) * 100))
        if tol_centavos == 0:
            out.append("Lançado" if cents in set_db_centavos else "Não Lançado")
        else:
            ok = any((cents + d) in set_db_centavos for d in range(-tol_centavos, tol_centavos+1))
            out.append("Lançado" if ok else "Não Lançado")
    return out

def garantir_colunas(df, cols):
    import numpy as np
    for c in cols:
        if c not in df.columns and c != "Verificação":
            df[c] = np.nan
    return df

def faixa_mes_anterior():
    hoje = datetime.now()
    p_atual = hoje.replace(day=1)
    u_ant = p_atual - timedelta(days=1)
    p_ant = u_ant.replace(day=1)
    return p_ant, u_ant

def inferir_periodo(df_nf: pd.DataFrame, dirpath: str):
    for c in ["DataEmissao","Data_Emissao","Data","Emissao"]:
        if c in df_nf.columns:
            try:
                datas = pd.to_datetime(df_nf[c], dayfirst=True, errors="coerce").dropna()
                if len(datas): return datas.min().normalize(), datas.max().normalize()
            except: pass
    m = re.search(r"(\d{2})[._-](\d{4})", os.path.basename(dirpath))
    if m:
        mes, ano = int(m.group(1)), int(m.group(2))
        ini = datetime(ano, mes, 1)
        prox = ini.replace(day=28) + timedelta(days=4)
        fim = prox - timedelta(days=prox.day)
        return ini, fim
    return faixa_mes_anterior()

def conectar_banco():
    _carregar_fbclient()
    return fdb.connect(
        host="Server", database="QUESTOR", user="SYSDBA", password="masterkey",
        port=3050, charset="ISO8859_1",
    )

def _colunas_relacao(con, relacao: str):
    cur = con.cursor()
    cur.execute("""
        SELECT TRIM(RF.RDB$FIELD_NAME), F.RDB$FIELD_TYPE
        FROM RDB$RELATION_FIELDS RF
        JOIN RDB$FIELDS F ON F.RDB$FIELD_NAME = RF.RDB$FIELD_SOURCE
        WHERE TRIM(RF.RDB$RELATION_NAME) = ?
        ORDER BY 1
    """, (relacao.upper(),))
    return [(r[0], r[1]) for r in cur.fetchall()]

def _top_amostras(con, relacao: str, coluna: str, limite=500):
    cur = con.cursor()
    try:
        cur.execute(f"SELECT {coluna} FROM {relacao} ROWS {limite}")
        return [r[0] for r in cur.fetchall()]
    except: return []

def descobrir_coluna_cnpj_estab(con):
    cols = _colunas_relacao(con, "ESTAB")
    nomes = ["CNPJ","CGC_CPF","CGCCPF","CNPJCPF","CNPJ_CPF","CGC"]
    mapa = {c.upper(): c for c,_ in cols}
    for n in nomes:
        if n.upper() in mapa: return mapa[n.upper()]
    texto = [c for c,t in cols if t in (14,37)]
    best, hits = None, -1
    for c in texto:
        am = _top_amostras(con, "ESTAB", c, 800)
        h = sum(1 for v in am if len(so_digitos(v)) == 14)
        if h > hits: hits, best = h, c
    if best and hits > 0: return best
    raise RuntimeError("Não identifiquei a coluna de CNPJ em ESTAB.")

def construir_mapa_cnpj_empresa(con, col_cnpj: str):
    cur = con.cursor()
    cur.execute(f"SELECT CODIGOEMPRESA, {col_cnpj} FROM ESTAB")
    mapa = {}
    for cod_emp, raw in cur.fetchall():
        cnpj = so_digitos(raw)
        if not cnpj: continue
        if len(cnpj) < 14: cnpj = cnpj.zfill(14)
        mapa[cnpj] = cod_emp
    return mapa

QUERY = """
SELECT
    L.CODIGOEMPRESA, E.NOMEEMPRESA, L.NUMERONF, PS.NOMEPESSOA AS FORNECEDOR,
    SUM(LCF.VALORCONTABILIMPOSTO) AS VALOR, L.DATALCTOFIS AS DATA_LANCAMENTO,
    SUM(LCF.BASECALCULOIMPOSTO) AS BASE_ICMS, SUM(LCF.VALORIMPOSTO) AS ICMS
FROM LCTOFISENT L
JOIN LCTOFISENTCFOP LCF
  ON L.CODIGOEMPRESA = LCF.CODIGOEMPRESA AND L.CHAVELCTOFISENT = LCF.CHAVELCTOFISENT
JOIN EMPRESA E ON E.CODIGOEMPRESA = L.CODIGOEMPRESA
JOIN PESSOA  PS ON PS.CODIGOPESSOA = L.CODIGOPESSOA
WHERE L.CODIGOEMPRESA = ?
  AND L.DATALCTOFIS >= ? AND L.DATALCTOFIS < ?
  AND L.ESPECIENF IN ('NFE','NFCE','CTE')
  AND LCF.TIPOIMPOSTO = 1
  AND L.CANCELADA = 0
GROUP BY L.CODIGOEMPRESA, E.NOMEEMPRESA, L.NUMERONF, PS.NOMEPESSOA, L.DATALCTOFIS
ORDER BY L.DATALCTOFIS, L.NUMERONF
"""

QUERY_SAI = """
SELECT
    L.CODIGOEMPRESA, E.NOMEEMPRESA, L.NUMERONF, PS.NOMEPESSOA AS CLIENTE,
    SUM(LCF.VALORCONTABILIMPOSTO) AS VALOR, L.DATALCTOFIS AS DATA_LANCAMENTO,
    SUM(LCF.BASECALCULOIMPOSTO) AS BASE_ICMS, SUM(LCF.VALORIMPOSTO) AS ICMS
FROM LCTOFISSAI L
JOIN LCTOFISSAICFOP LCF
  ON L.CODIGOEMPRESA = LCF.CODIGOEMPRESA AND L.CHAVELCTOFISSAI = LCF.CHAVELCTOFISSAI
JOIN CFOP C ON C.CODIGOEMPRESA = LCF.CODIGOEMPRESA AND C.CODIGOESTAB = LCF.CODIGOESTAB AND C.CODIGOCFOP = LCF.CODIGOCFOP
JOIN EMPRESA E ON E.CODIGOEMPRESA = L.CODIGOEMPRESA
JOIN PESSOA  PS ON PS.CODIGOPESSOA = L.CODIGOPESSOA
WHERE L.CODIGOEMPRESA = ?
  AND L.DATALCTOFIS >= ? AND L.DATALCTOFIS < ?
  AND L.ESPECIENF IN ('NFE','NFCE')
  AND LCF.TIPOIMPOSTO = 1
  AND L.CANCELADA = 0
GROUP BY L.CODIGOEMPRESA, E.NOMEEMPRESA, L.NUMERONF, PS.NOM
PS.NOMEPESSOA, L.DATALCTOFIS
ORDER BY L.DATALCTOFIS, L.NUMERONF
"""

def consultar_bd_por_empresa(con, cod_empresa: int, ini: datetime, fim: datetime):
    cur = con.cursor()
    di = ini.strftime("%Y-%m-%d")
    df = (fim + timedelta(days=1)).strftime("%Y-%m-%d")
    cur.execute(QUERY, (cod_empresa, di, df))
    rows = cur.fetchall()
    cols = ["CodigoEmpresa","Empresa","NumeroNota","Fornecedor","Valor","Data_Lancamento","BaseICMS","ICMS"]
    return pd.DataFrame(rows, columns=cols)

def consultar_bd_saida_por_empresa(con, cod_empresa: int, ini: datetime, fim: datetime):
    cur = con.cursor()
    di = ini.strftime("%Y-%m-%d")
    df = (fim + timedelta(days=1)).strftime("%Y-%m-%d")
    cur.execute(QUERY_SAI, (cod_empresa, di, df))
    rows = cur.fetchall()
    cols = ["CodigoEmpresa","Empresa","NumeroNota","Cliente","Valor","Data_Lancamento","BaseICMS","ICMS"]
    return pd.DataFrame(rows, columns=cols)

def ui_log(widget, msg):
    widget.configure(state="normal")
    widget.insert("end", msg)
    widget.see("end")
    widget.configure(state="disabled")

def carregar_mapa_nomes_df(df: pd.DataFrame):
    if "CNPJ" not in df.columns or "Nome" not in df.columns:
        raise ValueError("Planilha CNPJ precisa ter colunas 'CNPJ' e 'Nome'.")
    mapa = {}
    for _, r in df.iterrows():
        k = limpar_nome_para_chave(r["Nome"])
        mapa[k] = {"CNPJ": r["CNPJ"], "Nome": r["Nome"]}
    return mapa

def conciliar_pasta_ui(dirpath, mapa_nomes, mapa_cnpj_emp, quadro_log2, tolerancia=TOLERANCIA_REAIS, con=None):
    nf_path = os.path.join(dirpath, NOME_NF)
    if not os.path.exists(nf_path): return None
    ui_log(quadro_log2, f"Pasta: {dirpath}\n")
    try:
        df_nf = pd.read_excel(nf_path)
    except Exception as e:
        ui_log(quadro_log2, f"  Erro NF: {e}\n")
        return {"pasta": dirpath, "total": 0, "lancados": 0, "nao_lancados": 0, "obs": "erro NF"}

    ini, fim = inferir_periodo(df_nf, dirpath)
    nome_key = limpar_nome_para_chave(os.path.basename(dirpath))
    row_emp = mapa_nomes.get(nome_key)
    if row_emp is None:
        ui_log(quadro_log2, "  Nome não mapeado em CNPJ.xlsx\n")
        return {"pasta": dirpath, "total": 0, "lancados": 0, "nao_lancados": 0, "obs": "nome não mapeado"}

    cnpj = re.sub(r"\D", "", str(row_emp["CNPJ"]))
    cod_emp = mapa_cnpj_emp.get(cnpj)
    if cod_emp is None:
        ui_log(quadro_log2, f"  CNPJ {cnpj} sem ESTAB\n")
        return {"pasta": dirpath, "total": 0, "lancados": 0, "nao_lancados": 0, "obs": "cnpj sem ESTAB"}

    try:
        df_db = consultar_bd_por_empresa(con, cod_emp, ini, fim)
    except Exception as e:
        ui_log(quadro_log2, f"  Erro BD: {e}\n")
        return {"pasta": dirpath, "total": 0, "lancados": 0, "nao_lancados": 0, "obs": "erro BD"}

    df_db["_v_db"] = df_db["Valor"].apply(parse_money)
    df_db["_ch"] = df_db["NumeroNota"].astype(str).str.strip() + "|" + df_db["_v_db"].round(2).astype(str)
    df_db = df_db.sort_values(["Data_Lancamento","NumeroNota"], na_position="last")
    df_db_dedup = df_db.drop_duplicates(subset=["_ch"], keep="first").drop(columns=["_ch"])

    nome_bd = f"consulta_bd_{ini.strftime('%Y%m%d')}_{fim.strftime('%Y%m%d')}.xlsx"
    salvo_bd = salvar_dataframe(df_db_dedup, os.path.join(dirpath, nome_bd), sheet_name="ConsultaBD")
    ui_log(quadro_log2, f"  Consulta BD: {salvo_bd}\n")

    if "ValorTotalNota" not in df_nf.columns:
        ui_log(quadro_log2, "  NF sem ValorTotalNota\n")
        return {"pasta": dirpath, "total": 0, "lancados": 0, "nao_lancados": 0, "obs": "NF sem ValorTotalNota"}

    df_nf["_v_nf"] = df_nf["ValorTotalNota"].apply(parse_money)
    db_cent = set(int(round(v * 100)) for v in df_db_dedup["_v_db"].dropna())
    tol_cent = int(round(abs(tolerancia) * 100))
    df_nf["Verificação"] = marcar_lancado_series(df_nf["_v_nf"], db_cent, tol_cent)

    df_nf = garantir_colunas(df_nf, COLUNAS_FINAIS)
    df_out = df_nf[COLUNAS_FINAIS].copy()
    salvo = salvar_dataframe(df_out, os.path.join(dirpath, "nfedestinatario_conciliado.xlsx"), sheet_name="Conciliado")

    total = len(df_out)
    lanc = int((df_out["Verificação"] == "Lançado").sum())
    nao = total - lanc
    ui_log(quadro_log2, f"  OK: {salvo} | Lançados: {lanc} | Não: {nao}\n")
    return {"pasta": dirpath, "total": total, "lancados": lanc, "nao_lancados": nao, "obs": ""}

def _prep_nf_emitente(df_nf):
    df = df_nf.copy()
    for c in ("NumeroDocumento","ValorTotalNota","ValorBaseCalculoICMS","ValorTotalICMS"):
        if c not in df.columns: df[c] = None
    df["NumeroDocumento"] = df["NumeroDocumento"].astype(str).str.strip()
    df["_v_nf_total"] = df["ValorTotalNota"].apply(parse_money)
    df["_v_nf_base"]  = df["ValorBaseCalculoICMS"].apply(parse_money)
    df["_v_nf_icms"]  = df["ValorTotalICMS"].apply(parse_money)
    return df

def _prep_bd_saida(df_db):
    df = df_db.copy()
    df["NumeroNota"] = df["NumeroNota"].astype(str).str.strip()
    df["_v_bd_total"] = df["Valor"].apply(parse_money)
    df["_v_bd_base"]  = df["BaseICMS"].apply(parse_money)
    df["_v_bd_icms"]  = df["ICMS"].apply(parse_money)
    return df

def _flag_ok(a, b, tol):
    import math
    if pd.isna(a) or pd.isna(b): return "N/D"
    return "OK" if math.isfinite(a) and math.isfinite(b) and abs(float(a)-float(b)) <= tol else "DIF"

def conciliar_emitente_arquivo(dirpath, filename, mapa_nomes, mapa_cnpj_emp, quadro_log2, tolerancia=TOLERANCIA_REAIS, con=None):
    nf_path = os.path.join(dirpath, filename)
    if not os.path.exists(nf_path): return None
    ui_log(quadro_log2, f"  Emitente: {filename}\n")
    try:
        df_nf = pd.read_excel(nf_path)
    except Exception as e:
        ui_log(quadro_log2, f"    Erro NF emitente: {e}\n")
        return {"pasta": dirpath, "arquivo": filename, "total": 0, "ok_total": 0, "ok_base": 0, "ok_icms": 0, "obs": "erro NF emitente"}

    ini, fim = inferir_periodo(df_nf, dirpath)
    nome_key = limpar_nome_para_chave(os.path.basename(dirpath))
    row_emp = mapa_nomes.get(nome_key)
    if row_emp is None:
        ui_log(quadro_log2, "    Nome não mapeado\n")
        return {"pasta": dirpath, "arquivo": filename, "total": 0, "ok_total": 0, "ok_base": 0, "ok_icms": 0, "obs": "nome não mapeado"}

    cnpj = re.sub(r"\D", "", str(row_emp["CNPJ"]))
    cod_emp = mapa_cnpj_emp.get(cnpj)
    if cod_emp is None:
        ui_log(quadro_log2, f"    CNPJ {cnpj} sem ESTAB\n")
        return {"pasta": dirpath, "arquivo": filename, "total": 0, "ok_total": 0, "ok_base": 0, "ok_icms": 0, "obs": "cnpj sem ESTAB"}

    try:
        df_db = consultar_bd_saida_por_empresa(con, cod_emp, ini, fim)
    except Exception as e:
        ui_log(quadro_log2, f"    Erro BD saída: {e}\n")
        return {"pasta": dirpath, "arquivo": filename, "total": 0, "ok_total": 0, "ok_base": 0, "ok_icms": 0, "obs": "erro BD SAI"}

    df_nf = _prep_nf_emitente(df_nf)
    df_db = _prep_bd_saida(df_db)
    df = df_nf.merge(df_db[["NumeroNota","_v_bd_total","_v_bd_base","_v_bd_icms"]], left_on="NumeroDocumento", right_on="NumeroNota", how="left")

    tol = float(abs(tolerancia))
    df["OK_ValorTotal"] = df.apply(lambda r: _flag_ok(r["_v_nf_total"], r["_v_bd_total"], tol), axis=1)
    df["OK_BaseICMS"]   = df.apply(lambda r: _flag_ok(r["_v_nf_base"],  r["_v_bd_base"],  tol), axis=1)
    df["OK_ICMS"]       = df.apply(lambda r: _flag_ok(r["_v_nf_icms"],  r["_v_bd_icms"],  tol), axis=1)

    cols = ["NumeroDocumento","DataEmissao","NomeEmitente","ValorTotalNota","ValorBaseCalculoICMS","ValorTotalICMS",
            "_v_bd_total","_v_bd_base","_v_bd_icms","OK_ValorTotal","OK_BaseICMS","OK_ICMS"]
    cols = [c for c in cols if c in df.columns]
    df_out = df[cols].copy()

    outp = os.path.join(dirpath, f"{os.path.splitext(filename)[0]}_emitente_conciliado.xlsx")
    salvo = salvar_dataframe(df_out, outp, sheet_name="EmitenteConciliado")

    ok_total = int((df_out.get("OK_ValorTotal","") == "OK").sum())
    ok_base  = int((df_out.get("OK_BaseICMS","")   == "OK").sum())
    ok_icms  = int((df_out.get("OK_ICMS","")       == "OK").sum())
    ui_log(quadro_log2, f"    OK: {salvo} | Valor: {ok_total} | Base: {ok_base} | ICMS: {ok_icms}\n")
    return {"pasta": dirpath, "arquivo": filename, "total": len(df_out), "ok_total": ok_total, "ok_base": ok_base, "ok_icms": ok_icms, "obs": ""}

def processar_arvore_ui_por_planilha(raiz, df_cnpj, quadro_log2, barra_progresso2, mes_pasta):
    base = os.path.join(raiz, mes_pasta) if mes_pasta else raiz
    if not os.path.isdir(base): base = raiz
    ui_log(quadro_log2, f"Iniciando conciliação em: {os.path.abspath(base)}\n")

    try:
        con = conectar_banco()
    except Exception as e:
        ui_log(quadro_log2, f"Erro conectando BD: {e}\n"); return

    try:
        try:
            col = descobrir_coluna_cnpj_estab(con)
        except Exception:
            col = "INSCRFEDERAL"
        mapa_cnpj_emp = construir_mapa_cnpj_empresa(con, col)
        mapa_nomes = carregar_mapa_nomes_df(df_cnpj)

        pastas = listar_pastas_por_planilha(base, df_cnpj, quadro_log_widget=quadro_log2)
        tot = max(len(pastas), 1)
        done, resumo = 0, []

        for d in pastas:
            stats = conciliar_pasta_ui(d, mapa_nomes, mapa_cnpj_emp, quadro_log2, TOLERANCIA_REAIS, con=con)
            if stats: resumo.append(stats)
            for arq in ("nfe.xlsx", "nfce.xlsx"):
                conciliar_emitente_arquivo(d, arq, mapa_nomes, mapa_cnpj_emp, quadro_log2, TOLERANCIA_REAIS, con=con)
            done += 1
            barra_progresso2.set(done / tot)
            janela.update_idletasks()

        if resumo:
            df_resumo = pd.DataFrame(resumo).sort_values(by="pasta")
            salvo = salvar_dataframe(df_resumo, os.path.join(base, "conciliacao_consolidada.xlsx"), sheet_name="Resumo")
            ui_log(quadro_log2, f"Resumo: {salvo}\n")
        ui_log(quadro_log2, "Conciliação finalizada.\n")
    except Exception as e:
        ui_log(quadro_log2, f"Erro: {e}\n")
    finally:
        try: con.close()
        except: pass

def rodar_conciliacao_thread():
    if DF_CNPJ_PARTE2 is None or not RAIZ_PASTAS2:
        messagebox.showwarning("Atenção", "Selecione a planilha CNPJ.xlsx na Parte II antes de iniciar.")
        return
    barra_progresso2.set(0)
    threading.Thread(target=processar_arvore_ui_por_planilha, args=(RAIZ_PASTAS2, DF_CNPJ_PARTE2, quadro_log2, barra_progresso2, MES_PASTA), daemon=True).start()

def selecionar_pasta_raiz():
    global RAIZ_PASTAS2
    pasta = filedialog.askdirectory()
    if pasta:
        RAIZ_PASTAS2 = pasta
        quadro_log2.configure(state="normal")
        quadro_log2.insert("end", f"Raiz: {pasta}\n")
        quadro_log2.configure(state="disabled")
        quadro_log2.see("end")

def abrir_resultado2():
    if not RAIZ_PASTAS2:
        messagebox.showinfo("Info", "Selecione a pasta raiz primeiro."); return
    try:
        os.startfile(RAIZ_PASTAS2)
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível abrir: {e}")

# --------------------------------------------------------------------------------------
# Interface
# --------------------------------------------------------------------------------------
df_cnpjs = pd.DataFrame()

def selecionar_arquivo():
    global df_cnpjs
    caminho = filedialog.askopenfilengit --version
ame(filetypes=[("Arquivos Excel", "*.xlsx")])
    if caminho:
        df_cnpjs = pd.read_excel(caminho)
        quadro_log.configure(state="normal")
        quadro_log.insert("end", f"Planilha: {os.path.basename(caminho)}\n")
        quadro_log.configure(state="disabled")
        quadro_log.see("end")

def rodar_automacao_thread():
    if df_cnpjs.empty:
        quadro_log.configure(state="normal"); quadro_log.insert("end", "Nenhuma planilha carregada.\n"); quadro_log.configure(state="disabled"); return
    threading.Thread(target=iniciar_automacao, args=(df_cnpjs, quadro_log, barra_progresso), daemon=True).start()

def abrir_resultado():
    pasta = primeiro_e_ultimo_dia_mes_anterior()[1].strftime("%m.%Y")
    os.startfile(pasta)

COR_FUNDO = "#f5f5f5"
COR_TOPO = "#243947"
COR_ABA = "#ffffff"
COR_TEXTO = "#243947"
COR_BOTAO_PRINCIPAL = "#315166"
COR_BOTAO_HOVER = "#1e2f44"
COR_ACENTO = "#ED9E2B"

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("green")

janela = ctk.CTk()
janela.geometry("1000x800")
janela.title("Consulta NF-e - M&H Soluções")
janela.configure(fg_color=COR_FUNDO)
try: janela.iconbitmap("logo.ico")
except: pass

topo = ctk.CTkFrame(janela, fg_color=COR_TOPO, height=60); topo.pack(fill="x")
ctk.CTkLabel(topo, text="Sistema de Consulta - SAT", font=("Segoe UI", 22, "bold"), text_color="white").pack(pady=10)

try:
    logo_img = Image.open("rebellato_esticada.png").resize((400, 100))
    logo_ctk = ctk.CTkImage(light_image=logo_img, dark_image=logo_img, size=(400, 100))
    ctk.CTkLabel(janela, image=logo_ctk, text="").pack(pady=(15, 5))
except: pass

abas = ctk.CTkTabview(janela, width=940, height=640,
                      segmented_button_selected_color=COR_BOTAO_PRINCIPAL,
                      segmented_button_fg_color="#d9d9d9",
                      text_color=COR_TEXTO)
abas.pack(pady=10)
aba1, aba2, aba3 = abas.add("Parte I"), abas.add("Parte II"), abas.add("Parte III")

# Parte I
frame1 = ctk.CTkFrame(aba1, fg_color=COR_ABA, corner_radius=20); frame1.pack(fill="both", expand=True, padx=30, pady=30)
fb1 = ctk.CTkFrame(frame1, fg_color="transparent"); fb1.pack(pady=(40, 20))
ctk.CTkButton(fb1, text="📄 Inserir CNPJ's (Excel)", font=("Segoe UI", 16), height=40, width=220,
              fg_color=COR_BOTAO_PRINCIPAL, hover_color=COR_BOTAO_HOVER, command=selecionar_arquivo).pack(side="left", padx=20)
ctk.CTkButton(fb1, text="🚀 Iniciar Automação", font=("Segoe UI", 16), height=40, width=220,
              fg_color=COR_BOTAO_PRINCIPAL, hover_color=COR_BOTAO_HOVER, command=rodar_automacao_thread).pack(side="left", padx=20)
barra_progresso = ctk.CTkProgressBar(frame1, width=500, progress_color=COR_ACENTO, corner_radius=8); barra_progresso.pack(pady=30); barra_progresso.set(0)
ctk.CTkButton(frame1, text="📂 Abrir Resultado", font=("Segoe UI", 16), height=40, width=220,
              fg_color=COR_BOTAO_PRINCIPAL, hover_color=COR_BOTAO_HOVER, command=abrir_resultado).pack(pady=(10, 20))
quadro_log = ctk.CTkTextbox(frame1, width=700, height=150, corner_radius=10, font=("Consolas", 12),
                            fg_color="#eeeeee", text_color=COR_TEXTO, scrollbar_button_color=COR_BOTAO_PRINCIPAL)
quadro_log.pack(pady=(0, 10)); quadro_log.insert("end", "Pronto para iniciar...\n"); quadro_log.configure(state="disabled")

# Parte II
frame2 = ctk.CTkFrame(aba2, fg_color=COR_ABA, corner_radius=20); frame2.pack(fill="both", expand=True, padx=30, pady=30)
fb2 = ctk.CTkFrame(frame2, fg_color="transparent"); fb2.pack(pady=(40, 20))
ctk.CTkButton(fb2, text="📄 Selecionar CNPJ.xlsx", font=("Segoe UI", 16), height=40, width=220,
              fg_color=COR_BOTAO_PRINCIPAL, hover_color=COR_BOTAO_HOVER, command=selecionar_planilha_cnpj_parte2).pack(side="left", padx=20)
ctk.CTkButton(fb2, text="🔍 Iniciar Conciliação", font=("Segoe UI", 16), height=40, width=220,
              fg_color=COR_BOTAO_PRINCIPAL, hover_color=COR_BOTAO_HOVER, command=rodar_conciliacao_thread).pack(side="left", padx=20)
ctk.CTkButton(frame2, text="📂 Conferir Resultado", font=("Segoe UI", 16), height=40, width=220,
              fg_color=COR_BOTAO_PRINCIPAL, hover_color=COR_BOTAO_HOVER, command=abrir_resultado2).pack(pady=(10, 20))
barra_progresso2 = ctk.CTkProgressBar(frame2, width=500, progress_color=COR_ACENTO, corner_radius=8); barra_progresso2.pack(pady=30); barra_progresso2.set(0)
quadro_log2 = ctk.CTkTextbox(frame2, width=700, height=150, corner_radius=10, font=("Consolas", 12),
                             fg_color="#eeeeee", text_color=COR_TEXTO, scrollbar_button_color=COR_BOTAO_PRINCIPAL)
quadro_log2.pack(pady=(0, 10)); quadro_log2.insert("end", "Pronto para conciliar...\n"); quadro_log2.configure(state="disabled")

# Parte III
frame3 = ctk.CTkFrame(aba3, fg_color=COR_ABA, corner_radius=20); frame3.pack(fill="both", expand=True, padx=30, pady=30)
fb3 = ctk.CTkFrame(frame3, fg_color="transparent"); fb3.pack(pady=(40, 20))
ctk.CTkButton(fb3, text="📄 Inserir CNPJ's (Excel)", font=("Segoe UI", 16), height=40, width=220,
              fg_color=COR_BOTAO_PRINCIPAL, hover_color=COR_BOTAO_HOVER, command=selecionar_arquivo).pack(side="left", padx=20)
ctk.CTkButton(fb3, text="🚀 Emitir CTE", font=("Segoe UI", 16), height=40, width=220,
              fg_color=COR_BOTAO_PRINCIPAL, hover_color=COR_BOTAO_HOVER, command=rodar_emissao_cte_thread).pack(side="left", padx=20)
barra_progresso3 = ctk.CTkProgressBar(frame3, width=500, progress_color=COR_ACENTO, corner_radius=8); barra_progresso3.pack(pady=30); barra_progresso3.set(0)
ctk.CTkButton(frame3, text="📂 Abrir Resultado", font=("Segoe UI", 16), height=40, width=220,
              fg_color=COR_BOTAO_PRINCIPAL, hover_color=COR_BOTAO_HOVER, command=abrir_resultado).pack(pady=(10, 20))
quadro_log3 = ctk.CTkTextbox(frame3, width=700, height=150, corner_radius=10, font=("Consolas", 12),
                             fg_color="#eeeeee", text_color=COR_TEXTO, scrollbar_button_color=COR_BOTAO_PRINCIPAL)
quadro_log3.pack(pady=(0, 10)); quadro_log3.insert("end", "Pronto para emissão de CTE...\n"); quadro_log3.configure(state="disabled")

ctk.CTkLabel(janela, text="© M&H Soluções", font=("Segoe UI", 12), text_color=COR_TEXTO).pack(pady=(10, 15))
janela.mainloop()
