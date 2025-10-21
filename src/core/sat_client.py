from __future__ import annotations
import os, time, random, logging
from datetime import datetime
from typing import Iterable
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, BrowserContext, Page
from twocaptcha import TwoCaptcha

from .utils import primeiro_e_ultimo_dia_mes_anterior, limpar_nome
from .sat_pages import SatLoginPage, SatConsultaNFePage

log = logging.getLogger(__name__)
load_dotenv()

def _sleep(a=1, b=3):  # delay humano
    time.sleep(random.uniform(a, b))

def _resolver_captcha(page: Page):
    page.get_by_role("img").screenshot(path="captcha.png")
    api_key = os.getenv('APIKEY_2CAPTCHA', '')
    if not api_key:
        raise RuntimeError("APIKEY_2CAPTCHA não definida")
    result = TwoCaptcha(api_key).normal("captcha.png")
    page.fill('#Body_Main_Main_sepConsultaNfpe_ctl21 > div > input', result['code'])
    page.locator('#Body_Main_Main_sepConsultaNfpe_btnBuscar > span').click()
    time.sleep(2)

def login_e_abrir_consulta(context: BrowserContext) -> Page:
    page = context.new_page()
    login = os.getenv("SAT_USER", "")
    senha = os.getenv("SAT_PASSWORD", "")
    if not (login and senha):
        raise RuntimeError("SAT_USER/SAT_PASSWORD ausentes no .env")

    login_page = SatLoginPage(page)
    login_page.goto()
    _sleep(2, 4)
    login_page.login(login, senha)
    _sleep(1, 2)
    new_page = login_page.abrir_app_por_busca("NFe / NFCe - Consulta", context)
    return new_page

def processar_cnpj(new_page: Page, cnpj: str, nome_empresa: str, pasta_destino: str, ref_data: datetime):
    ini_dt, fim_dt = primeiro_e_ultimo_dia_mes_anterior(ref_data)
    ini, fim = ini_dt.strftime("%d/%m/%Y"), fim_dt.strftime("%d/%m/%Y")
    pasta = os.path.join(pasta_destino, limpar_nome(nome_empresa))
    os.makedirs(pasta, exist_ok=True)

    nfe = SatConsultaNFePage(new_page)

    def tentar_exportar(destino: str):
        try:
            nfe.exportar(destino)
        except RuntimeError as e:
            if "captcha" in str(e).lower():
                _resolver_captcha(new_page)
                nfe.exportar(destino)
            else:
                raise

    # LIMPAR CAMPOS
    for sel in ('#Body_Main_Main_sepConsultaNfpe_ctl10_idnEmitente_MaskedField',
                '#Body_Main_Main_sepConsultaNfpe_ctl11_idnDestinatario_MaskedField'):
        try:
            new_page.fill(sel, '')
        except:
            pass

    # NF-e Emitente
    nfe.set_emitente(cnpj)
    nfe.set_datas(ini, fim)
    tentar_exportar(os.path.join(pasta, 'nfe.xlsx'))

    # NF-e Destinatário
    new_page.locator('#Body_Main_Main_sepConsultaNfpe_ctl10_idnEmitente_MaskedField').select_text()
    new_page.locator('#Body_Main_Main_sepConsultaNfpe_ctl10_idnEmitente_MaskedField').press('Delete')
    nfe.set_destinatario(cnpj)
    nfe.set_datas(ini, fim)
    tentar_exportar(os.path.join(pasta, 'nfedestinatario.xlsx'))
    new_page.locator('#Body_Main_Main_sepConsultaNfpe_ctl11_idnDestinatario_MaskedField').select_text()
    new_page.locator('#Body_Main_Main_sepConsultaNfpe_ctl11_idnDestinatario_MaskedField').press('Delete')

    # NFC-e
    nfe.escolher_aba("NF-e")
    nfe.escolher_aba("NFC-e")
    nfe.set_emitente(cnpj)
    nfe.set_datas(ini, fim)
    tentar_exportar(os.path.join(pasta, 'nfce.xlsx'))

    new_page.get_by_role("link", name="NFC-e").click()
    new_page.locator("#select2-drop").get_by_text("NF-e").click()

def rodar_consulta_planilha(path_xlsx: str, saida_dir: str, headless=False):
    import pandas as pd
    from .utils import mes_aa_mm_pasta

    df = pd.read_excel(path_xlsx)
    if not {"CNPJ", "Nome"}.issubset(df.columns):
        raise ValueError("Planilha precisa ter colunas 'CNPJ' e 'Nome'.")

    pasta_destino = os.path.join(saida_dir, mes_aa_mm_pasta())
    os.makedirs(pasta_destino, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless, channel='chrome' if not headless else None)
        context = browser.new_context()
        new_page = login_e_abrir_consulta(context)

        for row in df.itertuples():
            cnpj = str(getattr(row, "CNPJ")).zfill(14)
            nome = str(getattr(row, "Nome"))
            log.info("Processando %s (%s)", nome, cnpj)
            processar_cnpj(new_page, cnpj, nome, pasta_destino, datetime.now())
            time.sleep(1.0)

        new_page.close()
        context.close()
        browser.close()
