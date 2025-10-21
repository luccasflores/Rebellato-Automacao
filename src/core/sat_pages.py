from __future__ import annotations
from playwright.sync_api import Page, BrowserContext
import time

class SatLoginPage:
    URL = "https://sat.sef.sc.gov.br/tax.NET/Login.aspx?ReturnUrl=%2ftax.NET%2f"

    def __init__(self, page: Page):
        self.page = page

    def goto(self):
        self.page.goto(self.URL)
        self.page.wait_for_load_state("networkidle")

    def login(self, user: str, password: str):
        self.page.fill('#Body_pnlMain_tbxUsername', user)
        self.page.fill('#Body_pnlMain_tbxUserPassword', password)
        self.page.click('#Body_pnlMain_btnLogin > span')
        self.page.wait_for_load_state("networkidle")

    def abrir_app_por_busca(self, termo: str, context: BrowserContext) -> Page:
        self.page.wait_for_selector('#Body_Main_ctl09_ctl07_rptAppList_ctl03_1 > span')
        self.page.locator('#s2id_Body_ApplicationMasterHeader_ApplicationMasterSearchAppsInput_txtSearchApp_hid_single_txtSearchApp_value > a').click()
        self.page.fill('#select2-drop > div > input', termo)
        with context.expect_page() as new_page_info:
            self.page.locator('#select2-drop > ul > li > div').click()
        new_page = new_page_info.value
        new_page.wait_for_load_state("networkidle")
        time.sleep(1)
        return new_page


class SatConsultaNFePage:
    def __init__(self, page: Page):
        self.page = page

    def escolher_aba(self, label: str):
        self.page.get_by_role("link", name=label).click()
        self.page.locator("#select2-drop").get_by_text(label).click()

    def _input(self, sel: str, val: str):
        el = self.page.locator(sel)
        el.click()
        self.page.keyboard.press("Control+A")
        self.page.keyboard.press("Delete")
        el.fill(val)

    def set_datas(self, ini: str, fim: str):
        self._input("#Body_Main_Main_sepConsultaNfpe_datDataInicial", ini)
        self._input("#Body_Main_Main_sepConsultaNfpe_datDataFinal", fim)

    def set_emitente(self, cnpj: str):
        self._input('#Body_Main_Main_sepConsultaNfpe_ctl10_idnEmitente_MaskedField', cnpj)

    def set_destinatario(self, cnpj: str):
        self._input('#Body_Main_Main_sepConsultaNfpe_ctl11_idnDestinatario_MaskedField', cnpj)

    def _captcha_visivel(self) -> bool:
        try:
            return self.page.locator('#Body_Main_Main_sepConsultaNfpe_ctl17 > img').is_visible()
        except:
            return False

    def acionar_busca(self):
        self.page.locator('#Body_Main_Main_sepConsultaNfpe_btnBuscar > span').click()

    def exportar(self, destino: str):
        if self._captcha_visivel():
            raise RuntimeError("captcha")  # o caller trata e resolve
        with self.page.expect_download(timeout=60000) as dlinfo:
            self.page.locator('#Body_Main_Main_sepConsultaNfpe_btnExportar > span').click()
        dlinfo.value.save_as(destino)
