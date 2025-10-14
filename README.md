# Rebellato – Automação Fiscal (SAT SC, NF-e/NFC-e e CTe) + Conciliação

App desktop em **Python** (Tkinter/CustomTkinter) que:
- **Baixa NF-e/NFC-e** (por emitente e destinatário) do SAT/SEF-SC via **Playwright**  
- **Baixa CTe** (por tomador) do SAT/SEF-SC  
- **Concilia** as planilhas exportadas com o **ERP Questor (Firebird)**  
- Gera saídas em **Excel** já filtradas e um **resumo consolidado**

> Foco: reduzir trabalho manual do fiscal, padronizar pastas por mês (`MM.AAAA`) e entregar **auditoria rápida** de lançamentos.

---

## ✨ Destaques
- UI em **CustomTkinter**, com barras de progresso e logs
- **Resiliência** a captcha (2Captcha)
- Normalização de nomes/pastas para localizar empresas de uma planilha
- Conciliação por **valor/base/ICMS** com tolerância
- Exporta resultados prontos para conferência

---

## 🛠️ Stack
- Python 3.11+
- **Playwright** (Chromium/Chrome)
- **Pandas**, **OpenPyXL/XlsxWriter**, **NumPy**
- **fdb** (Firebird Client)
- **TwoCaptcha**
- CustomTkinter

---

## 📸 Telas do Sistema

Interface moderna e funcional desenvolvida em **CustomTkinter**, com design adaptado à identidade visual da empresa Rebellato Contabilidade.

![Sistema de Consulta SAT](docs/sistema-consulta-sat.png)

## 🚀 Como rodar

### 1) Pré-requisitos
- Windows 10/11, Python 3.11+ e Git  
- Google Chrome instalado  
- **Firebird Client** (mesma arquitetura do Python). Deixe o `fbclient.dll` instalável em:
  - `C:\Program Files\Firebird\Firebird_5_0\bin\fbclient.dll` (ou 4.0)  
  - ou defina `FIREBIRD_CLIENT_PATH` com o caminho completo

### 2) Clonar e instalar
```powershell
git clone https://github.com/luccasflores/Rebellato-Automacao.git
cd Rebellato-Automacao

python -m venv .venv
.venv\Scripts\activate

pip install -r requirements.txt
python -m playwright install --with-deps
