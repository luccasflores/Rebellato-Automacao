from __future__ import annotations
import os, threading, logging
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image
from dotenv import load_dotenv

from ..core.utils import mes_aa_mm_pasta, primeiro_e_ultimo_dia_mes_anterior
from ..core.sat_client import rodar_consulta_planilha
from ..core.reconciliation import conciliar_pasta

load_dotenv()
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(name)s - %(message)s",
    handlers=[logging.FileHandler("app.log"), logging.StreamHandler()]
)
log = logging.getLogger("UI")

# ------------------ ESTADO ------------------
class AppState:
    def __init__(self):
        self.df_cnpjs = pd.DataFrame()
        self.planilha_cnpj_parte2_path = ""
        self.raiz_pastas2 = ""
        self.mes_pasta = ""

state = AppState()

# ------------------ AÇÕES ------------------
def selecionar_arquivo_parte1():
    caminho = filedialog.askopenfilename(title="Selecione a planilha (CNPJ.xlsx)",
                                         filetypes=[("Arquivos Excel", "*.xlsx")])
    if caminho:
        state.df_cnpjs = pd.read_excel(caminho)
        quadro_log.configure(state="normal")
        quadro_log.insert("end", f"Planilha: {os.path.basename(caminho)}\n")
        quadro_log.configure(state="disabled")
        quadro_log.see("end")

def iniciar_automacao_thread():
    if state.df_cnpjs.empty:
        quadro_log.configure(state="normal")
        quadro_log.insert("end", "Nenhuma planilha carregada.\n")
        quadro_log.configure(state="disabled")
        return
    def run():
        try:
            saida = mes_aa_mm_pasta()
            rodar_consulta_planilha(path_xlsx=planilha_temp(), saida_dir=saida, headless=False)
            quadro_log.configure(state="normal")
            quadro_log.insert("end", "Automação finalizada.\n")
            quadro_log.configure(state="disabled")
        except Exception as e:
            messagebox.showerror("Erro", str(e))
    threading.Thread(target=run, daemon=True).start()

def planilha_temp() -> str:
    """Salva df_cnpjs do estado numa planilha temporária para passar ao core (evita acoplamento GUI)."""
    path = os.path.join(os.getcwd(), "_cnpjs_tmp.xlsx")
    state.df_cnpjs.to_excel(path, index=False)
    return path

def abrir_resultado_parte1():
    pasta = primeiro_e_ultimo_dia_mes_anterior()[1].strftime("%m.%Y")
    try:
        os.startfile(pasta)
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível abrir: {e}")

# ----- Parte II (Conciliação) -----
def selecionar_planilha_cnpj_parte2():
    caminho = filedialog.askopenfilename(title="Selecione a planilha CNPJ.xlsx",
                                         filetypes=[("Arquivos Excel", "*.xlsx")])
    if not caminho:
        return
    state.planilha_cnpj_parte2_path = caminho
    state.mes_pasta = primeiro_e_ultimo_dia_mes_anterior()[1].strftime("%m.%Y")
    state.raiz_pastas2 = os.path.dirname(caminho)

    quadro_log2.configure(state="normal")
    quadro_log2.insert("end", f"Planilha CNPJ: {os.path.basename(caminho)}\n")
    quadro_log2.insert("end", f"Raiz base: {state.raiz_pastas2}\n")
    quadro_log2.insert("end", f"Pasta alvo: {state.mes_pasta}\n")
    quadro_log2.configure(state="disabled")
    quadro_log2.see("end")

def selecionar_pasta_raiz():
    pasta = filedialog.askdirectory()
    if pasta:
        state.raiz_pastas2 = pasta
        quadro_log2.configure(state="normal")
        quadro_log2.insert("end", f"Raiz: {pasta}\n")
        quadro_log2.configure(state="disabled")
        quadro_log2.see("end")

def rodar_conciliacao_thread():
    if not state.planilha_cnpj_parte2_path or not state.raiz_pastas2:
        messagebox.showwarning("Atenção", "Selecione a planilha CNPJ e a pasta raiz.")
        return
    df_cnpj = pd.read_excel(state.planilha_cnpj_parte2_path)
    base = os.path.join(state.raiz_pastas2, state.mes_pasta) if state.mes_pasta else state.raiz_pastas2
    if not os.path.isdir(base):
        base = state.raiz_pastas2

    def run():
        try:
            # lista pastas do mês e processa
            pastas = [os.path.join(base, d) for d in os.listdir(base) if os.path.isdir(os.path.join(base, d))]
            total = max(1, len(pastas))
            done = 0
            for d in pastas:
                # só processa se tiver nfedestinatario.xlsx dentro
                if not os.path.exists(os.path.join(d, "nfedestinatario.xlsx")):
                    continue
                salvo = conciliar_pasta(d, df_cnpj)
                quadro_log2.configure(state="normal")
                quadro_log2.insert("end", f"{'OK' if salvo else 'PULAR'}: {d}\n")
                quadro_log2.configure(state="disabled")
                done += 1
                barra_progresso2.set(done / total)
                janela.update_idletasks()
            quadro_log2.configure(state="normal")
            quadro_log2.insert("end", "Conciliação finalizada.\n")
            quadro_log2.configure(state="disabled")
        except Exception as e:
            messagebox.showerror("Erro", str(e))
    threading.Thread(target=run, daemon=True).start()

def abrir_resultado2():
    if not state.raiz_pastas2:
        messagebox.showinfo("Info", "Selecione a pasta raiz primeiro.")
        return
    try:
        os.startfile(state.raiz_pastas2)
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível abrir: {e}")

# ------------------ UI ------------------
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
try:
    janela.iconbitmap("logo.ico")
except Exception:
    pass

topo = ctk.CTkFrame(janela, fg_color=COR_TOPO, height=60); topo.pack(fill="x")
ctk.CTkLabel(topo, text="Sistema de Consulta - SAT", font=("Segoe UI", 22, "bold"), text_color="white").pack(pady=10)

try:
    logo_img = Image.open("rebellato_esticada.png").resize((400, 100))
    logo_ctk = ctk.CTkImage(light_image=logo_img, dark_image=logo_img, size=(400, 100))
    ctk.CTkLabel(janela, image=logo_ctk, text="").pack(pady=(15, 5))
except Exception:
    pass

abas = ctk.CTkTabview(janela, width=940, height=640,
                      segmented_button_selected_color=COR_BOTAO_PRINCIPAL,
                      segmented_button_fg_color="#d9d9d9",
                      text_color=COR_TEXTO)
abas.pack(pady=10)
aba1, aba2 = abas.add("Parte I"), abas.add("Parte II")  # (Parte III CTe pode ser adicionada depois)

# Parte I
frame1 = ctk.CTkFrame(aba1, fg_color=COR_ABA, corner_radius=20); frame1.pack(fill="both", expand=True, padx=30, pady=30)
fb1 = ctk.CTkFrame(frame1, fg_color="transparent"); fb1.pack(pady=(40, 20))
ctk.CTkButton(fb1, text="📄 Inserir CNPJ's (Excel)", font=("Segoe UI", 16), height=40, width=220,
              fg_color=COR_BOTAO_PRINCIPAL, hover_color=COR_BOTAO_HOVER, command=selecionar_arquivo_parte1).pack(side="left", padx=20)
ctk.CTkButton(fb1, text="🚀 Iniciar Automação", font=("Segoe UI", 16), height=40, width=220,
              fg_color=COR_BOTAO_PRINCIPAL, hover_color=COR_BOTAO_HOVER, command=iniciar_automacao_thread).pack(side="left", padx=20)
barra_progresso = ctk.CTkProgressBar(frame1, width=500, progress_color=COR_ACENTO, corner_radius=8); barra_progresso.pack(pady=30); barra_progresso.set(0)
ctk.CTkButton(frame1, text="📂 Abrir Resultado", font=("Segoe UI", 16), height=40, width=220,
              fg_color=COR_BOTAO_PRINCIPAL, hover_color=COR_BOTAO_HOVER, command=abrir_resultado_parte1).pack(pady=(10, 20))
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

ctk.CTkLabel(janela, text="© M&H Soluções", font=("Segoe UI", 12), text_color=COR_TEXTO).pack(pady=(10, 15))

if __name__ == "__main__":
    janela.mainloop()
