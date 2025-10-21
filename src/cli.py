from __future__ import annotations
import argparse, os
from dotenv import load_dotenv
from .core.sat_client import rodar_consulta_planilha

load_dotenv()

def main():
    ap = argparse.ArgumentParser(description="Consulta SAT e exportação (modo CLI).")
    ap.add_argument("--cnpj-xlsx", required=True, help="Planilha com colunas CNPJ e Nome.")
    ap.add_argument("--saida", default="./saida", help="Diretório de saída.")
    ap.add_argument("--headless", action="store_true", help="Executar sem UI do navegador.")
    args = ap.parse_args()

    os.makedirs(args.saida, exist_ok=True)
    rodar_consulta_planilha(args.cnpj_xlsx, args.saida, headless=args.headless)

if __name__ == "__main__":
    main()
