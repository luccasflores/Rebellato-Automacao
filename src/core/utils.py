from __future__ import annotations
import os, re, unicodedata, math, logging
from datetime import datetime, timedelta
import pandas as pd
from typing import Iterable, Tuple

log = logging.getLogger(__name__)

# ---------- Datas ----------
def primeiro_e_ultimo_dia_mes_anterior(ref: datetime | None = None) -> tuple[datetime, datetime]:
    hoje = ref or datetime.now()
    primeiro_atual = hoje.replace(day=1)
    ultimo_ant = primeiro_atual - timedelta(days=1)
    primeiro_ant = ultimo_ant.replace(day=1)
    return primeiro_ant, ultimo_ant

def faixa_mes_anterior(ref: datetime | None = None) -> tuple[datetime, datetime]:
    return primeiro_e_ultimo_dia_mes_anterior(ref)

def mes_aa_mm_pasta(ref: datetime | None = None) -> str:
    _, ultimo = primeiro_e_ultimo_dia_mes_anterior(ref)
    return ultimo.strftime("%m.%Y")

# ---------- Strings/Nomes ----------
def limpar_nome(s: str) -> str:
    s = unicodedata.normalize('NFKD', str(s)).encode('ASCII', 'ignore').decode('ASCII')
    s = re.sub(r'[^\w\s-]', '', s).strip()
    return re.sub(r'[-\s]+', '_', s)

def limpar_nome_para_chave(nome: str) -> str:
    nome = unicodedata.normalize('NFKD', str(nome)).encode('ASCII', 'ignore').decode('ASCII')
    nome = re.sub(r'[^\w\s.-]', '', nome).strip()
    nome = re.sub(r'[\s-]+', '_', nome)
    return nome.upper()

def so_digitos(x) -> str:
    return re.sub(r"\D", "", str(x or ""))

# ---------- Moeda ----------
def parse_money(v):
    import numpy as np
    if pd.isna(v): return np.nan
    s = str(v).strip().replace(" ", "")
    try:
        return float(s)
    except Exception:
        pass
    if "," in s and "." in s:
        s = s.replace(".", "") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
        s = s.replace(",", ".")
    elif "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        return float(s)
    except Exception:
        return np.nan

# ---------- Aux ----------
def marcar_lancado_series(series_val_nf: Iterable[float], set_db_centavos: set[int], tol_centavos=0) -> list[str]:
    out = []
    for v in series_val_nf:
        if pd.isna(v):
            out.append("Não Lançado");
            continue
        cents = int(round(float(v) * 100))
        if tol_centavos == 0:
            out.append("Lançado" if cents in set_db_centavos else "Não Lançado")
        else:
            ok = any((cents + d) in set_db_centavos for d in range(-tol_centavos, tol_centavos + 1))
            out.append("Lançado" if ok else "Não Lançado")
    return out

def escolher_engine_excel() -> str | None:
    try:
        import xlsxwriter  # noqa
        return "xlsxwriter"
    except Exception:
        try:
            import openpyxl  # noqa
            return "openpyxl"
        except Exception:
            return None

def salvar_dataframe(df: pd.DataFrame, out_path: str, sheet_name="Dados") -> str:
    import os
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    eng = escolher_engine_excel()
    if eng is None:
        csv_path = re.sub(r"\.xlsx$", "", out_path) + ".csv"
        df.to_csv(csv_path, index=False, sep=";", encoding="utf-8")
        return csv_path
    with pd.ExcelWriter(out_path, engine=eng) as w:
        df.to_excel(w, index=False, sheet_name=sheet_name)
    return out_path
