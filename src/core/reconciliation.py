from __future__ import annotations
import os, re, logging
import pandas as pd
from datetime import datetime, timedelta
from .utils import (
    parse_money, salvar_dataframe, limpar_nome_para_chave, so_digitos,
    marcar_lancado_series, faixa_mes_anterior
)
from .firebird import conectar

log = logging.getLogger(__name__)

COLUNAS_FINAIS = [
    "NomeEmitente","TipoDeOperacaoEntradaOuSaida","Situacao","ChaveAcesso",
    "DataEmissao","NumeroDocumento","ValorTotalNota","ValorTotalISS",
    "ValorTotalICMS","ValorBaseCalculoICMS","Verificação"
]

QUERY_ENT = """
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
GROUP BY L.CODIGOEMPRESA, E.NOMEEMPRESA, L.NUMERONF, PS.NOMEPESSOA, L.DATALCTOFIS
ORDER BY L.DATALCTOFIS, L.NUMERONF
"""

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
    except Exception:
        return []

def descobrir_coluna_cnpj_estab(con) -> str:
    cols = _colunas_relacao(con, "ESTAB")
    nomes = ["CNPJ","CGC_CPF","CGCCPF","CNPJCPF","CNPJ_CPF","CGC","INSCRFEDERAL"]
    mapa = {c.upper(): c for c,_ in cols}
    for n in nomes:
        if n.upper() in mapa:
            return mapa[n.upper()]
    texto = [c for c,_t in cols]
    best, hits = None, -1
    for c in texto:
        am = _top_amostras(con, "ESTAB", c, 800)
        h = sum(1 for v in am if len(so_digitos(v)) == 14)
        if h > hits:
            hits, best = h, c
    if not best:
        raise RuntimeError("Não identifiquei coluna de CNPJ em ESTAB.")
    return best

def construir_mapa_cnpj_empresa(con, col_cnpj: str):
    cur = con.cursor()
    cur.execute(f"SELECT CODIGOEMPRESA, {col_cnpj} FROM ESTAB")
    mapa = {}
    for cod_emp, raw in cur.fetchall():
        cnpj = so_digitos(raw)
        if not cnpj:
            continue
        if len(cnpj) < 14:
            cnpj = cnpj.zfill(14)
        mapa[cnpj] = cod_emp
    return mapa

def consultar_bd_entrada(con, cod_empresa: int, ini: datetime, fim: datetime) -> pd.DataFrame:
    cur = con.cursor()
    di = ini.strftime("%Y-%m-%d")
    df = (fim + timedelta(days=1)).strftime("%Y-%m-%d")
    cur.execute(QUERY_ENT, (cod_empresa, di, df))
    rows = cur.fetchall()
    cols = ["CodigoEmpresa","Empresa","NumeroNota","Fornecedor","Valor","Data_Lancamento","BaseICMS","ICMS"]
    return pd.DataFrame(rows, columns=cols)

def consultar_bd_saida(con, cod_empresa: int, ini: datetime, fim: datetime) -> pd.DataFrame:
    cur = con.cursor()
    di = ini.strftime("%Y-%m-%d")
    df = (fim + timedelta(days=1)).strftime("%Y-%m-%d")
    cur.execute(QUERY_SAI, (cod_empresa, di, df))
    rows = cur.fetchall()
    cols = ["CodigoEmpresa","Empresa","NumeroNota","Cliente","Valor","Data_Lancamento","BaseICMS","ICMS"]
    return pd.DataFrame(rows, columns=cols)

def conciliar_pasta(dirpath: str, df_cnpj: pd.DataFrame, tolerancia_reais: float = 0.0):
    nf_path = os.path.join(dirpath, "nfedestinatario.xlsx")
    if not os.path.exists(nf_path):
        return None

    try:
        df_nf = pd.read_excel(nf_path)
    except Exception as e:
        log.error("Erro NF: %s", e)
        return None

    # período pelos dados ou nome da pasta
    ini, fim = _inferir_periodo(df_nf, dirpath)

    # mapear nome → CNPJ
    if not {"CNPJ","Nome"}.issubset(df_cnpj.columns):
        raise ValueError("CNPJ.xlsx precisa ter colunas 'CNPJ' e 'Nome'.")

    mapa_nomes = {limpar_nome_para_chave(r["Nome"]): r for _, r in df_cnpj.iterrows()}
    nome_key = limpar_nome_para_chave(os.path.basename(dirpath))
    row_emp = mapa_nomes.get(nome_key)
    if row_emp is None:
        log.warning("Nome não mapeado em %s", dirpath)
        return None

    cnpj = so_digitos(row_emp["CNPJ"])

    con = conectar()
    try:
        col = descobrir_coluna_cnpj_estab(con)
    except Exception:
        col = "INSCRFEDERAL"

    mapa_cnpj_emp = construir_mapa_cnpj_empresa(con, col)
    cod_emp = mapa_cnpj_emp.get(cnpj)
    if cod_emp is None:
        log.warning("CNPJ %s sem ESTAB", cnpj)
        return None

    df_db = consultar_bd_entrada(con, cod_emp, ini, fim)
    con.close()

    df_db["_v_db"] = df_db["Valor"].apply(parse_money)
    db_cent = set(int(round(v * 100)) for v in df_db["_v_db"].dropna())
    tol_cent = int(round(abs(tolerancia_reais) * 100))

    if "ValorTotalNota" not in df_nf.columns:
        log.warning("NF sem ValorTotalNota em %s", nf_path)
        return None

    df_nf["_v_nf"] = df_nf["ValorTotalNota"].apply(parse_money)
    df_nf["Verificação"] = marcar_lancado_series(df_nf["_v_nf"], db_cent, tol_cent)

    for c in COLUNAS_FINAIS:
        if c not in df_nf.columns and c != "Verificação":
            df_nf[c] = None

    df_out = df_nf[COLUNAS_FINAIS].copy()
    salvo = salvar_dataframe(df_out, os.path.join(dirpath, "nfedestinatario_conciliado.xlsx"), sheet_name="Conciliado")
    log.info("Conciliado: %s", salvo)
    return salvo

def _inferir_periodo(df_nf: pd.DataFrame, dirpath: str):
    for c in ["DataEmissao","Data_Emissao","Data","Emissao"]:
        if c in df_nf.columns:
            try:
                datas = pd.to_datetime(df_nf[c], dayfirst=True, errors="coerce").dropna()
                if len(datas):
                    return datas.min().normalize(), datas.max().normalize()
            except Exception:
                pass
    m = re.search(r"(\d{2})[._-](\d{4})", os.path.basename(dirpath))
    if m:
        mes, ano = int(m.group(1)), int(m.group(2))
        ini = datetime(ano, mes, 1)
        prox = ini.replace(day=28) + timedelta(days=4)
        fim = prox - timedelta(days=prox.day)
        return ini, fim
    return faixa_mes_anterior()
