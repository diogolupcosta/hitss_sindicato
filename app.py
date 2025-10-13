import io
import re
import unicodedata
from typing import List, Optional, Dict, Any

import pandas as pd
import pdfplumber
import streamlit as st


# =========================
# Config da página
# =========================
# =========================
# Config da página
# =========================
st.set_page_config(page_title="Oposição - Extrator de Tabelas", layout="centered", initial_sidebar_state="expanded")

# --- LOGO acima do cabeçalho ---
from pathlib import Path
LOGO_PATH = Path(__file__).with_name("Customer-Logos-09.png")  # ajuste se estiver em outra pasta

c1, c2, c3 = st.columns([1, 2, 1])  # centraliza
with c2:
    st.image(str(LOGO_PATH), width=260)  # use_container_width=True se quiser ocupar toda a coluna

# Cabeçalho
st.title("Extrator de Tabelas do Sindicato | HITSS")
st.caption(
    "Carregue um PDF. O app extrai apenas **tabelas**, ignora cabeçalhos/rodapés institucionais, "
    "e mantém as colunas **nome**, **cpf** e **data de entrega**. Se o cabeçalho não vier claro, "
    "usamos um fallback por **regex** para identificar CPF e Data."
)


# =========================
# Constantes e utilidades
# =========================

# Textos a ignorar (rodapés/cabeçalhos):
IGNORES_RAW = [
    """SINDICATO     DOS     EMPREGADOS     EM     EMPRESAS     DE     PROCESSAMENTO     DE     DADOS,     DE SERVIÇOS    DE    COMPUTAÇÃO,    DE    INFORMÁTICA    E    TECNOLOGIA    DA    INFORMAÇÃO    E    DOS TRABALHADORES       EM       PROCESSAMENTO       DE       DADOS,       SERVIÇOS       DE       COMPUTAÇÃO, INFORMÁTICA E TECNOLOGIA DA INFORMAÇÃO DO ESTADO DE SÃO PAULO.""",
    """HITSS DO BRASIL SERVICOS TECNOLOGICOS LTDA.FUNDADO EM 14/08/1984  -  CNPJ 55.537.666/0001-75  -   Av Angelica 35  -  Santa Cecilia  -  Sao Paulo  -  SP  -  CEP 01227-000www.sindpd.org.br  -  sindpd@sindpd.org.br  -  Central de Atendimento: ( 11 ) 3823-5600""",
]

CPF_RE = re.compile(r"\b\d{3}\.\d{3}\.\d{3}-\d{2}\b")
DATE_RE = re.compile(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b")  # dd/mm/aaaa (ou aa)
INT_COL_RE = re.compile(r"^\s*\d+\s*$")


def norm_txt(s: Optional[str]) -> str:
    """Normaliza para comparação: sem acentos, minúscula, espaços colapsados."""
    if s is None:
        return ""
    if not isinstance(s, str):
        s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.lower()
    s = re.sub(r"\s+", " ", s).strip()
    return s


IGNORES = [norm_txt(x) for x in IGNORES_RAW]


def has_ignored_text(text: str) -> bool:
    t = norm_txt(text)
    return any(ign in t for ign in IGNORES)


def clean_cells(df: pd.DataFrame) -> pd.DataFrame:
    """Tira quebras de linha e espaços redundantes em todas as células."""
    df = df.applymap(lambda v: v.replace("\n", " ").strip() if isinstance(v, str) else v)
    return df


def drop_empty_rows_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Remove linhas/colunas totalmente vazias."""
    df = df.dropna(axis=1, how="all")
    df = df.dropna(axis=0, how="all")
    # também limpa colunas/linhas feitas só de strings vazias
    df = df[[c for c in df.columns if not all((str(x).strip() == "" or pd.isna(x)) for x in df[c])]]
    df = df.loc[~df.apply(lambda r: all((str(x).strip() == "" or pd.isna(x)) for x in r), axis=1)]
    return df.reset_index(drop=True)


def is_mostly_int_column(series_or_df: pd.Series | pd.DataFrame, min_ratio: float = 0.8) -> bool:
    """Retorna True se a coluna tem >= min_ratio de inteiros (ou vazios). Aceita Series ou DataFrame.

    - Se vier DataFrame (colunas duplicadas), reduz a uma Series pegando o primeiro valor não vazio por linha.
    """
    if isinstance(series_or_df, pd.DataFrame):
        # pega o primeiro valor "não vazio" por linha
        series = series_or_df.apply(
            lambda r: next((x for x in r if pd.notna(x) and str(x).strip() != ""), None), axis=1
        )
    else:
        series = series_or_df

    valid = series.dropna().astype(str).str.strip()
    if valid.empty:
        return False
    hits = valid.str.match(r"^\d+$").sum()
    ratio = hits / len(valid)
    return ratio >= min_ratio



# =========================
# Extração de tabelas (pdfplumber)
# =========================
def extract_tables_from_pdf(pdf_bytes: bytes) -> List[pd.DataFrame]:
    """Extrai tabelas usando lattice e stream. Retorna lista de DataFrames brutos."""
    dfs: List[pd.DataFrame] = []

    def safe_extract(page, settings: Dict[str, Any]):
        # Remove argumentos desconhecidos para compatibilidade entre versões
        while True:
            try:
                return page.extract_tables(settings) or []
            except TypeError as e:
                msg = str(e)
                if "unexpected keyword argument" in msg:
                    bad = msg.split("'")[1]
                    settings = {k: v for k, v in settings.items() if k != bad}
                    continue
                raise

    lattice = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 3,
        "join_tolerance": 3,
        "min_words_vertical": 1,
        "min_words_horizontal": 1,
        "intersection_tolerance": 3,
    }
    stream = {
        "vertical_strategy": "text",
        "horizontal_strategy": "text",
        "intersection_tolerance": 5,
        "text_x_tolerance": 2,
        "text_y_tolerance": 2,
        "snap_tolerance": 3,
        "min_words_vertical": 1,
        "min_words_horizontal": 1,
        "join_tolerance": 3,
    }

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            for settings in (lattice, stream):
                for tbl in safe_extract(page, dict(settings)):
                    if not tbl:
                        continue
                    df = pd.DataFrame(tbl)
                    if df.shape[0] and df.shape[1]:
                        dfs.append(df)

    return dfs


def dedupe_columns(cols):
    seen = {}
    out = []
    for c in [str(x) for x in cols]:
        base = c
        if base not in seen:
            seen[base] = 0
            out.append(base)
        else:
            seen[base] += 1
            out.append(f"{base}.{seen[base]}")
    return out


# =========================
# Mapeamento por cabeçalho
# =========================
def map_by_header(df: pd.DataFrame) -> Optional[pd.DataFrame]:
    """
    Usa a 1ª linha como cabeçalho. Procura colunas: nome / cpf / data (data entrega, data de entrega, data).
    Remove colunas-resíduo com inteiros.
    """
    if df.empty:
        return None

    df = clean_cells(df)
    df = drop_empty_rows_cols(df)
    if df.empty:
        return None

    header = [str(x) for x in df.iloc[0].tolist()]
    body = df.iloc[1:].reset_index(drop=True)
    # deduplica nomes repetidos do cabeçalho (CPF, CPF.1, ...)
    body.columns = dedupe_columns(header)


    # Drop colunas-resíduo (quase só inteiros)
    to_drop = []
    for c in body.columns:
        nc = norm_txt(c)
        if any(key in nc for key in ["cpf", "nome", "data"]):
            continue
        if is_mostly_int_column(body[c]):
            to_drop.append(c)
    body = body.drop(columns=to_drop) if to_drop else body


    norm_cols = [norm_txt(c) for c in body.columns]
    colmap: Dict[str, str] = {}

    # cpf
    for i, c in enumerate(norm_cols):
        if "cpf" in c:
            colmap["cpf"] = body.columns[i]
            break

    # nome
    for i, c in enumerate(norm_cols):
        if "nome" in c or "empregado" in c or "trabalhador" in c:
            colmap["nome"] = body.columns[i]
            break

    # data (aceita "data entrega", "data de entrega", somente "data")
    for i, c in enumerate(norm_cols):
        if ("data" in c and "entreg" in c) or (c == "data") or ("data" in c and "opos" in c):
            colmap["data de entrega"] = body.columns[i]
            break

    # fallback data: qualquer coluna com muitos padrões dd/mm/aaaa
    if "data de entrega" not in colmap:
        for col in body.columns:
            s = body[col].dropna().astype(str)
            hits = s.str.contains(DATE_RE).sum()
            if len(s) > 0 and hits / len(s) >= 0.5:
                colmap["data de entrega"] = col
                break

    required = ["nome", "cpf", "data de entrega"]
    if not all(k in colmap for k in required):
        return None

    out = body[[colmap["nome"], colmap["cpf"], colmap["data de entrega"]]].copy()
    out.columns = required

    # remove linhas com textos ignorados
    mask = out.apply(lambda r: any(has_ignored_text(str(v)) for v in r), axis=1)
    out = out.loc[~mask]

    # remove linhas totalmente vazias e duplicadas
    out = out.replace(r"^\s*$", pd.NA, regex=True).dropna(how="all").drop_duplicates()
    out = out.reset_index(drop=True)

    return out if not out.empty else None


# =========================
# Fallback por regex (linha a linha)
# =========================
def parse_by_regex(df: pd.DataFrame) -> Optional[pd.DataFrame]:
    """
    Junta as células de cada linha e tenta extrair: NOME (antes do CPF), CPF e DATA (primeira após CPF).
    Ignora linhas de cabeçalho e textos institucionais.
    Também ignora lixo numérico após a data (ex.: ... 10/01/2025 1).
    """
    if df.empty:
        return None

    df = clean_cells(df)
    df = drop_empty_rows_cols(df)
    if df.empty:
        return None

    rows = []
    for _, row in df.iterrows():
        cells = [str(x) for x in row.tolist() if str(x).strip() != ""]
        if not cells:
            continue
        line = " ".join(cells)

        # ignora cabeçalhos e textos institucionais
        nline = norm_txt(line)
        if ("nome" in nline and "cpf" in nline and "data" in nline) or has_ignored_text(line):
            continue

        m_cpf = CPF_RE.search(line)
        m_date = DATE_RE.search(line)

        if not m_cpf or not m_date:
            # às vezes a data pode estar em outra célula: tenta no conjunto de células
            joined = " ".join(cells)
            m_cpf = m_cpf or CPF_RE.search(joined)
            m_date = m_date or DATE_RE.search(joined)

        if m_cpf and m_date:
            cpf = m_cpf.group(0)
            date_pos = m_date.start()
            cpf_pos = m_cpf.start()
            # nome: conteúdo antes do CPF
            nome = line[:cpf_pos].strip()
            # data: primeira data encontrada após (ou antes) do cpf
            data = m_date.group(0)
            # validações simples
            if nome and CPF_RE.fullmatch(cpf) and DATE_RE.fullmatch(data):
                # filtra linhas com textos ignorados
                if not has_ignored_text(nome):
                    rows.append({"nome": nome, "cpf": cpf, "data de entrega": data})

    if not rows:
        return None

    out = pd.DataFrame(rows).drop_duplicates().reset_index(drop=True)
    return out if not out.empty else None


# =========================
# Pipeline: extrair -> normalizar
# =========================
def process_pdf(pdf_bytes: bytes) -> List[pd.DataFrame]:
    """Extrai tabelas e tenta normalizar: primeiro por cabeçalho; se falhar, usa regex."""
    raw_tables = extract_tables_from_pdf(pdf_bytes)
    usable: List[pd.DataFrame] = []

    for df in raw_tables:
        # 1) tenta por cabeçalho
        mapped = map_by_header(df)
        if mapped is not None and not mapped.empty:
            usable.append(mapped)
            continue

        # 2) fallback por regex
        parsed = parse_by_regex(df)
        if parsed is not None and not parsed.empty:
            usable.append(parsed)

    return usable


# =========================
# UI
# =========================
uploaded_pdf = st.file_uploader("Botão de upload de **planilha de oposição** (PDF)", type=["pdf"])
go = st.button("Extrair e unir")

if uploaded_pdf is None:
    st.info("Envie um PDF para começar.")
    st.stop()

with st.spinner("Lendo PDF e identificando tabelas..."):
    pdf_bytes = uploaded_pdf.read()
    tables = process_pdf(pdf_bytes)

if not tables:
    st.error("Encontrei tabelas, mas **não consegui mapear nome/CPF/data**. "
             "Se possível, envie 1 página de exemplo para ajustarmos as heurísticas.")
    st.stop()

st.subheader("Pré-visualização (somente: nome / cpf / data de entrega)")
for i, tdf in enumerate(tables, start=1):
    with st.expander(f"Tabela {i} — {tdf.shape[0]} linhas", expanded=False):
        st.dataframe(tdf, use_container_width=True, height=320)

if go:
    final_df = pd.concat(tables, ignore_index=True)
    # limpeza final: tirar duplicadas e linhas vazias
    final_df = final_df.replace(r"^\s*$", pd.NA, regex=True).dropna(how="all").drop_duplicates().reset_index(drop=True)

    st.success(f"Tabelas unidas: {final_df.shape[0]} linhas")
    st.dataframe(final_df, use_container_width=True, height=420)

    csv_bytes = final_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Baixar CSV",
        data=csv_bytes,
        file_name="oposicoes_unificadas.csv",
        mime="text/csv",
    )
