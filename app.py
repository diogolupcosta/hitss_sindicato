import io
import re
import unicodedata
from pathlib import Path
from typing import List, Optional, Dict, Any
import pandas as pd
import pdfplumber
import streamlit as st

# função utilitária para converter DataFrame em bytes XLSX


def df_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Dados") -> bytes:
    buffer = io.BytesIO()
    try:
        # usa XlsxWriter se disponível (recomendado para escrita)
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            # auto-ajuste de largura de colunas
            ws = writer.sheets[sheet_name]
            for i, col in enumerate(df.columns):
                max_len = max([len(str(x))
                              for x in df[col].astype(str)] + [len(col)])
                ws.set_column(i, i, min(max_len + 2, 60))
    except Exception:
        # fallback para openpyxl caso xlsxwriter não esteja instalado
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer.getvalue()


FGTS_SAMPLE_NAME = "fgts.csv"
ATIVOS_SAMPLE_NAME = "Ativos_22102025_modelo.csv"


def _load_semicolon_csv(
    raw: bytes,
    *,
    encodings: tuple[str, ...] = ("utf-8-sig", "latin-1"),
    **read_csv_kwargs: Any,
) -> pd.DataFrame:
    """Carrega CSV separado por ';' tentando diferentes codificações antes de falhar."""
    last_error: Optional[Exception] = None
    for enc in encodings:
        try:
            buffer = io.BytesIO(raw)
            df = pd.read_csv(
                buffer,
                sep=";",
                decimal=",",
                quotechar='"',
                dtype=str,
                encoding=enc,
                **read_csv_kwargs,
            )
            return df
        except Exception as exc:  # captura UnicodeDecodeError e outros
            last_error = exc
            continue
    raise last_error if last_error else ValueError("Não foi possível decodificar o CSV fornecido.")


def load_fgts_csv(raw: bytes) -> pd.DataFrame:
    df = _load_semicolon_csv(raw)
    return df.applymap(lambda v: v.strip() if isinstance(v, str) else v)


def load_ativos_csv(raw: bytes) -> pd.DataFrame:
    df_raw = _load_semicolon_csv(raw, skiprows=8, header=None)
    df_raw = df_raw.applymap(lambda v: v.strip() if isinstance(v, str) else v)

    header_idx = None
    if not df_raw.empty:
        for idx, value in df_raw.iloc[:, 0].items():
            if isinstance(value, str) and value.strip().upper() == "MATRICULA":
                header_idx = idx
                break

    if header_idx is None:
        df = df_raw.copy()
    else:
        header = df_raw.loc[header_idx].fillna("")
        df = df_raw.loc[header_idx + 1:].reset_index(drop=True)
        df.columns = header

    columns = []
    for name in df.columns:
        if isinstance(name, str):
            name = name.strip()
        columns.append(name)

    clean_cols = []
    keep_indices = []
    for idx, name in enumerate(columns):
        if name and not str(name).startswith("Unnamed"):
            clean_cols.append(name)
            keep_indices.append(idx)

    if keep_indices:
        df = df.iloc[:, keep_indices]
        df.columns = dedupe_columns(clean_cols)

    if "CARGO" in df.columns:
        mask = df["CARGO"].str.upper().isin({"APRENDIZ OP MICROCOMPUTADOR", "ESTAGIARIO"})
        df = df.loc[~mask]

    return df.reset_index(drop=True)
    return df


# ========== Config ==========
st.set_page_config(page_title="Oposição - Extrator de Tabelas",
                   layout="wide", initial_sidebar_state="expanded")

# Logo + cabeçalho
LOGO_PATH = Path(__file__).with_name("Customer-Logos-09.png")
# remova vertical_alignment se der erro
col1, col2, col3 = st.columns([1, 0.4, 1])
with col2:
    st.image(str(LOGO_PATH), use_column_width=True)

st.markdown("<h1 style='text-align: center;'>Extrator de Tabelas do Sindicato </h1>",
            unsafe_allow_html=True)

# ========== Constantes (Oposição) ==========
IGNORES_RAW = [
    "SINDICATO     DOS     EMPREGADOS     EM     EMPRESAS     DE     PROCESSAMENTO     DE     DADOS,     DE SERVIÇOS    DE    COMPUTAÇÃO,    DE    INFORMÁTICA    E    TECNOLOGIA    DA    INFORMAÇÃO    E    DOS TRABALHADORES       EM       PROCESSAMENTO       DE       DADOS,       SERVIÇOS       DE       COMPUTAÇÃO, INFORMÁTICA E TECNOLOGIA DA INFORMAÇÃO DO ESTADO DE SÃO PAULO.",
    "HITSS DO BRASIL SERVICOS TECNOLOGICOS LTDA.FUNDADO EM 14/08/1984  -  CNPJ 55.537.666/0001-75  -   Av Angelica 35  -  Santa Cecilia  -  Sao Paulo  -  SP  -  CEP 01227-000www.sindpd.org.br  -  sindpd@sindpd.org.br  -  Central de Atendimento: ( 11 ) 3823-5600",
]
CPF_RE = re.compile(r"\b\d{3}\.\d{3}\.\d{3}-\d{2}\b")
DATE_RE = re.compile(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b")

# ========== Constantes (Associados) ==========
ASSOC_IGNORES_RAW = [
    "SINDICATO DOS EMPREGADOS EM EMPRESAS DE PROCESSAMENTO DE DADOS, DE SERVIÇOS DE COMPUTAÇÃO, DE INFORMÁTICA E TECNOLOGIA DA INFORMAÇÃO E DOS TRABALHADORES EM PROCESSAMENTO DE DADOS, SERVIÇOS DE COMPUTAÇÃO, INFORMÁTICA E TECNOLOGIA DA INFORMAÇÃO DO ESTADO DE SÃO PAULO.",
    "Sede: AV ANGELICA 35 SANTA CECILIA Sao Paulo Cep: 01227-000",
    "FUNDADO EM 14/08/1984 - CNPJ 55.537.666/0001-75",
    "Home Page: www.sindpd.org.br",
    "Fone: 3823-5600 Fax: 3667-6475",
]
# aceita 3-8 dígitos + hífen + 1 dígito (matrícula vista no documento)
MATRICULA_RE = re.compile(r"\b\d{3,8}-\d\b")
# valor no formato brasileiro: 1.234,56 ou 15,00
BRL_VAL_RE = re.compile(r"\b\d{1,3}(?:\.\d{3})*,\d{2}\b")
PAG_RE = re.compile(r"\bpag\.?\s*:\s*\d+\s*/\s*\d+", re.IGNORECASE)

# ========== Helpers ==========


def norm_txt(s: Optional[str]) -> str:
    if s is None:
        return ""
    if not isinstance(s, str):
        s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch)).lower()
    return re.sub(r"\s+", " ", s).strip()


IGNORES = [norm_txt(x) for x in IGNORES_RAW]
ASSOC_IGNORES = [norm_txt(x) for x in ASSOC_IGNORES_RAW]


def has_ignored_text(text: str) -> bool:
    return any(ign in norm_txt(text) for ign in IGNORES)


def has_ignored_text_assoc(text: str) -> bool:
    t = norm_txt(text)
    if any(ign in t for ign in ASSOC_IGNORES):
        return True
    if "relacao de associados do sindicato" in t:
        return True
    if PAG_RE.search(t):
        return True
    return False


def clean_cells(df: pd.DataFrame) -> pd.DataFrame:
    return df.applymap(lambda v: v.replace("\n", " ").strip() if isinstance(v, str) else v)


def drop_empty_rows_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
    df = df[[c for c in df.columns if not all(
        (str(x).strip() == "" or pd.isna(x)) for x in df[c])]]
    df = df.loc[~df.apply(lambda r: all(
        (str(x).strip() == "" or pd.isna(x)) for x in r), axis=1)]
    return df.reset_index(drop=True)


def is_mostly_int_column(series_or_df: pd.Series | pd.DataFrame, min_ratio: float = 0.8) -> bool:
    if isinstance(series_or_df, pd.DataFrame):
        series = series_or_df.apply(lambda r: next(
            (x for x in r if pd.notna(x) and str(x).strip() != ""), None), axis=1)
    else:
        series = series_or_df
    valid = series.dropna().astype(str).str.strip()
    if valid.empty:
        return False
    return (valid.str.match(r"^\d+$").sum() / len(valid)) >= min_ratio


def parse_brl_float(s: str) -> Optional[float]:
    if not s:
        return None
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None

# ========== Extração genérica de tabelas ==========


def extract_tables_from_pdf(pdf_bytes: bytes) -> List[pd.DataFrame]:
    dfs: List[pd.DataFrame] = []

    def safe_extract(page, settings: Dict[str, Any]):
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
        "vertical_strategy": "lines", "horizontal_strategy": "lines",
        "snap_tolerance": 3, "join_tolerance": 3,
        "min_words_vertical": 1, "min_words_horizontal": 1,
        "intersection_tolerance": 3,
    }
    stream = {
        "vertical_strategy": "text", "horizontal_strategy": "text",
        "intersection_tolerance": 5, "text_x_tolerance": 2, "text_y_tolerance": 2,
        "snap_tolerance": 3, "min_words_vertical": 1, "min_words_horizontal": 1,
        "join_tolerance": 3,
    }

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            for settings in (lattice, stream):
                for tbl in safe_extract(page, dict(settings)):
                    if tbl:
                        df = pd.DataFrame(tbl)
                        if df.shape[0] and df.shape[1]:
                            dfs.append(df)
    return dfs


def dedupe_columns(cols) -> list[str]:
    seen, out = {}, []
    for c in map(str, cols):
        if c not in seen:
            seen[c] = 0
            out.append(c)
        else:
            seen[c] += 1
            out.append(f"{c}.{seen[c]}")
    return out

# ========== Oposição: normalização ==========


def map_by_header(df: pd.DataFrame) -> Optional[pd.DataFrame]:
    df = drop_empty_rows_cols(clean_cells(df))
    if df.empty:
        return None
    header = [str(x) for x in df.iloc[0].tolist()]
    body = df.iloc[1:].reset_index(drop=True)
    body.columns = dedupe_columns(header)

    drop_cols = []
    for c in body.columns:
        if any(k in norm_txt(c) for k in ["cpf", "nome", "data"]):
            continue
        if is_mostly_int_column(body[c]):
            drop_cols.append(c)
    if drop_cols:
        body = body.drop(columns=drop_cols)

    norm_cols = [norm_txt(c) for c in body.columns]
    colmap: Dict[str, str] = {}
    for i, c in enumerate(norm_cols):
        if "cpf" in c:
            colmap["cpf"] = body.columns[i]
            break
    for i, c in enumerate(norm_cols):
        if "nome" in c or "empregado" in c or "trabalhador" in c:
            colmap["nome"] = body.columns[i]
            break
    for i, c in enumerate(norm_cols):
        if ("data" in c and "entreg" in c) or c == "data" or ("data" in c and "opos" in c):
            colmap["data de entrega"] = body.columns[i]
            break

    if "data de entrega" not in colmap:
        for col in body.columns:
            s = body[col].dropna().astype(str)
            if len(s) and (s.str.contains(DATE_RE).sum() / len(s)) >= 0.5:
                colmap["data de entrega"] = col
                break

    required = ["nome", "cpf", "data de entrega"]
    if not all(k in colmap for k in required):
        return None

    out = body[[colmap["nome"], colmap["cpf"],
                colmap["data de entrega"]]].copy()
    out.columns = required
    out = out.loc[~out.apply(lambda r: any(
        has_ignored_text(str(v)) for v in r), axis=1)]
    out = out.replace(r"^\s*$", pd.NA, regex=True).dropna(
        how="all").drop_duplicates().reset_index(drop=True)
    return out if not out.empty else None


def parse_by_regex(df: pd.DataFrame) -> Optional[pd.DataFrame]:
    df = drop_empty_rows_cols(clean_cells(df))
    if df.empty:
        return None
    rows = []
    for _, row in df.iterrows():
        cells = [str(x) for x in row.tolist() if str(x).strip()]
        if not cells:
            continue
        line = " ".join(cells)
        nline = norm_txt(line)
        if ("nome" in nline and "cpf" in nline and "data" in nline) or has_ignored_text(line):
            continue
        m_cpf, m_date = CPF_RE.search(line), DATE_RE.search(line)
        if not (m_cpf and m_date):
            joined = " ".join(cells)
            m_cpf, m_date = m_cpf or CPF_RE.search(
                joined), m_date or DATE_RE.search(joined)
        if m_cpf and m_date:
            nome = line[: m_cpf.start()].strip()
            cpf = m_cpf.group(0)
            data = m_date.group(0)
            if nome and CPF_RE.fullmatch(cpf) and DATE_RE.fullmatch(data) and not has_ignored_text(nome):
                rows.append({"nome": nome, "cpf": cpf,
                            "data de entrega": data})
    out = pd.DataFrame(rows).drop_duplicates(
    ).reset_index(drop=True) if rows else None
    return out if out is not None and not out.empty else None


def process_oposicao(pdf_bytes: bytes) -> List[pd.DataFrame]:
    usable: List[pd.DataFrame] = []
    for df in extract_tables_from_pdf(pdf_bytes):
        mapped = map_by_header(df)
        if mapped is not None and not mapped.empty:
            usable.append(mapped)
            continue
        parsed = parse_by_regex(df)
        if parsed is not None and not parsed.empty:
            usable.append(parsed)
    return usable

# ========== Associados: normalização ==========


def parse_associados_df(df: pd.DataFrame) -> Optional[pd.DataFrame]:
    df = drop_empty_rows_cols(clean_cells(df))
    if df.empty:
        return None
    rows = []
    for _, row in df.iterrows():
        cells = [str(x) for x in row.tolist() if str(x).strip()]
        if not cells:
            continue
        line = " ".join(cells)
        if has_ignored_text_assoc(line):
            continue
        m_mat = MATRICULA_RE.search(line)
        m_val = BRL_VAL_RE.search(line)
        if not (m_mat and m_val):
            continue
        matricula = m_mat.group(0)
        valor_raw = m_val.group(0)
        nome = line[m_mat.end(): m_val.start()].strip()
        valor = parse_brl_float(valor_raw)
        if nome and valor is not None:
            rows.append({"matricula": matricula, "nome": nome, "valor": valor})
    out = pd.DataFrame(rows).drop_duplicates(
    ).reset_index(drop=True) if rows else None
    return out if out is not None and not out.empty else None


def process_associados(pdf_bytes: bytes) -> pd.DataFrame:
    tables = extract_tables_from_pdf(pdf_bytes)
    parts: List[pd.DataFrame] = []
    for df in tables:
        parsed = parse_associados_df(df)
        if parsed is not None and not parsed.empty:
            parts.append(parsed)
    if not parts:
        return pd.DataFrame(columns=["matricula", "nome", "valor"])
    out = (pd.concat(parts, ignore_index=True)
             .replace(r"^\s*$", pd.NA, regex=True)
             .dropna(how="all")
             .drop_duplicates()
             .reset_index(drop=True))
    return out


# ========== UI quadripartida ==========
col_a, col_b, col_c, col_d = st.columns(4, gap="large")

# --- Coluna 1: Oposição ---
with col_a:
    st.subheader("Planilha de Oposição")
    up_op = st.file_uploader("Upload do Oposições (PDF)", type=[
                             "pdf"], key="oposicao")
    if up_op is None:
        st.info("Envie o PDF de Oposições para extrair (nome / cpf / data de entrega).")
    else:
        with st.spinner("Extraindo dados de Oposição..."):
            tables_op = process_oposicao(up_op.read())
        if not tables_op:
            st.error(
                "Não consegui mapear nome/CPF/data. Envie 1 página de exemplo para ajustarmos.")
        else:
            df_op = (pd.concat(tables_op, ignore_index=True)
                       .replace(r"^\s*$", pd.NA, regex=True)
                       .dropna(how="all")
                       .drop_duplicates()
                       .reset_index(drop=True))
            st.caption(f"{df_op.shape[0]} linhas • 3 colunas")
            st.dataframe(df_op, use_container_width=True, height=440)
            st.download_button(
                "Baixar XLSX (Oposição)",
                data=df_to_xlsx_bytes(df_op, sheet_name="Oposicao"),
                file_name="oposicoes_unificadas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


# --- Coluna 2: Relação de Associados ---
with col_b:
    st.subheader("Relação de Associados")
    up_assoc = st.file_uploader("Upload da Relação de Associados (PDF)", type=[
                                "pdf"], key="associados")
    if up_assoc is None:
        st.info(
            "Envie o PDF da Relação de Associados para extrair (matrícula / nome / valor).")
    else:
        with st.spinner("Extraindo dados da Relação de Associados..."):
            df_assoc = process_associados(up_assoc.read())
        if df_assoc.empty:
            st.error(
                "Não consegui mapear matrícula/nome/valor. Envie 1 página de exemplo para ajustarmos.")
        else:
            st.caption(f"{df_assoc.shape[0]} linhas • 3 colunas")
            st.dataframe(df_assoc, use_container_width=True, height=440)
        st.download_button(
            "Baixar XLSX (Associados)",
            data=df_to_xlsx_bytes(df_assoc, sheet_name="Associados"),
            file_name="associados_unificados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# --- Coluna 3: FGTS (CSV) ---
with col_c:
    st.subheader("FGTS (CSV)")
    uploaded_csv = st.file_uploader(
        "Upload do FGTS (CSV)", type=["csv"], key="fgts_csv")

    csv_bytes: Optional[bytes] = None

    if uploaded_csv is not None:
        csv_bytes = uploaded_csv.read()

    if csv_bytes is None:
        st.info(
            "Envie o arquivo FGTS em formato CSV para visualizar e baixar os dados.")
    else:
        try:
            df_fgts = load_fgts_csv(csv_bytes)
        except Exception as exc:
            st.error(f"Não foi possível ler o CSV: {exc}")
        else:
            st.caption(
                f"{df_fgts.shape[0]} linhas • {df_fgts.shape[1]} colunas")
            st.dataframe(df_fgts, use_container_width=True, height=440)
            st.download_button(
                "Baixar XLSX (FGTS)",
                data=df_to_xlsx_bytes(df_fgts, sheet_name="FGTS"),
                file_name="fgts.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


# --- Coluna 4: Ativos (CSV) ---
with col_d:
    st.subheader("Ativos (CSV)")
    uploaded_ativos = st.file_uploader(
        "Upload de Ativos (CSV)", type=["csv"], key="ativos_csv")

    ativos_bytes: Optional[bytes] = None

    if uploaded_ativos is not None:
        ativos_bytes = uploaded_ativos.read()

    if ativos_bytes is None:
        st.info("Envie o arquivo de Ativos em formato CSV para visualizar e baixar os dados.")
    else:
        try:
            df_ativos = load_ativos_csv(ativos_bytes)
        except Exception as exc:
            st.error(f"Não foi possível ler o CSV de Ativos: {exc}")
        else:
            st.caption(
                f"{df_ativos.shape[0]} linhas • {df_ativos.shape[1]} colunas")

            preview_cols = [
                "MATRICULA",
                "NOME",
                "SITUACAO",
                "DATA_ADMISSAO",
                "CARGO",
                "EMPRESA",
                "LOCALIZACAO",
                "DEPARTAMENTO",
                "SALARIO",
                "EMAIL",
            ]
            chosen_cols = [c for c in preview_cols if c in df_ativos.columns]
            df_preview = df_ativos[chosen_cols].copy() if chosen_cols else df_ativos

            st.dataframe(df_preview, use_container_width=True, height=440)
            st.download_button(
                "Baixar XLSX (Ativos)",
                data=df_to_xlsx_bytes(df_ativos, sheet_name="Ativos"),
                file_name="ativos.xlsx",
                mime="application/vnd.openxmlformats-officedocument-spreadsheetml.sheet",
            )
