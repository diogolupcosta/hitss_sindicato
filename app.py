import io
import re
import unicodedata
from pathlib import Path
from typing import List, Optional, Dict, Any

import pandas as pd
import pdfplumber
import streamlit as st

# funÃ§Ã£o utilitÃ¡ria para converter DataFrame em bytes XLSX
def df_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Dados") -> bytes:
    buffer = io.BytesIO()
    try:
        # usa XlsxWriter se disponÃ­vel (recomendado para escrita)
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            # auto-ajuste de largura de colunas
            ws = writer.sheets[sheet_name]
            for i, col in enumerate(df.columns):
                max_len = max([len(str(x)) for x in df[col].astype(str)] + [len(col)])
                ws.set_column(i, i, min(max_len + 2, 60))
    except Exception:
        # fallback para openpyxl caso xlsxwriter nÃ£o esteja instalado
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer.getvalue()


# ========== Config ==========
st.set_page_config(page_title="OposiÃ§Ã£o - Extrator de Tabelas", layout="wide", initial_sidebar_state="expanded")

# Logo + cabeÃ§alho
LOGO_PATH = Path(__file__).with_name("Customer-Logos-09.png")
col1, col2, col3 = st.columns([1, 0.4, 1])  # remova vertical_alignment se der erro
with col2:
    st.image(str(LOGO_PATH), use_column_width=True)

st.markdown("<h1 style='text-align: center;'>Extrator de Tabelas do Sindicato </h1>", unsafe_allow_html=True)

# ========== Constantes (OposiÃ§Ã£o) ==========
IGNORES_RAW = [
    "SINDICATO     DOS     EMPREGADOS     EM     EMPRESAS     DE     PROCESSAMENTO     DE     DADOS,     DE SERVIÃ‡OS    DE    COMPUTAÃ‡ÃƒO,    DE    INFORMÃTICA    E    TECNOLOGIA    DA    INFORMAÃ‡ÃƒO    E    DOS TRABALHADORES       EM       PROCESSAMENTO       DE       DADOS,       SERVIÃ‡OS       DE       COMPUTAÃ‡ÃƒO, INFORMÃTICA E TECNOLOGIA DA INFORMAÃ‡ÃƒO DO ESTADO DE SÃƒO PAULO.",
    "HITSS DO BRASIL SERVICOS TECNOLOGICOS LTDA.FUNDADO EM 14/08/1984  -  CNPJ 55.537.666/0001-75  -   Av Angelica 35  -  Santa Cecilia  -  Sao Paulo  -  SP  -  CEP 01227-000www.sindpd.org.br  -  sindpd@sindpd.org.br  -  Central de Atendimento: ( 11 ) 3823-5600",
]
CPF_RE = re.compile(r"\b\d{3}\.\d{3}\.\d{3}-\d{2}\b")
DATE_RE = re.compile(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b")
MONTH_YEAR_RE = re.compile(r"\b\d{1,2}/\d{4}\b")

# ========== Constantes (Associados) ==========
ASSOC_IGNORES_RAW = [
    "SINDICATO DOS EMPREGADOS EM EMPRESAS DE PROCESSAMENTO DE DADOS, DE SERVIÃ‡OS DE COMPUTAÃ‡ÃƒO, DE INFORMÃTICA E TECNOLOGIA DA INFORMAÃ‡ÃƒO E DOS TRABALHADORES EM PROCESSAMENTO DE DADOS, SERVIÃ‡OS DE COMPUTAÃ‡ÃƒO, INFORMÃTICA E TECNOLOGIA DA INFORMAÃ‡ÃƒO DO ESTADO DE SÃƒO PAULO.",
    "Sede: AV ANGELICA 35 SANTA CECILIA Sao Paulo Cep: 01227-000",
    "FUNDADO EM 14/08/1984 - CNPJ 55.537.666/0001-75",
    "Home Page: www.sindpd.org.br",
    "Fone: 3823-5600 Fax: 3667-6475",
]
# aceita 3-8 dÃ­gitos + hÃ­fen + 1 dÃ­gito (matrÃ­cula vista no documento)
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
    df = df[[c for c in df.columns if not all((str(x).strip() == "" or pd.isna(x)) for x in df[c])]]
    df = df.loc[~df.apply(lambda r: all((str(x).strip() == "" or pd.isna(x)) for x in r), axis=1)]
    return df.reset_index(drop=True)

def is_mostly_int_column(series_or_df: pd.Series | pd.DataFrame, min_ratio: float = 0.8) -> bool:
    if isinstance(series_or_df, pd.DataFrame):
        series = series_or_df.apply(lambda r: next((x for x in r if pd.notna(x) and str(x).strip() != ""), None), axis=1)
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

# ========== ExtraÃ§Ã£o genÃ©rica de tabelas ==========
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

# ========== OposiÃ§Ã£o: normalizaÃ§Ã£o ==========
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
        if "cpf" in c: colmap["cpf"] = body.columns[i]; break
    for i, c in enumerate(norm_cols):
        if "nome" in c or "empregado" in c or "trabalhador" in c: colmap["nome"] = body.columns[i]; break
    for i, c in enumerate(norm_cols):
        if ("data" in c and "entreg" in c) or c == "data" or ("data" in c and "opos" in c):
            colmap["data de entrega"] = body.columns[i]; break

    if "data de entrega" not in colmap:
        for col in body.columns:
            s = body[col].dropna().astype(str)
            if len(s) and (s.str.contains(DATE_RE).sum() / len(s)) >= 0.5:
                colmap["data de entrega"] = col
                break

    required = ["nome", "cpf", "data de entrega"]
    if not all(k in colmap for k in required):
        return None

    out = body[[colmap["nome"], colmap["cpf"], colmap["data de entrega"]]].copy()
    out.columns = required
    out = out.loc[~out.apply(lambda r: any(has_ignored_text(str(v)) for v in r), axis=1)]
    out = out.replace(r"^\s*$", pd.NA, regex=True).dropna(how="all").drop_duplicates().reset_index(drop=True)
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
            m_cpf, m_date = m_cpf or CPF_RE.search(joined), m_date or DATE_RE.search(joined)
        if m_cpf and m_date:
            nome = line[: m_cpf.start()].strip()
            cpf = m_cpf.group(0)
            data = m_date.group(0)
            if nome and CPF_RE.fullmatch(cpf) and DATE_RE.fullmatch(data) and not has_ignored_text(nome):
                rows.append({"nome": nome, "cpf": cpf, "data de entrega": data})
    out = pd.DataFrame(rows).drop_duplicates().reset_index(drop=True) if rows else None
    return out if out is not None and not out.empty else None

def process_oposicao(pdf_bytes: bytes) -> List[pd.DataFrame]:
    usable: List[pd.DataFrame] = []
    for df in extract_tables_from_pdf(pdf_bytes):
        mapped = map_by_header(df)
        if mapped is not None and not mapped.empty:
            usable.append(mapped); continue
        parsed = parse_by_regex(df)
        if parsed is not None and not parsed.empty:
            usable.append(parsed)
    return usable

# ========== Associados: normalizaÃ§Ã£o ==========
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
    out = pd.DataFrame(rows).drop_duplicates().reset_index(drop=True) if rows else None
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

# ========== UI bipartida ==========
col_a, col_b, col_c = st.columns(3, gap="large")

# --- Lado esquerdo: OposiÃ§Ã£o ---
with col_a:
    st.subheader("Planilha de OposiÃ§Ã£o")
    up_op = st.file_uploader("Upload do OposiÃ§Ãµes (PDF)", type=["pdf"], key="oposicao")
    if up_op is None:
        st.info("Envie o PDF de OposiÃ§Ãµes para extrair (nome / cpf / data de entrega).")
    else:
        with st.spinner("Extraindo dados de OposiÃ§Ã£o..."):
            tables_op = process_oposicao(up_op.read())
        if not tables_op:
            st.error("NÃ£o consegui mapear nome/CPF/data. Envie 1 pÃ¡gina de exemplo para ajustarmos.")
        else:
            df_op = (pd.concat(tables_op, ignore_index=True)
                       .replace(r"^\s*$", pd.NA, regex=True)
                       .dropna(how="all")
                       .drop_duplicates()
                       .reset_index(drop=True))
            st.caption(f"{df_op.shape[0]} linhas â€¢ 3 colunas")
            st.dataframe(df_op, use_container_width=True, height=440)
            st.download_button(
                "Baixar XLSX (OposiÃ§Ã£o)",
                data=df_to_xlsx_bytes(df_op, sheet_name="Oposicao"),
                file_name="oposicoes_unificadas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


# --- Lado direito: RelaÃ§Ã£o de Associados ---
with col_b:
    st.subheader("RelaÃ§Ã£o de Associados")
    up_assoc = st.file_uploader("Upload da RelaÃ§Ã£o de Associados (PDF)", type=["pdf"], key="associados")
    if up_assoc is None:
        st.info("Envie o PDF da RelaÃ§Ã£o de Associados para extrair (matrÃ­cula / nome / valor).")
    else:
        with st.spinner("Extraindo dados da RelaÃ§Ã£o de Associados..."):
            df_assoc = process_associados(up_assoc.read())
        if df_assoc.empty:
            st.error("NÃ£o consegui mapear matrÃ­cula/nome/valor. Envie 1 pÃ¡gina de exemplo para ajustarmos.")
        else:
            st.caption(f"{df_assoc.shape[0]} linhas â€¢ 3 colunas")
            st.dataframe(df_assoc, use_container_width=True, height=440)
        st.download_button(
            "Baixar XLSX (Associados)",
            data=df_to_xlsx_bytes(df_assoc, sheet_name="Associados"),
            file_name="associados_unificados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# --- Terceira coluna: Modelo FGTS (Sistema) ---
FGTS_IGNORES_RAW = [
    "caixa economica federal",
    "extrato fgts",
    "conta vinculada",
    "data/hora",
    "pÃ¡gina",
    "pagina",
    "total geral",
    "somatÃ³rio",
]
FGTS_IGNORES = [norm_txt(x) for x in FGTS_IGNORES_RAW]

def has_ignored_text_fgts(text: str) -> bool:
    t = norm_txt(text)
    if any(ign in t for ign in FGTS_IGNORES):
        return True
    # nÃºmeros de pÃ¡gina do tipo "pag: 1/5"
    if PAG_RE.search(t):
        return True
    return False

# HeurÃ­stica para corrigir textos extraÃ­dos ao contrÃ¡rio (e.g., "5202/90/91")
def _score_text_variant(s: str) -> int:
    score = 0
    t = s or ""
    # datas tipo dd/mm/aaaa
    if DATE_RE.fullmatch(t):
        score += 3
    # competÃªncias mm/aaaa
    if MONTH_YEAR_RE.fullmatch(t):
        score += 2
    # valores BRL
    if BRL_VAL_RE.fullmatch(t):
        score += 3
    # CPF
    if CPF_RE.fullmatch(t):
        score += 3
    # palavras-chave do domÃ­nio
    kw = [
        "nome", "trabalhador", "matricula", "categoria", "competencia",
        "apuracao", "referencia", "vencimento", "tipo", "deposito",
        "valor", "guia", "fgts", "total", "base", "remuneracao",
        "multa", "juros", "atualizacao", "cpf", "pis", "cnpj",
        "de", "do", "da", "dos", "das", "mensal",
    ]
    nt = norm_txt(t)
    if any(k in nt for k in kw):
        score += 1
    return score

def choose_best_orientation(val: Any) -> Any:
    if not isinstance(val, str):
        return val
    s = val.strip()
    if len(s) <= 1:
        return s
    rev = s[::-1]
    score_orig = _score_text_variant(s)
    score_rev = _score_text_variant(rev)
    if score_rev > score_orig:
        return rev
    return s


def choose_best_orientation_text(val: Any) -> Any:
    if not isinstance(val, str):
        return val
    s = val.strip()
    if len(s) <= 1:
        return s
    rev = s[::-1]
    score_orig = _score_text_variant(s)
    score_rev = _score_text_variant(rev)
    if score_rev >= score_orig:
        return rev
    return s


FGTS_COLUMN_ORDER = [
    "nome",
    "cpf",
    "pis_pasep",
    "matricula",
    "categoria",
    "tipo_deposito",
    "competencia_apuracao",
    "competencia_referencia",
    "vencimento",
    "valor_guia_fgts",
    "remuneracao_total_base",
    "multa",
    "juros",
    "atualizacao_monetaria",
    "total",
]

FGTS_COLUMN_DISPLAY = {
    "nome": "Nome Trabalhador",
    "cpf": "CPF",
    "pis_pasep": "PIS/PASEP",
    "matricula": "MatrÃ­cula",
    "categoria": "Categoria",
    "tipo_deposito": "Tipo DepÃ³sito",
    "competencia_apuracao": "Comp. ApuraÃ§Ã£o",
    "competencia_referencia": "Comp. ReferÃªncia",
    "vencimento": "Vencimento",
    "valor_guia_fgts": "Valor FGTS na Guia",
    "remuneracao_total_base": "RemuneraÃ§Ã£o Total Base",
    "multa": "Multa",
    "juros": "Juros",
    "atualizacao_monetaria": "AtualizaÃ§Ã£o MonetÃ¡ria",
    "total": "Total",
}

FGTS_WORD_SETTINGS = {
    "x_tolerance": 2,
    "y_tolerance": 2,
    "keep_blank_chars": False,
    "use_text_flow": True,
}
FGTS_COLUMN_TOLERANCE = 6.0
FGTS_MONTH_ONLY_RE = re.compile(r"\d{2}/\d{4}")
FGTS_FULLDATE_RE = re.compile(r"\d{2}/\d{2}/\d{4}")

def parse_mensal_block(text: str) -> Dict[str, str]:
    numbers = re.findall(r"\d{1,3}(?:\.\d{3})*,\d{2}", text)
    data = {
        "tipo_deposito": "Mensal" if "mensal" in norm_txt(text) else "",
        "remuneracao_total_base": numbers[0] if len(numbers) > 0 else "",
        "valor_guia_fgts": numbers[1] if len(numbers) > 1 else "",
        "multa": numbers[2] if len(numbers) > 2 else "0,00",
        "juros": numbers[3] if len(numbers) > 3 else "0,00",
        "atualizacao_monetaria": "0,00",
        "total": numbers[-1] if len(numbers) > 1 else (numbers[0] if numbers else ""),
    }
    if len(numbers) >= 6:
        data["atualizacao_monetaria"] = numbers[2]
        data["multa"] = numbers[3]
        data["juros"] = numbers[4]
        data["total"] = numbers[5]
    return data


def extract_matricula_cpf_from_token(text: str) -> tuple[str, str]:
    match = CPF_RE.search(text)
    if not match:
        return "", ""
    cpf = match.group(0)
    mat_raw = text[: match.start()]
    matricula = re.sub(r"[^0-9A-Z]", "", mat_raw)
    return matricula.strip(), cpf


def clean_name_fragment(text: str) -> str:
    if not text:
        return ""
    text = re.sub(r"\d", " ", text)
    text = text.replace("/", " ")
    text = text.replace(",", " ")
    text = text.replace(".", " ")
    text = text.replace(":", " ")
    text = text.replace(";", " ")
    text = re.sub(r"[^A-Za-zÀ-ÿ'\- ]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def build_fgts_record_from_column(words: List[Dict[str, Any]], mensal_word: Optional[Dict[str, Any]]) -> Optional[Dict[str, str]]:
    if mensal_word is None:
        return None
    record = parse_mensal_block(mensal_word.get("text", ""))
    words_sorted = sorted(words, key=lambda w: w.get("top", 0.0))
    name_parts: List[tuple[float, str]] = []
    comp_dates: List[tuple[float, str]] = []
    categoria: Optional[str] = None
    matricula, cpf = "", ""
    vencimento: str = ""

    for word in words_sorted:
        text = word.get("text", "")
        if not text or text == mensal_word.get("text"):
            continue

        if not cpf:
            mat_token, cpf_token = extract_matricula_cpf_from_token(text)
            if cpf_token:
                cpf = cpf_token
                if mat_token:
                    matricula = mat_token
                continue

        m_full = FGTS_FULLDATE_RE.search(text)
        if m_full:
            vencimento = m_full.group(0)
            continue

        residual = text
        month_match = FGTS_MONTH_ONLY_RE.search(residual)
        if month_match:
            comp_dates.append((word.get("top", 0.0), month_match.group(0)))
            residual = residual.replace(month_match.group(0), " ")

        if categoria is None:
            tail_digits = re.search(r"(\d{3})$", residual)
            if tail_digits:
                categoria = tail_digits.group(1)
                residual = residual[: -len(tail_digits.group(1))]

        fragment = clean_name_fragment(residual)
        if fragment:
            name_parts.append((word.get("top", 0.0), fragment))

    if not record.get("tipo_deposito"):
        record["tipo_deposito"] = "Mensal"

    record["matricula"] = matricula
    record["cpf"] = cpf
    record["vencimento"] = vencimento
    record["categoria"] = categoria or "101"

    comp_dates = sorted(comp_dates, key=lambda item: item[0], reverse=True)
    if comp_dates:
        record["competencia_apuracao"] = comp_dates[0][1]
        record["competencia_referencia"] = comp_dates[1][1] if len(comp_dates) > 1 else comp_dates[0][1]
    else:
        record["competencia_apuracao"] = ""
        record["competencia_referencia"] = ""

    name_parts = sorted(name_parts, key=lambda item: item[0], reverse=True)
    name = " ".join(part for _, part in name_parts).strip()
    record["nome"] = re.sub(r"\s+", " ", name)
    record.setdefault("pis_pasep", "")

    return record if record.get("nome") or record.get("matricula") else None


def parse_fgts_page_words(page) -> List[Dict[str, str]]:
    words = page.extract_words(**FGTS_WORD_SETTINGS)
    if not words:
        return []
    mensal_words = [w for w in words if "mensal" in norm_txt(w.get("text", ""))]
    if not mensal_words:
        return []
    mensal_words.sort(key=lambda w: w.get("x0", 0.0))
    col_positions = [w.get("x0", 0.0) for w in mensal_words]
    columns: Dict[int, Dict[str, Any]] = {i: {"words": [], "mensal": mensal_words[i]} for i in range(len(col_positions))}

    for word in words:
        x0 = word.get("x0", 0.0)
        idx = min(range(len(col_positions)), key=lambda i: abs(col_positions[i] - x0))
        if abs(col_positions[idx] - x0) <= FGTS_COLUMN_TOLERANCE:
            columns[idx]["words"].append(word)

    records: List[Dict[str, str]] = []
    for data in columns.values():
        rec = build_fgts_record_from_column(data.get("words", []), data.get("mensal"))
        if rec:
            records.append(rec)
    return records


def extract_fgts_records_from_words(pdf_bytes: bytes) -> List[Dict[str, str]]:
    records: List[Dict[str, str]] = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            records.extend(parse_fgts_page_words(page))
    return records
def map_fgts_label(raw: Any) -> Optional[str]:
    if raw is None:
        return None
    if not isinstance(raw, str):
        if pd.isna(raw):
            return None
        raw_str = str(raw).strip()
    else:
        raw_str = raw.strip()
    if not raw_str:
        return None
    nt = norm_txt(raw_str)
    if not nt:
        return None
    if "document" in nt or "observ" in nt:
        return None
    if "comp" in nt and "apur" in nt:
        return "competencia_apuracao"
    if "comp" in nt and ("refer" in nt or "ref" in nt):
        return "competencia_referencia"
    if "valor" in nt and "fgts" in nt and "guia" in nt:
        return "valor_guia_fgts"
    if "remuner" in nt and "base" in nt:
        return "remuneracao_total_base"
    if "atualiz" in nt or "monet" in nt:
        return "atualizacao_monetaria"
    if "tipo" in nt and "deposit" in nt:
        return "tipo_deposito"
    if "matric" in nt:
        return "matricula"
    if "categoria" in nt:
        return "categoria"
    if "venc" in nt:
        return "vencimento"
    if "multa" in nt:
        return "multa"
    if "juros" in nt:
        return "juros"
    if "cpf" in nt:
        return "cpf"
    if "pis" in nt:
        return "pis_pasep"
    if "nome" in nt:
        return "nome"
    if "total" in nt:
        return "total"
    if "guia" in nt and "fgts" in nt:
        return "valor_guia_fgts"
    return None


def sanitize_fgts_value(val: Any) -> Any:
    if isinstance(val, str):
        cleaned = val.strip().strip("'\"")
        if cleaned == "":
            return pd.NA
        if cleaned.lower() in {"nan", "none"}:
            return pd.NA
        return cleaned
    return val


def reshape_fgts_vertical_table(df: pd.DataFrame) -> Optional[pd.DataFrame]:
    if df.shape[1] <= 1:
        return None
    label_series = df.iloc[:, 0].apply(choose_best_orientation_text)
    label_series = label_series.apply(lambda x: x.strip() if isinstance(x, str) else x)
    labels = label_series.tolist()
    mapped_labels = [map_fgts_label(lbl) for lbl in labels]
    if not any(mapped_labels):
        return None
    data = df.iloc[:, 1:].copy()
    data = data.applymap(choose_best_orientation)
    data = data.applymap(lambda v: v.strip() if isinstance(v, str) else v)
    records: list[Dict[str, Any]] = []
    for col_idx in range(data.shape[1]):
        record: Dict[str, Any] = {}
        for row_idx, label in enumerate(mapped_labels):
            if label is None:
                continue
            value = data.iat[row_idx, col_idx]
            if isinstance(value, str):
                value = value.strip()
            if value in (None, ""):
                continue
            record[label] = sanitize_fgts_value(value)
        if not record:
            continue
        if not any(record.get(k) for k in ("nome", "cpf", "matricula")):
            continue
        records.append(record)
    if not records:
        return None
    out = pd.DataFrame(records)
    bias_text_cols = {
        "nome",
        "categoria",
        "tipo_deposito",
    }
    for col in bias_text_cols & set(out.columns):
        out[col] = out[col].apply(choose_best_orientation_text)
        out[col] = out[col].astype(str).str.strip()

    strict_text_cols = {
        "competencia_apuracao",
        "competencia_referencia",
        "vencimento",
        "pis_pasep",
        "matricula",
        "cpf",
    }
    for col in strict_text_cols & set(out.columns):
        out[col] = out[col].apply(choose_best_orientation)
        out[col] = out[col].astype(str).str.strip()

    value_cols = {
        "valor_guia_fgts",
        "remuneracao_total_base",
        "multa",
        "juros",
        "atualizacao_monetaria",
        "total",
    }
    for col in value_cols & set(out.columns):
        out[col] = out[col].apply(choose_best_orientation)
        out[col] = out[col].apply(sanitize_fgts_value)
    return drop_empty_rows_cols(out)

def extract_tables_from_pdf_advanced(pdf_bytes: bytes, top_margin_ratio: float = 0.08, bottom_margin_ratio: float = 0.08) -> list[pd.DataFrame]:
    dfs: list[pd.DataFrame] = []

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
            for angle in (0, 90, 270):
                # Try a safe rotation if supported; otherwise, skip rotation
                if angle == 0:
                    rpage = page
                else:
                    rpage = None
                    if hasattr(page, "with_rotation"):
                        try:
                            rpage = page.with_rotation(angle)
                        except Exception:
                            rpage = None
                    if rpage is None:
                        # rotation not supported; continue without this angle
                        continue

                h, w = rpage.height, rpage.width
                y0 = max(0, h * bottom_margin_ratio)
                y1 = min(h, h * (1 - top_margin_ratio))
                cpage = rpage.crop((0, y0, w, y1))
                for settings in (lattice, stream):
                    for tbl in safe_extract(cpage, dict(settings)):
                        if not tbl:
                            continue
                        df = pd.DataFrame(tbl)
                        if df.shape[0] and df.shape[1]:
                            dfs.append(df)
    return dfs

def process_fgts_vertical(pdf_bytes: bytes) -> list[pd.DataFrame]:
    word_records = extract_fgts_records_from_words(pdf_bytes)
    if word_records:
        df_word = pd.DataFrame(word_records)
        if not df_word.empty:
            df_word = df_word.replace(r"^\s*$", pd.NA, regex=True).fillna("")
            for col in FGTS_COLUMN_ORDER:
                if col not in df_word.columns:
                    if col in {"multa", "juros", "atualizacao_monetaria"}:
                        df_word[col] = "0,00"
                    else:
                        df_word[col] = ""
            df_word = df_word[[c for c in FGTS_COLUMN_ORDER if c in df_word.columns] + [c for c in df_word.columns if c not in FGTS_COLUMN_ORDER]]
            return [df_word.reset_index(drop=True)]

    raw_tables = extract_tables_from_pdf_advanced(pdf_bytes)
    reshaped_tables: list[pd.DataFrame] = []
    for df in raw_tables:
        df2 = drop_empty_rows_cols(clean_cells(df))
        if df2.empty:
            continue
        df2 = df2.applymap(choose_best_orientation)
        df2 = drop_empty_rows_cols(clean_cells(df2))
        mask = []
        for _, row in df2.iterrows():
            row_text = " ".join(str(x) for x in row.tolist() if pd.notna(x)).strip()
            mask.append(not has_ignored_text_fgts(row_text))
        if mask:
            df2 = df2.loc[mask]
        df2 = drop_empty_rows_cols(df2)
        if df2.empty:
            continue
        reshaped = reshape_fgts_vertical_table(df2)
        if reshaped is not None and not reshaped.empty:
            reshaped_tables.append(reshaped.reset_index(drop=True))
    return reshaped_tables

with col_c:
    st.subheader("Modelo FGTS (Sistema)")
    up_fgts = st.file_uploader("Upload do FGTS (PDF)", type=["pdf"], key="fgts")
    if up_fgts is None:
        # se existir o arquivo de exemplo na raiz do projeto, ofereÃ§a um botÃ£o de teste
        sample_path = Path.cwd().parent / "modelo FGTS sistema.pdf"
        if sample_path.exists():
            if st.button("Usar arquivo de exemplo (modelo FGTS sistema.pdf)"):
                with open(sample_path, "rb") as f:
                    up_fgts = io.BytesIO(f.read())
        if up_fgts is None:
            st.info("Envie o PDF do modelo FGTS (tabelas verticais) para prÃ©-visualizaÃ§Ã£o.")
    if up_fgts is not None:
        pdf_bytes = up_fgts.read() if hasattr(up_fgts, "read") else up_fgts.getvalue()
        with st.spinner("Extraindo tabelas do FGTS (inclui rotaÃ§Ã£o e remoÃ§Ã£o de cabeÃ§alho/rodapÃ©)..."):
            tables_fgts = process_fgts_vertical(pdf_bytes)
        if not tables_fgts:
            st.error("NÃ£o consegui localizar tabelas Ãºteis no FGTS. Verifique se o layout Ã© compatÃ­vel.")
        else:
            df_fgts = (pd.concat(tables_fgts, ignore_index=True)
                         .replace(r"^\s*$", pd.NA, regex=True)
                         .dropna(how="all")
                         .drop_duplicates()
                         .reset_index(drop=True))
            ordered_cols = [c for c in FGTS_COLUMN_ORDER if c in df_fgts.columns]
            remaining_cols = [c for c in df_fgts.columns if c not in ordered_cols]
            df_fgts = df_fgts[ordered_cols + remaining_cols]
            display_df = df_fgts.rename(columns=FGTS_COLUMN_DISPLAY)
            st.caption(f"{len(tables_fgts)} tabela(s) processadas â€¢ {display_df.shape[0]} linhas â€¢ {display_df.shape[1]} colunas")
            st.dataframe(display_df, use_container_width=True, height=440)
            st.download_button(
                "Baixar XLSX (FGTS)",
                data=df_to_xlsx_bytes(display_df, sheet_name="FGTS"),
                file_name="fgts_tabelas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )






