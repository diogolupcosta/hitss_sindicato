import io
import re
import unicodedata
from pathlib import Path
from typing import List, Optional, Dict, Any, Tuple

import pandas as pd
import pdfplumber
import streamlit as st
from pypdf import PdfReader, PdfWriter

# ========= util XLSX =========
def df_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Dados") -> bytes:
    buffer = io.BytesIO()
    try:
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            ws = writer.sheets[sheet_name]
            for i, col in enumerate(df.columns):
                max_len = max([len(str(x)) for x in df[col].astype(str)] + [len(col)])
                ws.set_column(i, i, min(max_len + 2, 60))
    except Exception:
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer.getvalue()

# ========= Config =========
st.set_page_config(page_title="Oposição - Extrator de Tabelas", layout="wide", initial_sidebar_state="expanded")

# Logo + cabeçalho
LOGO_PATH = Path(__file__).with_name("Customer-Logos-09.png")
col1, col2, col3 = st.columns([1, 0.4, 1])
with col2:
    try:
        st.image(str(LOGO_PATH), use_container_width=True)
    except TypeError:
        st.image(str(LOGO_PATH), use_column_width=True)

st.markdown("<h1 style='text-align: center;'>Extrator de Tabelas do Sindicato</h1>", unsafe_allow_html=True)

# ========= Constantes (Oposição) =========
IGNORES_RAW = [
    "SINDICATO     DOS     EMPREGADOS     EM     EMPRESAS     DE     PROCESSAMENTO     DE     DADOS,     DE SERVIÇOS    DE    COMPUTAÇÃO,    DE    INFORMÁTICA    E    TECNOLOGIA    DA    INFORMAÇÃO    E    DOS TRABALHADORES       EM       PROCESSAMENTO       DE       DADOS,       SERVIÇOS       DE       COMPUTAÇÃO, INFORMÁTICA E TECNOLOGIA DA INFORMAÇÃO DO ESTADO DE SÃO PAULO.",
    "HITSS DO BRASIL SERVICOS TECNOLOGICOS LTDA.FUNDADO EM 14/08/1984  -  CNPJ 55.537.666/0001-75  -   Av Angelica 35  -  Santa Cecilia  -  Sao Paulo  -  SP  -  CEP 01227-000www.sindpd.org.br  -  sindpd@sindpd.org.br  -  Central de Atendimento: ( 11 ) 3823-5600",
]
CPF_RE = re.compile(r"\b\d{3}\.\d{3}\.\d{3}-\d{2}\b")
DATE_RE = re.compile(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b")

# ========= Constantes (Associados) =========
ASSOC_IGNORES_RAW = [
    "SINDICATO DOS EMPREGADOS EM EMPRESAS DE PROCESSAMENTO DE DADOS, DE SERVIÇOS DE COMPUTAÇÃO, DE INFORMÁTICA E TECNOLOGIA DA INFORMAÇÃO E DOS TRABALHADORES EM PROCESSAMENTO DE DADOS, SERVIÇOS DE COMPUTAÇÃO, INFORMÁTICA E TECNOLOGIA DA INFORMAÇÃO DO ESTADO DE SÃO PAULO.",
    "Sede: AV ANGELICA 35 SANTA CECILIA Sao Paulo Cep: 01227-000",
    "FUNDADO EM 14/08/1984 - CNPJ 55.537.666/0001-75",
    "Home Page: www.sindpd.org.br",
    "Fone: 3823-5600 Fax: 3667-6475",
]
MATRICULA_RE = re.compile(r"\b\d{3,8}-\d\b")
BRL_VAL_RE = re.compile(r"\b\d{1,3}(?:\.\d{3})*,\d{2}\b")
PAG_RE = re.compile(r"\bpag\.?\s*:\s*\d+\s*/\s*\d+", re.IGNORECASE)

# ========= Helpers comuns =========
def norm_txt(s: Optional[str]) -> str:
    if s is None:
        return ""
    if not isinstance(s, str):
        s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch)).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
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
    s = str(s).replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None

# ========= Rotação de PDF =========
def rotate_pdf_bytes(pdf_bytes: bytes, degrees_clockwise: int) -> bytes:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    for page in reader.pages:
        try:
            page.rotate(degrees_clockwise)
        except Exception:
            page.rotate_clockwise(degrees_clockwise)
        writer.add_page(page)
    out = io.BytesIO()
    writer.write(out)
    out.seek(0)
    return out.getvalue()

# ========= Extração genérica de tabelas =========
def extract_tables_from_pdf(pdf_bytes: bytes, rotate_degrees: int = 0) -> List[pd.DataFrame]:
    if rotate_degrees:
        pdf_bytes = rotate_pdf_bytes(pdf_bytes, rotate_degrees)

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

# ========= Oposição =========
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

# ========= Associados =========
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

# ========= FGTS – REFINADO =========
FGTS_CANON = [
    "comp. apuração",
    "comp. referência",
    "nome trabalhador",
    "matrícula",
    "cpf",
    "categoria",
    "vencimento",
    "tipo depósito",
    "base remuneração total",
    "valor fgts na guia",
    "juros",
    "atualiz. monetária",
    "multa",
    "total",
]

FGTS_ALIASES = {
    "comp apuracao": "comp. apuração",
    "comp apuração": "comp. apuração",
    "comp referencia": "comp. referência",
    "nome trabalhador": "nome trabalhador",
    "matricula": "matrícula",
    "cpf": "cpf",
    "categoria": "categoria",
    "vencimento": "vencimento",
    "tipo deposito": "tipo depósito",
    "base remuneracao total": "base remuneração total",
    "valor fgts na guia": "valor fgts na guia",
    "juros": "juros",
    "atualiz monetaria": "atualiz. monetária",
    "atualiz  monetaria": "atualiz. monetária",
    "multa": "multa",
    "total": "total",
}

def _best_header_map(df: pd.DataFrame) -> Optional[Tuple[pd.DataFrame, Dict[str, int]]]:
    if df.shape[0] < 2:
        return None
    header = [str(x) for x in df.iloc[0].tolist()]
    body = df.iloc[1:].reset_index(drop=True)
    colmap: Dict[str, int] = {}
    for idx, h in enumerate(header):
        key = norm_txt(h)
        if key in FGTS_ALIASES:
            colmap[FGTS_ALIASES[key]] = idx
    if len(colmap) >= 6:
        return body, colmap
    return None

def _words_to_df_by_grid(page, header_words: Dict[str, Tuple[float, float, float, float]]) -> pd.DataFrame:
    """
    Reconstrói a tabela extraindo TODAS as palavras da linha e depois mapeando por padrões.
    """
    if "nome trabalhador" not in header_words or "cpf" not in header_words:
        return pd.DataFrame(columns=FGTS_CANON)
    
    header_bottom = header_words["nome trabalhador"][3]
    
    words = page.extract_words(keep_blank_chars=True, use_text_flow=False, x_tolerance=1, y_tolerance=2)
    if not words:
        return pd.DataFrame(columns=FGTS_CANON)
    
    table_words = [w for w in words if w["top"] > header_bottom + 2]

    # Agrupa palavras por linha (Y)
    table_words.sort(key=lambda w: w["top"])
    rows: List[List[dict]] = []
    current: List[dict] = []
    last_top = None
    row_gap = 5

    for w in table_words:
        if last_top is None or abs(w["top"] - last_top) <= row_gap:
            current.append(w)
            last_top = w["top"] if last_top is None else (last_top + w["top"]) / 2
        else:
            if current:
                rows.append(current)
            current = [w]
            last_top = w["top"]
    if current:
        rows.append(current)

    out_rows = []
    for row_words in rows:
        # Ordena palavras da linha por posição X
        row_words_sorted = sorted(row_words, key=lambda w: w["x0"])
        
        # Concatena todo o texto da linha para análise
        full_line = " ".join([w["text"] for w in row_words_sorted])
        
        # Extrai campos por padrões de regex
        comp = ""
        comp_match = re.search(r'\b(\d{2}/\d{4})\b', full_line)
        if comp_match:
            comp = comp_match.group(1)
        
        cpf = ""
        cpf_match = re.search(r'\b(\d{3}\.\d{3}\.\d{3}-\d{2})\b', full_line)
        if cpf_match:
            cpf = cpf_match.group(1)
        
        # Se não tem CPF, provavelmente não é linha de dados válida
        if not cpf:
            continue
        
        matricula = ""
        mat_match = re.search(r'\b(\d{18,25})\b', full_line)
        if mat_match:
            matricula = mat_match.group(1)
        
        categoria = ""
        cat_match = re.search(r'\b(10[0-9]|11[0-9])\b', full_line)
        if cat_match:
            categoria = cat_match.group(1)
        
        vencimento = ""
        venc_match = re.search(r'\b(\d{2}/\d{2}/\d{4})\b', full_line)
        if venc_match:
            vencimento = venc_match.group(1)
        
        tipo_deposito = ""
        if 'Mensal' in full_line or 'mensal' in full_line.lower():
            tipo_deposito = "Mensal"
        
        # Nome: extrai da posição X do cabeçalho "nome trabalhador"
        nome = ""
        if "nome trabalhador" in header_words:
            nome_x0, _, nome_x1, _ = header_words["nome trabalhador"]
            nome_words = [w for w in row_words if nome_x0 - 20 <= w["x0"] <= nome_x1 + 20]
            if nome_words:
                nome_parts = []
                for w in sorted(nome_words, key=lambda x: x["x0"]):
                    # Pega apenas palavras em maiúsculas (nomes próprios)
                    if w["text"] and (w["text"].isupper() or w["text"][0].isupper()):
                        nome_parts.append(w["text"])
                nome = " ".join(nome_parts)
        
        # Valores monetários: extrai TODOS os números no formato brasileiro
        valores_raw = re.findall(r'\b(\d{1,3}(?:\.\d{3})*,\d{2})\b', full_line)
        valores = [parse_brl_float(v) for v in valores_raw]
        
        # Mapeia valores para as colunas (ordem esperada do PDF)
        # Base Remuneração, Valor FGTS, Juros, Atualiz, Multa, Total
        base_remun = valores[0] if len(valores) > 0 else None
        valor_fgts = valores[1] if len(valores) > 1 else None
        juros = valores[2] if len(valores) > 2 else 0.0
        atualiz = valores[3] if len(valores) > 3 else 0.0
        multa = valores[4] if len(valores) > 4 else 0.0
        total = valores[5] if len(valores) > 5 else None
        
        # Se não tem total mas tem valor FGTS, assume que são iguais
        if total is None and valor_fgts is not None:
            total = valor_fgts
        
        row_data = {
            "comp. apuração": comp,
            "comp. referência": "",
            "nome trabalhador": nome,
            "matrícula": matricula,
            "cpf": cpf,
            "categoria": categoria,
            "vencimento": vencimento,
            "tipo depósito": tipo_deposito,
            "base remuneração total": base_remun,
            "valor fgts na guia": valor_fgts,
            "juros": juros,
            "atualiz. monetária": atualiz,
            "multa": multa,
            "total": total,
        }
        
        out_rows.append(row_data)

    return pd.DataFrame(out_rows)

def process_fgts(pdf_bytes: bytes) -> pd.DataFrame:
    pdf_bytes = rotate_pdf_bytes(pdf_bytes, 90)

    # Usa extract_tables que JÁ FUNCIONA para capturar as linhas
    tables = extract_tables_from_pdf(pdf_bytes, rotate_degrees=0)

    best = None
    best_body = None
    for df in tables:
        m = _best_header_map(df)
        if m:
            best = m
            best_body = df
            break

    if not best:
        return pd.DataFrame(columns=FGTS_CANON)
    
    body, colmap = best
    
    # Processa linha por linha
    out_rows = []
    for idx, row in body.iterrows():
        row_data = {}
        
        # Extrai valores de cada coluna mapeada
        for can in FGTS_CANON:
            if can in colmap:
                val = str(row.iloc[colmap[can]]).strip() if pd.notna(row.iloc[colmap[can]]) else ""
            else:
                val = ""
            
            # Conversão especial para campos numéricos
            if can in ["base remuneração total", "valor fgts na guia", "juros", "atualiz. monetária", "multa", "total"]:
                parsed = parse_brl_float(val) if re.search(r"\d", val) else None
                row_data[can] = parsed
            else:
                row_data[can] = val
        
        # Se a linha não tem valores numéricos mas tem nome, tenta buscar na linha original da tabela
        if (row_data.get("base remuneração total") is None and 
            row_data.get("valor fgts na guia") is None and 
            row_data.get("nome trabalhador")):
            
            # Pega a linha RAW completa da tabela original
            if idx < len(best_body):
                raw_row = best_body.iloc[idx].tolist()
                raw_line = " ".join([str(x) for x in raw_row if pd.notna(x)])
                
                # Extrai TODOS os valores monetários da linha raw
                valores_raw = re.findall(r'\b(\d{1,3}(?:\.\d{3})*,\d{2})\b', raw_line)
                valores = [parse_brl_float(v) for v in valores_raw]
                
                # Mapeia valores
                if len(valores) >= 2:
                    row_data["base remuneração total"] = valores[0]
                    row_data["valor fgts na guia"] = valores[1]
                    row_data["juros"] = valores[2] if len(valores) > 2 else 0.0
                    row_data["atualiz. monetária"] = valores[3] if len(valores) > 3 else 0.0
                    row_data["multa"] = valores[4] if len(valores) > 4 else 0.0
                    row_data["total"] = valores[5] if len(valores) > 5 else valores[1]
        
        out_rows.append(row_data)
    
    out = pd.DataFrame(out_rows)
    out = out.replace(r"^\s*$", pd.NA, regex=True).dropna(how="all")
    
    # Garante todas as colunas
    for c in FGTS_CANON:
        if c not in out.columns:
            out[c] = pd.NA
    
    out = out[FGTS_CANON].drop_duplicates().reset_index(drop=True)
    
    return out

# ========= UI (3 colunas) =========
col_op, col_assoc, col_fgts = st.columns(3, gap="large")

# --- Oposição ---
with col_op:
    st.subheader("Planilha de Oposição")
    up_op = st.file_uploader("Upload do Oposições (PDF)", type=["pdf"], key="oposicao")
    if up_op is None:
        st.info("Envie o PDF de Oposições para extrair (nome / cpf / data de entrega).")
    else:
        with st.spinner("Extraindo dados de Oposição..."):
            tables_op = process_oposicao(up_op.read())
        if not tables_op:
            st.error("Não consegui mapear nome/CPF/data. Envie 1 página de exemplo para ajustarmos.")
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

# --- Associados ---
with col_assoc:
    st.subheader("Relação de Associados")
    up_assoc = st.file_uploader("Upload da Relação de Associados (PDF)", type=["pdf"], key="associados")
    if up_assoc is None:
        st.info("Envie o PDF da Relação de Associados para extrair (matrícula / nome / valor).")
    else:
        with st.spinner("Extraindo dados da Relação de Associados..."):
            df_assoc = process_associados(up_assoc.read())
        if df_assoc.empty:
            st.error("Não consegui mapear matrícula/nome/valor. Envie 1 página de exemplo para ajustarmos.")
        else:
            st.caption(f"{df_assoc.shape[0]} linhas • 3 colunas")
            st.dataframe(df_assoc, use_container_width=True, height=440)
            st.download_button(
                "Baixar XLSX (Associados)",
                data=df_to_xlsx_bytes(df_assoc, sheet_name="Associados"),
                file_name="associados_unificados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# --- FGTS ---
with col_fgts:
    st.subheader("FGTS – Relação de Trabalhadores (14 colunas)")
    up_fgts = st.file_uploader("Upload do PDF FGTS (em retrato; será girado 90°)", type=["pdf"], key="fgts")
    if up_fgts is None:
        st.info("Envie o PDF para extrair as colunas (Comp. Apuração, Comp. Referência, …, Total).")
    else:
        with st.spinner("Girando páginas e extraindo dados do FGTS..."):
            df_fgts = process_fgts(up_fgts.read())

        if df_fgts.empty:
            st.error("Não consegui mapear as colunas do FGTS após a rotação. Se persistir, me envie 1 página para refino.")
        else:
            st.caption(f"{df_fgts.shape[0]} linhas • {df_fgts.shape[1]} colunas")
            st.dataframe(df_fgts, use_container_width=True, height=440)
            st.download_button(
                "Baixar XLSX (FGTS)",
                data=df_to_xlsx_bytes(df_fgts, sheet_name="FGTS"),
                file_name="fgts_relacao_trabalhadores.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )