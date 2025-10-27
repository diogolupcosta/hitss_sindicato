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
                max_len = max([len(str(x)) for x in df[col].astype(str)] + [len(col)])
                ws.set_column(i, i, min(max_len + 2, 60))
    except Exception:
        # fallback para openpyxl caso xlsxwriter não esteja instalado
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer.getvalue()


# ========== Config ==========
st.set_page_config(page_title="Oposição - Extrator de Tabelas", layout="wide", initial_sidebar_state="expanded")

# Logo + cabeçalho
LOGO_PATH = Path(__file__).with_name("Customer-Logos-09.png")
col1, col2, col3 = st.columns([1, 0.4, 1])  # remova vertical_alignment se der erro
with col2:
    st.image(str(LOGO_PATH), use_column_width=True)

st.markdown("<h1 style='text-align: center;'>Extrator de Tabelas do Sindicato </h1>", unsafe_allow_html=True)

# ========== Constantes (Oposição) ==========
IGNORES_RAW = [
    "SINDICATO     DOS     EMPREGADOS     EM     EMPRESAS     DE...VIÇOS       DE       COMPUTAÇÃO, INFORMÁTICA E TECNOLOGIA DA INFORMAÇÃO DO ESTADO DE SÃO PAULO.",
    "HITSS DO BRASIL SERVICOS TECNOLOGICOS LTDA.FUNDADO EM 14/08...dpd@sindpd.org.br  -  Central de Atendimento: ( 11 ) 3823-5600",
]
CPF_RE = re.compile(r"\b\d{3}\.\d{3}\.\d{3}-\d{2}\b")
DATE_RE = re.compile(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b")

# ========== Constantes (Associados) ==========
ASSOC_IGNORES_RAW = [
    "SINDICATO DOS EMPREGADOS EM EMPRESAS DE PROCESSAMENTO DE DA...INFORMÁTICA E TECNOLOGIA DA INFORMAÇÃO DO ESTADO DE SÃO PAULO.",
    "Sede: AV ANGELICA 35 SANTA CECILIA Sao Paulo Cep: 01227-000",
    "FUNDADO EM 14/08/1984 - CNPJ 55.537.666/0001-75",
    "Home Page: www.sindpd.org.br",
    "Fone: 3823-5600 Fax: 3667-6475",
]
# aceita 3-8 dígitos + hífen + 1 dígito (matrícula vista no documento)
MATRICULA_RE = re.compile(r"\b\d{3,8}-\d\b")

PAG_RE = re.compile(r"\bp[aá]gina\s*\d+\s*(de|/)\s*\d+\b", re.IGNORECASE)

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
    s = series.dropna().astype(str).str.replace(r"[.\s]", "", regex=True)
    is_int = s.str.fullmatch(r"\d+")
    if not len(s):
        return False
    return (is_int.sum() / len(s)) >= min_ratio

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

    # remove colunas 100% numéricas que não são chaves
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
    # valida: linhas com CPF/Data válidos
    out = out.dropna(how="any")
    mask = out["cpf"].astype(str).str.fullmatch(CPF_RE).fillna(False) & out["data de entrega"].astype(str).str.fullmatch(DATE_RE).fillna(False)
    out = out.loc[mask].reset_index(drop=True)
    return out if not out.empty else None

def extract_tables_from_pdf(pdf_bytes: bytes) -> List[pd.DataFrame]:
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
                        yield df

def parse_by_regex(df: pd.DataFrame) -> Optional[pd.DataFrame]:
    df = clean_cells(df)
    rows: list[dict] = []

    for _, row in df.iterrows():
        cells = [v for v in row.tolist() if isinstance(v, str) and v.strip()]
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

# ========== Associados: normalização ==========
def parse_associados_df(df: pd.DataFrame) -> Optional[pd.DataFrame]:
    df = drop_empty_rows_cols(clean_cells(df))
    if df.empty:
        return None

    # tenta identificar colunas esperadas
    cols_norm = [norm_txt(c) for c in df.iloc[0].astype(str).tolist()]
    body = df.iloc[1:].reset_index(drop=True)

    # fallback: se a primeira linha não parece cabeçalho, usa heurística por conteúdo
    if not any("matric" in c or "cpf" in c or "valor" in c or "nome" in c for c in cols_norm):
        body = df.copy()
        body.columns = [f"col_{i}" for i in range(body.shape[1])]
    else:
        body.columns = dedupe_columns(df.iloc[0].astype(str).tolist())

    # limpa e reduz
    body = drop_empty_rows_cols(clean_cells(body))

    # mapeia colunas por nome aproximado
    colmap = {}
    for i, c in enumerate([norm_txt(x) for x in body.columns]):
        if "matric" in c or "registro" in c:
            colmap["matricula"] = body.columns[i]
        if "nome" in c or "empregado" in c or "trabalhador" in c:
            colmap["nome"] = body.columns[i]
        if "valor" in c or "mensal" in c or "contrib" in c or "descont" in c:
            colmap["valor"] = body.columns[i]

    # fallback por conteúdo
    if "valor" not in colmap:
        for col in body.columns:
            s = body[col].dropna().astype(str)
            if len(s) and is_mostly_int_column(s.str.replace(r"[^\d]", "", regex=True)):
                colmap["valor"] = col
                break

    # se não há 2+ colunas úteis, aborta
    need = [k for k in ["matricula", "nome", "valor"] if k in colmap]
    if len(need) < 2:
        return None

    out = body[[colmap[k] for k in need]].copy()
    out.columns = need

    # normalizações
    if "valor" in out.columns:
        out["valor"] = (
            out["valor"].astype(str)
            .str.replace(r"[^\d,.\-]", "", regex=True)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        # mantém como texto (não converte para float aqui para não perder zeros)
    if "matricula" in out.columns:
        out["matricula"] = out["matricula"].astype(str).str.strip()
    if "nome" in out.columns:
        out["nome"] = out["nome"].astype(str).str.strip()

    out = out.replace(r"^\s*$", pd.NA, regex=True).dropna(how="all").drop_duplicates().reset_index(drop=True)
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

# ========= FGTS =========
# Extração e normalização de tabela de trabalhadores (FGTS)
FGTS_HEADER_HINTS = [
    "nome", "trabalhador", "matricula", "cpf", "vencimento",
    "tipo", "deposito", "base", "remuneracao", "total",
    "valor", "fgts", "na", "guia", "comp", "apuracao", "referencia",
    "juros", "atualiz", "monet", "multa", "categoria", "tomador"
]

def _looks_like_fgts_header(cells: list[str]) -> bool:
    t = " ".join(norm_txt(" ".join(cells))).strip()
    score = sum(1 for h in FGTS_HEADER_HINTS if h in t)
    # considera cabeçalho se tiver pelo menos 4 pistas
    return score >= 4

def _merge_multiline_header(df: pd.DataFrame, max_rows_to_merge: int = 3) -> tuple[list[str], int]:
    """
    Alguns PDFs trazem o cabeçalho quebrado em 2-3 linhas.
    Tenta juntar até 'max_rows_to_merge' linhas para formar o cabeçalho final.
    Retorna (headers, linhas_consumidas).
    """
    df_head = df.head(max_rows_to_merge).fillna("").astype(str)
    candidates = []
    # tenta: só 1ª linha; 1ª+2ª; 1ª+2ª+3ª
    for take in range(1, max_rows_to_merge + 1):
        block = df_head.iloc[:take, :]
        merged = [" ".join([str(x) for x in block.iloc[:, j].tolist() if str(x).strip()]) for j in range(block.shape[1])]
        if _looks_like_fgts_header(merged):
            candidates.append((merged, take))
    if candidates:
        # escolhe o com mais pistas (maior cobertura)
        candidates.sort(key=lambda x: sum(1 for h in FGTS_HEADER_HINTS if h in norm_txt(" ".join(x[0]))), reverse=True)
        return candidates[0]
    # fallback: primeira linha
    return [str(x) for x in df.iloc[0].tolist()], 1

def _keep_only_table_rows(df: pd.DataFrame) -> pd.DataFrame:
    """
    Mantém somente linhas que parecem pertencer à tabela (têm ao menos um CPF),
    removendo linhas de totais, rodapés e metadados.
    """
    df2 = df.copy()
    # remove linhas vazias
    df2 = df2.replace(r"^\s*$", pd.NA, regex=True).dropna(how="all")
    mask_cpf = df2.apply(lambda r: any(CPF_RE.search(str(v)) for v in r if pd.notna(v)), axis=1)
    df2 = df2.loc[mask_cpf]
    return df2.reset_index(drop=True)

def _best_table_candidate(dfs: list[pd.DataFrame]) -> Optional[pd.DataFrame]:
    """
    Entre vários dataframes extraídos, escolhe o que mais parece tabela de FGTS (tem cabeçalho bom e muitas linhas com CPF).
    """
    best = None
    best_score = -1
    for df in dfs:
        d = drop_empty_rows_cols(clean_cells(df))
        if d.empty or d.shape[1] < 4:
            continue
        header, consumed = _merge_multiline_header(d)
        body = d.iloc[consumed:].reset_index(drop=True)
        body.columns = dedupe_columns(header)
        body = _keep_only_table_rows(body)
        if body.empty:
            continue
        # pontua por qtd de linhas e presença de colunas-chave
        cols_norm = [norm_txt(c) for c in body.columns]
        col_score = sum(1 for h in ["nome", "cpf", "matricula", "valor", "fgts", "vencimento"] if any(h in c for c in cols_norm))
        score = body.shape[0] * 2 + col_score
        if score > best_score:
            best_score = score
            best = body
    return best

def process_fgts(pdf_bytes: bytes) -> pd.DataFrame:
    """
    Extrai apenas a Tabela de Trabalhadores do FGTS (sem metadados/rodapés),
    corrigindo rotação (90° anti-horário) e retornando somente cabeçalhos+linhas.
    """
    dfs_all: list[pd.DataFrame] = []

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
            # O PDF veio 90° anti-horário; testamos rotações para garantir
            for angle in ( -90, 90, 0 ):
                try:
                    p = page.rotate(angle)
                except Exception:
                    p = page
                for settings in (lattice, stream):
                    try:
                        tbls = p.extract_tables(dict(settings)) or []
                    except TypeError as e:
                        # remove chaves desconhecidas se houver
                        msg = str(e)
                        if "unexpected keyword argument" in msg:
                            bad = msg.split("'")[1]
                            s2 = {k: v for k, v in settings.items() if k != bad}
                            tbls = p.extract_tables(s2) or []
                        else:
                            continue
                    for t in tbls:
                        df = pd.DataFrame(t)
                        if df.shape[0] and df.shape[1]:
                            dfs_all.append(df)

    if not dfs_all:
        return pd.DataFrame()

    body = _best_table_candidate(dfs_all)
    if body is None or body.empty:
        return pd.DataFrame()

    # Normaliza cabeçalhos: remove quebras e espaços extras
    body.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in body.columns]

    # Remove linhas repetidas e completamente vazias
    body = (body.replace(r"^\s*$", pd.NA, regex=True)
                 .dropna(how="all")
                 .drop_duplicates()
                 .reset_index(drop=True))

    return body

# ========== UI bipartida ==========
left, middle, right = st.columns(3, gap="large")

# --- Lado esquerdo: Oposição ---
with left:
    st.subheader("Planilha de Oposição")
    up_op = st.file_uploader("Upload do Oposições (PDF)", type=["pdf"], key="oposicao")
    if up_op is None:
        st.info("Envie o PDF de Oposições para extrair (nome / cpf / data de entrega).")
    else:
        with st.spinner("Extraindo dados de Oposição..."):
            tables_op = process_oposicao(up_op.read())
        if not tables_op:
            st.error("Não consegui localizar dados (Nome/CPF/Data). Envie 1 página de exemplo para ajustarmos os padrões.")
        else:
            if len(tables_op) > 1:
                st.caption(f"Encontrei {len(tables_op)} quadros válidos. Unifiquei abaixo.")
                df_op = pd.concat(tables_op, ignore_index=True).drop_duplicates().reset_index(drop=True)
            else:
                df_op = tables_op[0]
            st.caption(f"{df_op.shape[0]} linhas • 3 colunas")
            st.dataframe(df_op, use_container_width=True, height=440)
            st.download_button(
                "Baixar XLSX (Oposição)",
                data=df_to_xlsx_bytes(df_op, sheet_name="Oposicao"),
                file_name="oposicao_unificada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# --- Coluna do meio: Extração de FGTS ---
with middle:
    st.subheader("Extração de FGTS")
    up_fgts = st.file_uploader("Upload do FGTS (PDF rotacionado)", type=["pdf"], key="fgts")
    if up_fgts is None:
        st.info("Envie o PDF do FGTS para extrair apenas cabeçalhos e linhas da tabela de trabalhadores.")
    else:
        with st.spinner("Extraindo tabela de FGTS..."):
            df_fgts = process_fgts(up_fgts.read())

        if df_fgts.empty:
            st.error("Não consegui localizar a tabela de trabalhadores. Envie 1 página de exemplo para ajustarmos os padrões.")
        else:
            st.caption(f"{df_fgts.shape[0]} linhas • {df_fgts.shape[1]} colunas (somente cabeçalhos e linhas da tabela)")
            st.dataframe(df_fgts, use_container_width=True, height=440)
            st.download_button(
                "Baixar XLSX (FGTS)",
                data=df_to_xlsx_bytes(df_fgts, sheet_name="FGTS"),
                file_name="fgts_trabalhadores.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


# --- Lado direito: Relação de Associados ---
with right:
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
