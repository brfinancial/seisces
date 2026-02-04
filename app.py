from __future__ import annotations

import logging
import os
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ============================ Config / UI =====================================

st.set_page_config(
    page_title="Conciliação WBA x Contabilidade",
    layout="wide"
)

st.title("Conciliação WBA x Contabilidade (Diasdário)")
st.caption(
    "Faça upload dos arquivos da Contabilidade (layout Diário) e do WBA, ajuste os parâmetros e gere o Excel final."
)

# ============================ Logging =========================================

def setup_logging_streamlit(verbosity: int = 1) -> None:
    level = logging.INFO if verbosity == 1 else logging.DEBUG
    fmt = "[%(asctime)s] %(levelname)s - %(message)s"
    logging.basicConfig(format=fmt, level=level, datefmt="%H:%M:%S")


# ============================ Utils texto/validação ===========================

def strip_accents(text: str) -> str:
    if not isinstance(text, str):
        text = "" if text is None else str(text)
    text = unicodedata.normalize("NFKD", text)
    return "".join([c for c in text if not unicodedata.combining(c)])

def normalize_text(text: str) -> str:
    text = strip_accents(text).lower()
    text = re.sub(r"[^\w\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text

def similarity(a: str, b: str) -> float:
    import difflib
    return difflib.SequenceMatcher(None, normalize_text(a), normalize_text(b)).ratio()

def to_date(x) -> Optional[pd.Timestamp]:
    return pd.to_datetime(x, dayfirst=True, errors="coerce")

def to_float(x) -> float:
    try:
        if pd.isna(x):
            return np.nan
        return float(str(x).replace(",", "."))
    except Exception:
        return np.nan


def only_digits(x) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""

    # Se for número inteiro (int / numpy int)
    if isinstance(x, (int, np.integer)):
        return str(int(x))

    # Se for float “inteiro” tipo 10018.0 -> vira 10018
    if isinstance(x, (float, np.floating)):
        if np.isfinite(x) and float(x).is_integer():
            return str(int(x))

    s = str(x).strip()
    # remove casos tipo "10018.0", "10018,0"
    s = s.replace(",", ".")
    s = re.sub(r"\.0+$", "", s)

    return re.sub(r"\D", "", s)

def account_norm(x) -> Optional[str]:
    d = only_digits(x)
    return d if d != "" else None
    
def account_norm(x) -> Optional[str]:
    d = only_digits(x)
    return d if d != "" else None

def cents(v: float) -> Optional[int]:
    if v is None or pd.isna(v):
        return None
    return int(np.rint(float(v) * 100.0))


# ============================ Leitura Contabilidade (Diário) ==================

def _try_header_map(df: pd.DataFrame) -> Dict[str, str]:
    if df is None or df.shape[1] == 0:
        return {}

    cols_norm = {normalize_text(c): c for c in df.columns if isinstance(c, str)}

    def find_first(keys: List[str]) -> Optional[str]:
        for k in keys:
            for cname_norm, cname_real in cols_norm.items():
                if k in cname_norm:
                    return cname_real
        return None

    m = {}
    m["data"]           = find_first(["data"])
    m["conta_debito"]   = find_first(["conta debito", "conta débito", "debito", "débito", "conta origem", "cta deb"])
    m["conta_credito"]  = find_first(["conta credito", "conta crédito", "credito", "crédito", "conta destino", "cta part", "cta c part"])
    m["historico"]      = find_first(["historico", "hist", "histórico"])
    m["valor"]          = find_first(["valor", "vlr"])

    essentials = ["data", "conta_debito", "conta_credito", "historico", "valor"]
    if all(m.get(k) for k in essentials):
        return m
    return {}

def _parse_diario_positional(df: pd.DataFrame) -> pd.DataFrame:
    ncols = df.shape[1]

    def col(i):
        return df.iloc[:, i] if i < ncols else pd.Series([None] * len(df))

    # Layout posicional:
    # A: data | D: conta_debito | F/G: conta_credito | I: historico | N/O: valor
    data_col = col(0).apply(to_date).dt.date
    conta_debito_col = col(3).apply(account_norm)

    conta_credito_col = []
    for f, g in zip(col(5), col(6)):
        cc = account_norm(f) if pd.notna(f) and str(f).strip() != "" else account_norm(g)
        conta_credito_col.append(cc)
    conta_credito_col = pd.Series(conta_credito_col)

    historico_col = col(8).astype(str)

    valor_col = []
    for n, o in zip(col(13), col(14)):
        v = to_float(n) if pd.notna(n) and str(n).strip() != "" else to_float(o)
        valor_col.append(v)
    valor_col = pd.Series(valor_col)

    df2 = pd.DataFrame({
        "data": data_col,
        "conta_debito": conta_debito_col,
        "conta_credito": conta_credito_col,
        "historico": historico_col,
        "valor": valor_col
    })

    df2 = df2.dropna(subset=["data", "conta_debito", "conta_credito", "valor"])
    df2 = df2[df2["valor"] > 0]
    return df2.reset_index(drop=True)

def read_contabilidade_to_standard(path_or_buffer: Union[str, BytesIO]) -> pd.DataFrame:
    logging.info("Lendo Contabilidade (Diário)")
    try:
        df_hdr = pd.read_excel(path_or_buffer, sheet_name=0, header=0, engine="openpyxl")
        map_hdr = _try_header_map(df_hdr)
        if map_hdr:
            logging.debug(f"Mapeamento por cabeçalho detectado: {map_hdr}")
            df = df_hdr.rename(columns={
                map_hdr["data"]: "data",
                map_hdr["conta_debito"]: "conta_debito",
                map_hdr["conta_credito"]: "conta_credito",
                map_hdr["historico"]: "historico",
                map_hdr["valor"]: "valor",
            }).copy()

            df["data"] = pd.to_datetime(df["data"], dayfirst=True, errors="coerce").dt.date
            df["conta_debito"] = df["conta_debito"].apply(account_norm)
            df["conta_credito"] = df["conta_credito"].apply(account_norm)
            df["valor"] = pd.to_numeric(df["valor"], errors="coerce")
            df["historico"] = df["historico"].astype(str)

            df = df.dropna(subset=["data", "conta_debito", "conta_credito", "valor"])
            df = df[df["valor"] > 0]
        else:
            # Sem cabeçalho útil -> posicional
            df_raw = pd.read_excel(path_or_buffer, sheet_name=0, header=None, engine="openpyxl")
            df = _parse_diario_positional(df_raw)
    except Exception:
        logging.error("Falha ao ler o Diário pelo cabeçalho; tentando posicional...", exc_info=True)
        df_raw = pd.read_excel(path_or_buffer, sheet_name=0, header=None, engine="openpyxl")
        df = _parse_diario_positional(df_raw)

    logging.info(f"Contabilidade normalizada (sem deduplicação): {len(df)} linhas.")
    return df


# ============================ Leitura WBA =====================================

def _wba_map_columns(df: pd.DataFrame) -> Dict[str, str]:
    cols = {normalize_text(str(c)): c for c in df.columns}

    def pick(candidates: List[str]) -> Optional[str]:
        for can in candidates:
            for k, v in cols.items():
                if can in k:
                    return v
        return None

    deb  = pick(["deb", "cta.", "conta debito", "cta deb", "cta debito"])
    cred = pick(["cred", "cta.c.part", "cta c part", "conta credito", "cta part", "cta credito"])
    vlr  = pick(["vlr", "valor"])
    dt   = pick(["data"])
    hist = pick(["hist", "descr", "descricao"])

    need = [deb, cred, vlr, dt, hist]
    if any(x is None for x in need):
        raise ValueError(f"Falha ao mapear colunas do WBA. Recebi: {list(df.columns)}")

    return {
        "conta_debito": deb,
        "conta_credito": cred,
        "valor": vlr,
        "data": dt,
        "historico": hist,
    }

def _strip_leading_equals_in_series(s: pd.Series) -> pd.Series:
    s = s.astype("string")
    s = s.str.replace(r'^\s*=\s*', '', regex=True)
    s = s.str.strip()
    return s

def read_wba_to_standard(path_or_buffer: Union[str, BytesIO]) -> pd.DataFrame:
    logging.info("Lendo WBA")
    df_raw = pd.read_excel(path_or_buffer, engine="openpyxl")
    if df_raw.empty:
        return pd.DataFrame(columns=["conta_debito", "conta_credito", "valor", "data", "historico"])

    colmap = _wba_map_columns(df_raw)

    # Limpa '=' no começo do histórico ANTES do rename
    if colmap["historico"] in df_raw.columns:
        df_raw[colmap["historico"]] = _strip_leading_equals_in_series(df_raw[colmap["historico"]])

    df = df_raw.rename(columns=colmap)[["conta_debito", "conta_credito", "valor", "data", "historico"]].copy()

    # Segurança extra
    df["historico"] = _strip_leading_equals_in_series(df["historico"])

    for c in ["conta_debito", "conta_credito"]:
        df[c] = df[c].apply(account_norm)

    df["valor"] = pd.to_numeric(df["valor"], errors="coerce")
    df["data"] = pd.to_datetime(df["data"], dayfirst=True, errors="coerce").dt.date
    df["historico"] = df["historico"].astype(str)

    df = df.dropna(subset=["conta_debito", "conta_credito", "valor", "data"]).copy()
    df = df[df["valor"] > 0]

    logging.info(f"WBA normalizado (sem deduplicação): {len(df)} linhas.")
    return df


# ============================ Matching Engine =================================

@dataclass(frozen=True)
class MatchRow:
    contab_idx: int
    wba_idx: int
    tipo: str
    score: float
    diff_dias: int
    diff_valor: float
    sim_desc: float

def _days_diff(d1, d2) -> int:
    return abs((pd.to_datetime(d1) - pd.to_datetime(d2)).days)

def _acct_set_equal(a1, b1, a2, b2) -> bool:
    return {a1, b1} == {a2, b2}

def build_candidates(
    contab: pd.DataFrame,
    wba: pd.DataFrame,
    janela_dias: int,
    tol_valor_cents: int,
    limiar_desc: float,
) -> Dict[str, List[MatchRow]]:
    logging.info("Gerando candidatos...")
    contab = contab.copy().reset_index(drop=True)
    wba = wba.copy().reset_index(drop=True)
    
    contab["valor_cents"] = contab["valor"].apply(cents)
    wba["valor_cents"] = wba["valor"].apply(cents)
    
    # remove inválidos e RESETA índice posicional
    contab = contab.dropna(subset=["valor_cents"]).reset_index(drop=True)
    wba = wba.dropna(subset=["valor_cents"]).reset_index(drop=True)
    
    contab["valor_cents"] = contab["valor_cents"].astype(int)
    wba["valor_cents"] = wba["valor_cents"].astype(int)
    
    # cria ids fixos para export (opcional, mas ótimo)
    contab["contab_id"] = np.arange(len(contab))
    wba["wba_id"] = np.arange(len(wba))

    def acct_key(a, b) -> Tuple[str, str]:
        return (a, b) if a <= b else (b, a)

    idx_wba_exact: Dict[Tuple, List[int]] = {}
    for row in wba.itertuples(index=False):
        key = (row.data, row.valor_cents, acct_key(row.conta_debito, row.conta_credito))
        idx_wba_exact.setdefault(key, []).append(int(row.wba_idx))

    cand = {
        "exato": [],
        "mesmo_valor_data_perto": [],
        "mesmo_dia_valor_parecido": [],
        "valor_parecido_data_perto": [],
        "fuzzy": [],
    }

    # Para acesso rápido por índice
    wba_by_idx = {int(r.wba_idx): r for r in wba.itertuples(index=False)}

    # 1) EXATO
    for c in contab.itertuples(index=False):
        key = (c.data, c.valor_cents, acct_key(c.conta_debito, c.conta_credito))
        for wi in idx_wba_exact.get(key, []):
            w = wba_by_idx.get(int(wi))
            if w is None:
                continue
            sd = similarity(c.historico, w.historico)
            cand["exato"].append(
                MatchRow(int(c.contab_idx), int(wi), "exato", score=3.0 + sd,
                         diff_dias=0, diff_valor=0.0, sim_desc=sd)
            )

    idx_wba_by_val: Dict[int, List[int]] = {}
    for row in wba.itertuples(index=False):
        idx_wba_by_val.setdefault(int(row.valor_cents), []).append(int(row.wba_idx))

    idx_wba_by_date: Dict[object, List[int]] = {}
    for row in wba.itertuples(index=False):
        idx_wba_by_date.setdefault(row.data, []).append(int(row.wba_idx))

    # 2) valor igual, data perto
    for c in contab.itertuples(index=False):
        for wi in idx_wba_by_val.get(int(c.valor_cents), []):
            w = wba_by_idx.get(int(wi))
            if w is None:
                continue
            if not _acct_set_equal(c.conta_debito, c.conta_credito, w.conta_debito, w.conta_credito):
                continue
            dd = _days_diff(c.data, w.data)
            if 0 < dd <= janela_dias:
                sd = similarity(c.historico, w.historico)
                dv = abs(float(c.valor) - float(w.valor))
                score = (1.0 / (1 + dd)) + (1.0 / (1 + dv)) + sd
                cand["mesmo_valor_data_perto"].append(
                    MatchRow(int(c.contab_idx), int(wi), "mesmo_valor_data_perto", score, dd, dv, sd)
                )

    # 3) mesma data, valor parecido
    for c in contab.itertuples(index=False):
        for wi in idx_wba_by_date.get(c.data, []):
            w = wba_by_idx.get(int(wi))
            if w is None:
                continue
            if not _acct_set_equal(c.conta_debito, c.conta_credito, w.conta_debito, w.conta_credito):
                continue
            dv_cents = abs(int(c.valor_cents) - int(w.valor_cents))
            if 0 < dv_cents <= tol_valor_cents:
                dd = 0
                dv = abs(float(c.valor) - float(w.valor))
                sd = similarity(c.historico, w.historico)
                score = (1.0 / (1 + dv)) + 1.5 + sd
                cand["mesmo_dia_valor_parecido"].append(
                    MatchRow(int(c.contab_idx), int(wi), "mesmo_dia_valor_parecido", score, dd, dv, sd)
                )

    # 4) valor parecido + data perto (atenção: O(n*m) — pode ser pesado em arquivos grandes)
    wba_rows = list(wba.itertuples(index=False))
    for c in contab.itertuples(index=False):
        for w in wba_rows:
            if not _acct_set_equal(c.conta_debito, c.conta_credito, w.conta_debito, w.conta_credito):
                continue
            dd = _days_diff(c.data, w.data)
            if 0 < dd <= janela_dias:
                dv_cents = abs(int(c.valor_cents) - int(w.valor_cents))
                if 0 < dv_cents <= tol_valor_cents:
                    dv = abs(float(c.valor) - float(w.valor))
                    score = (1.0 / (1 + dd)) + (1.0 / (1 + dv))
                    cand["valor_parecido_data_perto"].append(
                        MatchRow(int(c.contab_idx), int(w.wba_idx), "valor_parecido_data_perto",
                                 score, dd, dv, 0.0)
                    )

    # 5) fuzzy
    for c in contab.itertuples(index=False):
        for w in wba_rows:
            if not _acct_set_equal(c.conta_debito, c.conta_credito, w.conta_debito, w.conta_credito):
                continue
            dd = _days_diff(c.data, w.data)
            if dd <= janela_dias:
                dv_cents = abs(int(c.valor_cents) - int(w.valor_cents))
                if dv_cents <= tol_valor_cents:
                    sd = similarity(c.historico, w.historico)
                    if sd >= limiar_desc:
                        dv = abs(float(c.valor) - float(w.valor))
                        score = (1.0 / (1 + dd)) + (1.0 / (1 + dv)) + sd
                        cand["fuzzy"].append(
                            MatchRow(int(c.contab_idx), int(w.wba_idx), "fuzzy",
                                     score, dd, dv, sd)
                        )

    for k in cand:
        cand[k].sort(key=lambda r: r.score, reverse=True)

    return cand

def greedy_resolve(
    contab: pd.DataFrame,
    wba: pd.DataFrame,
    cand: Dict[str, List[MatchRow]],
) -> Dict[str, List[MatchRow]]:
    logging.info("Resolvendo conflitos (1-para-1)...")
    assigned_c = set()
    assigned_w = set()
    result = {k: [] for k in cand.keys()}

    priority = ["exato", "mesmo_valor_data_perto",
                "mesmo_dia_valor_parecido", "valor_parecido_data_perto", "fuzzy"]
    for tier in priority:
        for m in cand[tier]:
            if (m.contab_idx not in assigned_c) and (m.wba_idx not in assigned_w):
                result[tier].append(m)
                assigned_c.add(m.contab_idx)
                assigned_w.add(m.wba_idx)

    all_c = set(range(len(contab)))
    all_w = set(range(len(wba)))

    matched_c = set().union(*[set(x.contab_idx for x in result[t]) for t in result])
    matched_w = set().union(*[set(x.wba_idx for x in result[t]) for t in result])

    result["so_contabilidade"] = [
        MatchRow(ci, -1, "so_contabilidade", 0.0, 0, 0.0, 0.0)
        for ci in sorted(all_c - matched_c)
    ]
    result["so_wba"] = [
        MatchRow(-1, wi, "so_wba", 0.0, 0, 0.0, 0.0)
        for wi in sorted(all_w - matched_w)
    ]
    return result


# ============================ Exportação Excel =================================

def _format_dual_area_sheet(ws):
    blue_fill  = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    bold_font  = Font(bold=True)
    center_v   = Alignment(vertical="center")

    for col_idx in range(1, 7):   # A..F (Contab)
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = blue_fill
        cell.font = bold_font
        cell.alignment = center_v

    for col_idx in range(7, 12+1):   # G..L (WBA)
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = green_fill
        cell.font = bold_font
        cell.alignment = center_v

    thick = Side(style="thick", color="000000")
    for row in range(1, ws.max_row + 1):
        ws.cell(row=row, column=6).border = Border(right=thick)  # F
        ws.cell(row=row, column=7).border = Border(left=thick)   # G

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

    widths = {
        1: 14, 2: 16, 3: 16, 4: 12, 5: 14, 6: 40,
        7: 14, 8: 16, 9: 16, 10: 12, 11: 14, 12: 40
    }
    for col_idx, w in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = w

    date_fmt = "DD/MM/YY"
    currency_fmt = 'R$ #,##0.00'
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=4).number_format = date_fmt
        ws.cell(row=r, column=5).number_format = currency_fmt
        ws.cell(row=r, column=10).number_format = date_fmt
        ws.cell(row=r, column=11).number_format = currency_fmt

def _format_simple_sheet(ws):
    bold = Font(bold=True)
    if ws.max_row >= 1:
        for c in ws[1]:
            c.font = bold
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

    # Autoajuste de coluna (limite de 60)
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            v = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(v))
        ws.column_dimensions[col_letter].width = min(max(12, max_len + 2), 60)

def _write_sheet(writer: pd.ExcelWriter, df: pd.DataFrame, name: str, dual: bool = False):
    if df is None:
        df = pd.DataFrame()

    # Se vier vazio sem colunas, ainda escreve um cabeçalho padrão
    if df.shape[1] == 0:
        df = pd.DataFrame(columns=[
            "ID CONTAB", "CONTA DEB CONTAB", "CONTA CRED CONTAB", "DATA CONTAB", "VALOR CONTAB", "HIST CONTAB",
            "ID WBA", "CONTA DEB WBA", "CONTA CRED WBA", "DATA WBA", "VALOR WBA", "HIST WBA"
        ])

    df.to_excel(writer, sheet_name=name, index=False)
    ws = writer.book[name]
    if dual:
        _format_dual_area_sheet(ws)
    else:
        _format_simple_sheet(ws)
    return True

def to_dual_records(contab: pd.DataFrame, wba: pd.DataFrame, matches: List[MatchRow]) -> pd.DataFrame:
    cols = [
        "ID CONTAB", "CONTA DEB CONTAB", "CONTA CRED CONTAB", "DATA CONTAB", "VALOR CONTAB", "HIST CONTAB",
        "ID WBA", "CONTA DEB WBA", "CONTA CRED WBA", "DATA WBA", "VALOR WBA", "HIST WBA"
    ]
    if matches is None or len(matches) == 0:
        return pd.DataFrame(columns=cols)

    rows = []
    for m in matches:
        c = contab.iloc[m.contab_idx] if m.contab_idx != -1 and len(contab) > 0 else None
        w = wba.iloc[m.wba_idx]       if m.wba_idx != -1 and len(wba) > 0 else None
        rows.append({
            "ID CONTAB":         (int(m.contab_idx) if c is not None else None),
            "CONTA DEB CONTAB":  (c["conta_debito"] if c is not None else None),
            "CONTA CRED CONTAB": (c["conta_credito"] if c is not None else None),
            "DATA CONTAB":       (c["data"] if c is not None else None),
            "VALOR CONTAB":      (round(float(c["valor"]), 2) if c is not None else None),
            "HIST CONTAB":       (c["historico"] if c is not None else None),

            "ID WBA":            (int(m.wba_idx) if w is not None else None),
            "CONTA DEB WBA":     (w["conta_debito"] if w is not None else None),
            "CONTA CRED WBA":    (w["conta_credito"] if w is not None else None),
            "DATA WBA":          (w["data"] if w is not None else None),
            "VALOR WBA":         (round(float(w["valor"]), 2) if w is not None else None),
            "HIST WBA":          (w["historico"] if w is not None else None),
        })

    df = pd.DataFrame(rows)
    for c in cols:
        if c not in df.columns:
            df[c] = pd.Series(dtype="object")
    return df[cols]

def compute_data_valor_conta_divergente(so_contab: pd.DataFrame, so_wba: pd.DataFrame) -> pd.DataFrame:
    base_cols = [
        "DATA CONTAB", "VALOR CONTAB_R",
        "ID CONTAB", "CONTA DEB CONTAB", "CONTA CRED CONTAB", "HIST CONTAB",
        "ID WBA", "CONTA DEB WBA", "CONTA CRED WBA", "DATA WBA", "VALOR WBA", "HIST WBA",
        "DEBITO_IGUAL", "CREDITO_IGUAL"
    ]

    min_sc = {"DATA CONTAB", "VALOR CONTAB", "ID CONTAB", "CONTA DEB CONTAB", "CONTA CRED CONTAB"}
    min_sw = {"DATA WBA", "VALOR WBA", "ID WBA", "CONTA DEB WBA", "CONTA CRED WBA"}

    if (so_contab is None) or (so_wba is None):
        return pd.DataFrame(columns=base_cols)
    if not min_sc.issubset(set(so_contab.columns)) or not min_sw.issubset(set(so_wba.columns)):
        return pd.DataFrame(columns=base_cols)
    if so_contab.empty or so_wba.empty:
        return pd.DataFrame(columns=base_cols)

    sc = so_contab.copy()
    sw = so_wba.copy()

    sc["DATA CONTAB"] = pd.to_datetime(sc["DATA CONTAB"], errors="coerce").dt.date
    sw["DATA WBA"] = pd.to_datetime(sw["DATA WBA"], errors="coerce").dt.date
    sc["VALOR CONTAB_R"] = pd.to_numeric(sc["VALOR CONTAB"], errors="coerce").round(2)
    sw["VALOR WBA_R"] = pd.to_numeric(sw["VALOR WBA"], errors="coerce").round(2)

    sc = sc.dropna(subset=["DATA CONTAB", "VALOR CONTAB_R"])
    sw = sw.dropna(subset=["DATA WBA", "VALOR WBA_R"])
    if sc.empty or sw.empty:
        return pd.DataFrame(columns=base_cols)

    sc["key"] = list(zip(sc["DATA CONTAB"], sc["VALOR CONTAB_R"]))
    sw["key"] = list(zip(sw["DATA WBA"], sw["VALOR WBA_R"]))

    key_c = sc.groupby("key").size().rename("n_c")
    key_w = sw.groupby("key").size().rename("n_w")
    if key_c.empty or key_w.empty:
        return pd.DataFrame(columns=base_cols)

    unique_keys = (
        key_c.to_frame()
        .join(key_w, how="inner")
        .query("n_c == 1 and n_w == 1")
        .index
    )
    if len(unique_keys) == 0:
        return pd.DataFrame(columns=base_cols)

    sc_sub = (
        sc.set_index("key").loc[list(unique_keys)].reset_index()[[
            "key", "ID CONTAB", "CONTA DEB CONTAB", "CONTA CRED CONTAB", "DATA CONTAB", "VALOR CONTAB", "HIST CONTAB", "VALOR CONTAB_R"
        ]]
    )
    sw_sub = (
        sw.set_index("key").loc[list(unique_keys)].reset_index()[[
            "key", "ID WBA", "CONTA DEB WBA", "CONTA CRED WBA", "DATA WBA", "VALOR WBA", "HIST WBA", "VALOR WBA_R"
        ]]
    )

    pairs = sc_sub.merge(sw_sub, on="key", how="inner")
    if pairs.empty:
        return pd.DataFrame(columns=base_cols)

    pairs["DEBITO_IGUAL"] = (pairs["CONTA DEB CONTAB"] == pairs["CONTA DEB WBA"])
    pairs["CREDITO_IGUAL"] = (pairs["CONTA CRED CONTAB"] == pairs["CONTA CRED WBA"])

    conta_diff = pairs[~(pairs["DEBITO_IGUAL"] & pairs["CREDITO_IGUAL"])].copy()
    if conta_diff.empty:
        return pd.DataFrame(columns=base_cols)

    conta_diff = conta_diff[[
        "DATA CONTAB", "VALOR CONTAB_R",
        "ID CONTAB", "CONTA DEB CONTAB", "CONTA CRED CONTAB", "HIST CONTAB",
        "ID WBA", "CONTA DEB WBA", "CONTA CRED WBA", "DATA WBA", "VALOR WBA", "HIST WBA",
        "DEBITO_IGUAL", "CREDITO_IGUAL"
    ]].sort_values(["DATA CONTAB", "VALOR CONTAB_R", "ID CONTAB"]).reset_index(drop=True)

    return conta_diff

def export_excel_bytes(
    base_contab: pd.DataFrame,
    base_wba: pd.DataFrame,
    resolved: Dict[str, List[MatchRow]],
) -> bytes:
    logging.info("Exportando Excel (BytesIO)...")

    title_map = [
        ("exato", "Acertos_Exatos"),
        ("mesmo_valor_data_perto", "Mesmo_Valor_Data_Perto"),
        ("mesmo_dia_valor_parecido", "Mesmo_Dia_Valor_Parecido"),
        ("valor_parecido_data_perto", "Valor_Parecido_Data_Perto"),
        ("fuzzy", "Fuzzy_Valor+Data+Desc"),
        ("so_contabilidade", "So_Contabilidade"),
        ("so_wba", "So_WBA"),
    ]

    dfs_dual: Dict[str, pd.DataFrame] = {}
    for key, title in title_map:
        dfs_dual[title] = to_dual_records(base_contab, base_wba, resolved.get(key, []))

    # DataValor_ContaDiff (com base nas SOBRAS)
    try:
        conta_diff_df = compute_data_valor_conta_divergente(
            so_contab=dfs_dual.get("So_Contabilidade"),
            so_wba=dfs_dual.get("So_WBA")
        )
    except Exception:
        logging.exception("Falha ao calcular DataValor_ContaDiff; seguindo com aba vazia.")
        conta_diff_df = pd.DataFrame(columns=[
            "DATA CONTAB", "VALOR CONTAB_R", "ID CONTAB", "CONTA DEB CONTAB", "CONTA CRED CONTAB", "HIST CONTAB",
            "ID WBA", "CONTA DEB WBA", "CONTA CRED WBA", "DATA WBA", "VALOR WBA", "HIST WBA",
            "DEBITO_IGUAL", "CREDITO_IGUAL"
        ])

    # Remove IDs migrados de So_Contabilidade / So_WBA
    try:
        if conta_diff_df is not None and not conta_diff_df.empty:
            ids_sc = set(pd.to_numeric(conta_diff_df["ID CONTAB"], errors="coerce").dropna().astype(int))
            ids_sw = set(pd.to_numeric(conta_diff_df["ID WBA"], errors="coerce").dropna().astype(int))

            if "So_Contabilidade" in dfs_dual:
                sc_df = dfs_dual["So_Contabilidade"].copy()
                if not sc_df.empty and "ID CONTAB" in sc_df.columns and ids_sc:
                    idx_series = pd.to_numeric(sc_df["ID CONTAB"], errors="coerce")
                    dfs_dual["So_Contabilidade"] = sc_df[~idx_series.isin(ids_sc)].reset_index(drop=True)

            if "So_WBA" in dfs_dual:
                sw_df = dfs_dual["So_WBA"].copy()
                if not sw_df.empty and "ID WBA" in sw_df.columns and ids_sw:
                    idx_series = pd.to_numeric(sw_df["ID WBA"], errors="coerce")
                    dfs_dual["So_WBA"] = sw_df[~idx_series.isin(ids_sw)].reset_index(drop=True)

    except Exception:
        logging.exception("Falha ao filtrar 'sobras' após DataValor_ContaDiff; mantendo originais.")

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        # Bases
        _write_sheet(writer, base_contab, "Base_Contabilidade", dual=False)
        _write_sheet(writer, base_wba, "Base_WBA", dual=False)

        # Camadas
        for _, title in title_map:
            _write_sheet(writer, dfs_dual.get(title, pd.DataFrame()), title, dual=True)

        # Aba nova
        if conta_diff_df is None:
            conta_diff_df = pd.DataFrame(columns=[
                "DATA CONTAB", "VALOR CONTAB_R",
                "ID CONTAB", "CONTA DEB CONTAB", "CONTA CRED CONTAB", "HIST CONTAB",
                "ID WBA", "CONTA DEB WBA", "CONTA CRED WBA", "DATA WBA", "VALOR WBA", "HIST WBA",
                "DEBITO_IGUAL", "CREDITO_IGUAL"
            ])
        conta_diff_df.to_excel(writer, index=False, sheet_name="DataValor_ContaDiff")
        _format_simple_sheet(writer.book["DataValor_ContaDiff"])

    buffer.seek(0)
    return buffer.getvalue()


# ============================ Streamlit UI ====================================

with st.sidebar:
    st.header("Parâmetros")
    verbosity = st.selectbox("Log", options=["INFO", "DEBUG"], index=0)
    janela_dias = st.number_input("Janela de dias (data perto)", min_value=0, value=7, step=1)
    tol_valor = st.number_input("Tolerância de valor (R$)", min_value=0.0, value=1.00, step=0.10, format="%.2f")
    limiar_desc = st.number_input("Limiar de similaridade (fuzzy)", min_value=0.0, max_value=1.0, value=0.62, step=0.01)

setup_logging_streamlit(verbosity=1 if verbosity == "INFO" else 2)

colA, colB = st.columns(2)
with colA:
    contab_file = st.file_uploader("📄 Contabilidade (Diário) - Excel", type=["xlsx", "xls"])
with colB:
    wba_file = st.file_uploader("📄 WBA - Excel", type=["xlsx", "xls"])

st.divider()

run = st.button("🚀 Gerar conciliação", type="primary", use_container_width=True)

def _to_buffer(uploaded_file) -> BytesIO:
    return BytesIO(uploaded_file.getvalue())

if run:
    if contab_file is None or wba_file is None:
        st.error("Envie os dois arquivos (Contabilidade e WBA).")
        st.stop()

    with st.spinner("Processando conciliação..."):
        try:
            contab_buf = _to_buffer(contab_file)
            wba_buf = _to_buffer(wba_file)

            base_contab = read_contabilidade_to_standard(contab_buf)
            base_wba = read_wba_to_standard(wba_buf)

            tol_cents = int(round(float(tol_valor) * 100))
            cand = build_candidates(
                base_contab, base_wba,
                janela_dias=int(janela_dias),
                tol_valor_cents=tol_cents,
                limiar_desc=float(limiar_desc),
            )
            resolved = greedy_resolve(base_contab, base_wba, cand)

            # Resumo
            resumo = {
                "exato": len(resolved.get("exato", [])),
                "mesmo_valor_data_perto": len(resolved.get("mesmo_valor_data_perto", [])),
                "mesmo_dia_valor_parecido": len(resolved.get("mesmo_dia_valor_parecido", [])),
                "valor_parecido_data_perto": len(resolved.get("valor_parecido_data_perto", [])),
                "fuzzy": len(resolved.get("fuzzy", [])),
                "so_contabilidade": len(resolved.get("so_contabilidade", [])),
                "so_wba": len(resolved.get("so_wba", [])),
            }

            st.success("Conciliação concluída!")
            st.subheader("Resumo")
            st.json(resumo)

            xlsx_bytes = export_excel_bytes(base_contab, base_wba, resolved)
            ts = datetime.now().strftime("%Y%m%d-%H%M%S")
            st.download_button(
                "⬇️ Baixar Excel (.xlsx)",
                data=xlsx_bytes,
                file_name=f"reconciliacao_WBA_vs_Contabilidade_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            # Prévia opcional
            with st.expander("Prévia (primeiras linhas da Base_Contabilidade e Base_WBA)"):
                st.write("Base_Contabilidade")
                st.dataframe(base_contab.head(30), use_container_width=True)
                st.write("Base_WBA")
                st.dataframe(base_wba.head(30), use_container_width=True)

        except Exception as e:
            st.exception(e)
