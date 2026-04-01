"""
src.py  –  Motore di calcolo per il Computo Metrico Estimativo
==============================================================
Logica pura, zero dipendenze da Streamlit.

Struttura dati principale
--------------------------
VOCE DI COMPUTO (dict):
    id                : int   – identificatore univoco
    wbs               : str   – codice WBS es. "1.2.3"
    categoria         : str   – etichetta categoria
    sottocategoria    : str   – etichetta sottocategoria (opzionale)
    codice            : str   – codice tariffa prezziario
    descrizione       : str   – testo descrittivo
    um                : str   – unità di misura
    prezzo_unitario   : float
    note              : str
    misurazioni       : list[MISURAZIONE]
    quantita_totale   : float  (calcolato)
    importo           : float  (calcolato)

MISURAZIONE (dict):
    descrizione : str
    parti       : float  (parti uguali, default 1)
    lung        : float  (lunghezza)
    larg        : float  (larghezza)
    alt         : float  (altezza / peso)
    quantita    : float  (quantità diretta, usata se lung/larg/alt = 0)

PREZZIARIO (pd.DataFrame):
    colonne: CODICE, DESCRIZIONE, UM, PREZZO, FONTE
"""

from __future__ import annotations

import io
import json
import re
from typing import Any

import pandas as pd


# ══════════════════════════════════════════════════════════════════════════════
# COSTANTI
# ══════════════════════════════════════════════════════════════════════════════

COLONNE_PREZZIARIO = ["CODICE", "DESCRIZIONE", "UM", "PREZZO", "FONTE"]

MISURAZIONE_VUOTA: dict = {
    "descrizione": "",
    "parti":       1.0,
    "lung":        0.0,
    "larg":        0.0,
    "alt":         0.0,
    "quantita":    0.0,
}


# ══════════════════════════════════════════════════════════════════════════════
# UTILITÀ GENERALI
# ══════════════════════════════════════════════════════════════════════════════

def parse_price(val_str: Any) -> float:
    """Converte stringa prezzo → float. Gestisce '1.234,56' e '1234.56'."""
    if val_str is None:
        return 0.0
    s = str(val_str).strip()
    if re.search(r"\d\.\d{3},\d", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    m = re.search(r"\d+\.?\d*", s)
    return float(m.group()) if m else 0.0


def _safe_str(val: Any) -> str:
    s = str(val or "").strip()
    return "" if s in ("nan", "None", "NaN") else s


def _safe_float(val: Any) -> float:
    try:
        return float(str(val).replace(",", "."))
    except (ValueError, TypeError):
        return 0.0


# ══════════════════════════════════════════════════════════════════════════════
# LIBRETTO DELLE MISURE
# ══════════════════════════════════════════════════════════════════════════════

def quantita_misurazione(m: dict) -> float:
    """
    Calcola la quantità di una riga di misurazione.
    - Se lung/larg/alt valorizzati → parti × lung × larg × alt (0 → 1)
    - Altrimenti → parti × quantita (campo diretto)
    """
    parti = float(m.get("parti", 1) or 1)
    lung  = float(m.get("lung",  0) or 0)
    larg  = float(m.get("larg",  0) or 0)
    alt   = float(m.get("alt",   0) or 0)

    if lung or larg or alt:
        l = lung if lung else 1.0
        w = larg if larg else 1.0
        h = alt  if alt  else 1.0
        return round(parti * l * w * h, 4)
    return round(parti * float(m.get("quantita", 0) or 0), 4)


def quantita_totale_voce(voce: dict) -> float:
    """Somma le quantità di tutte le righe di misurazione."""
    misurazioni = voce.get("misurazioni") or []
    if not misurazioni:
        return float(voce.get("quantita", 0) or 0)
    return round(sum(quantita_misurazione(m) for m in misurazioni), 4)


def calcola_importo(voce: dict) -> float:
    """Importo = quantità totale × prezzo unitario."""
    return round(quantita_totale_voce(voce) * float(voce.get("prezzo_unitario", 0) or 0), 2)


def aggiorna_importi(computo: list[dict]) -> None:
    """Ricalcola quantita_totale e importo per tutte le voci. In-place."""
    for v in computo:
        v["quantita_totale"] = quantita_totale_voce(v)
        v["importo"]         = calcola_importo(v)


def nuova_misurazione(**kwargs) -> dict:
    m = dict(MISURAZIONE_VUOTA)
    m.update(kwargs)
    return m


def nuova_voce(next_id: int, **kwargs) -> dict:
    """Crea una voce di computo con struttura completa e valori di default."""
    v: dict = {
        "id":              next_id,
        "wbs":             "",
        "categoria":       "",
        "sottocategoria":  "",
        "codice":          "",
        "descrizione":     "",
        "um":              "",
        "prezzo_unitario": 0.0,
        "note":            "",
        "misurazioni":     [nuova_misurazione()],
        "quantita_totale": 0.0,
        "importo":         0.0,
    }
    v.update(kwargs)
    return v


# ══════════════════════════════════════════════════════════════════════════════
# AGGREGAZIONE E RIEPILOGO WBS
# ══════════════════════════════════════════════════════════════════════════════

def totale_computo(computo: list[dict]) -> float:
    return round(sum(v.get("importo", 0.0) for v in computo), 2)


def riepilogo_wbs(computo: list[dict]) -> pd.DataFrame:
    """Aggrega per WBS/categoria. Restituisce DataFrame ordinato per importo."""
    if not computo:
        return pd.DataFrame(columns=["WBS", "Categoria", "Sottocategoria", "N. voci", "Importo €"])

    rows = [
        {
            "WBS":            v.get("wbs", ""),
            "Categoria":      v.get("categoria") or "— Senza categoria —",
            "Sottocategoria": v.get("sottocategoria", ""),
            "Importo":        v.get("importo", 0.0),
        }
        for v in computo
    ]
    df  = pd.DataFrame(rows)
    agg = (
        df.groupby(["WBS", "Categoria", "Sottocategoria"])["Importo"]
        .agg(N_voci="count", Importo="sum")
        .reset_index()
        .rename(columns={"Importo": "Importo €", "N_voci": "N. voci"})
        .sort_values("Importo €", ascending=False)
        .reset_index(drop=True)
    )
    return agg


def computo_to_dataframe(computo: list[dict]) -> pd.DataFrame:
    """DataFrame piatto per visualizzazione e CSV export."""
    if not computo:
        return pd.DataFrame()
    return pd.DataFrame([
        {
            "N.":             v.get("id", ""),
            "WBS":            v.get("wbs", ""),
            "Categoria":      v.get("categoria", ""),
            "Sottocategoria": v.get("sottocategoria", ""),
            "Codice":         v.get("codice", ""),
            "Descrizione":    str(v.get("descrizione", "") or "")[:120],
            "UM":             v.get("um", ""),
            "Quantità":       v.get("quantita_totale", 0),
            "P.U. €":         v.get("prezzo_unitario", 0),
            "Importo €":      v.get("importo", 0),
        }
        for v in computo
    ])


# ══════════════════════════════════════════════════════════════════════════════
# PARSING PREZZIARIO DA PDF  (PyMuPDF → pdfplumber fallback)
# ══════════════════════════════════════════════════════════════════════════════

_RE_CODICE = re.compile(r"^[A-Za-z]{1,3}\.\d{2,3}(?:\.\d{3})?(?:\.[a-z])?$")


def _parse_rows_from_table(table: list, nome: str) -> list[dict]:
    rows = []
    for row in table:
        if row is None or len(row) < 4:
            continue
        code_cell  = str(row[0] or "").strip()
        desc_cell  = str(row[1] or "").strip()
        um_cell    = str(row[2] or "").strip()
        price_cell = str(row[3] or "").strip()

        if code_cell.upper() in ("CODICE", "COD.", "", "-"):
            continue
        if "descrizione" in code_cell.lower():
            continue

        codes  = [c.strip() for c in code_cell.split("\n")  if c.strip()]
        prices = [p.strip() for p in price_cell.split("\n") if p.strip()]
        ums    = [u.strip() for u in um_cell.split("\n")    if u.strip()]
        descs  = desc_cell.split("\n")

        for i, cod in enumerate(codes):
            if not _RE_CODICE.match(cod):
                continue
            pr = prices[i] if i < len(prices) else (prices[-1] if prices else "0")
            um = ums[i]    if i < len(ums)    else (ums[-1]    if ums    else "")

            desc_chunk = descs[i].strip() if i < len(descs) else ""
            if len(desc_chunk) < 5 and desc_cell:
                desc_chunk = desc_cell[:300].replace("\n", " ")
            desc_clean = re.sub(r"\s+", " ", desc_chunk).strip()

            price_val = parse_price(pr)
            if price_val <= 0:
                continue

            rows.append({"CODICE": cod, "DESCRIZIONE": desc_clean,
                         "UM": um, "PREZZO": price_val, "FONTE": nome})
    return rows


def _extract_rows_from_text(text: str, nome: str) -> list[dict]:
    """Parser regex fallback per PDF a testo libero."""
    rows = []
    pattern = re.compile(
        r"^([A-Za-z]{1,3}\.\d{2,3}(?:\.\d{3})?(?:\.[a-z])?)"
        r"\s+(.+?)\s+([A-Za-z²³/]{1,5})\s+(\d[\d.,]+)",
        re.MULTILINE,
    )
    for m in pattern.finditer(text):
        price_val = parse_price(m.group(4))
        if price_val <= 0:
            continue
        rows.append({
            "CODICE":      m.group(1).strip(),
            "DESCRIZIONE": re.sub(r"\s+", " ", m.group(2)).strip(),
            "UM":          m.group(3).strip(),
            "PREZZO":      price_val,
            "FONTE":       nome,
        })
    return rows


def _finalize_df(rows: list[dict]) -> pd.DataFrame:
    return (
        pd.DataFrame(rows)
        .drop_duplicates(subset=["CODICE"])
        .reset_index(drop=True)
    )


def extract_pdf_prezziario(pdf_bytes: bytes, nome: str) -> pd.DataFrame:
    """
    Estrae voci da PDF prezziario.
    Strategia: PyMuPDF (veloce) → pdfplumber (fallback).
    """
    rows: list[dict] = []

    # ── PyMuPDF ───────────────────────────────────────────────────────────────
    try:
        import fitz
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            for page in doc:
                try:
                    for tab in page.find_tables():
                        rows.extend(_parse_rows_from_table(tab.extract(), nome))
                except AttributeError:
                    rows.extend(_extract_rows_from_text(page.get_text("text"), nome))
        if rows:
            return _finalize_df(rows)
    except ImportError:
        pass

    # ── pdfplumber ────────────────────────────────────────────────────────────
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                for table in (page.extract_tables() or []):
                    rows.extend(_parse_rows_from_table(table, nome))
        if rows:
            return _finalize_df(rows)
    except ImportError as exc:
        raise ImportError(
            "Installa almeno uno tra PyMuPDF e pdfplumber:\n"
            "  pip install pymupdf\n  pip install pdfplumber"
        ) from exc

    return pd.DataFrame(columns=COLONNE_PREZZIARIO)


# ══════════════════════════════════════════════════════════════════════════════
# PARSING PREZZIARIO DA XLSX
# ══════════════════════════════════════════════════════════════════════════════

_COL_KEYWORDS: dict[str, list[str]] = {
    "CODICE":      ["TARIFFA", "CODICE", "COD"],
    "DESCRIZIONE": ["DESCRIZIONE", "DESCR"],
    "UM":          ["UM", "U.M.", "UDM", "UNITÀ", "UNITA"],
    "PREZZO":      ["PREZZO", "COSTO"],
}


def _detect_header_row(df: pd.DataFrame) -> int | None:
    keywords = {"tariffa", "codice", "prezzo", "descrizione"}
    for i, row in df.iterrows():
        row_str = " ".join(str(v).lower() for v in row if pd.notna(v))
        if any(kw in row_str for kw in keywords):
            return int(i)
    return None


def _map_columns(columns: list[str]) -> dict[str, str]:
    col_map: dict[str, str] = {}
    for col in columns:
        for standard, kws in _COL_KEYWORDS.items():
            if standard not in col_map and any(k in col.upper() for k in kws):
                col_map[standard] = col
                break
    return col_map


def extract_xlsx_prezziario(xlsx_bytes: bytes, nome: str) -> pd.DataFrame:
    """Legge XLSX/XLS come prezziario con rilevamento automatico colonne."""
    xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    dfs: list[pd.DataFrame] = []

    for sheet in xls.sheet_names:
        raw        = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=sheet, header=None)
        header_row = _detect_header_row(raw)
        if header_row is None:
            continue

        raw.columns = raw.iloc[header_row]
        raw         = raw.iloc[header_row + 1:].copy()
        raw.columns = [str(c).strip().upper() for c in raw.columns]
        col_map     = _map_columns(list(raw.columns))

        if not all(k in col_map for k in ("CODICE", "DESCRIZIONE", "PREZZO")):
            continue

        wanted  = [k for k in ("CODICE", "DESCRIZIONE", "UM", "PREZZO") if k in col_map]
        sub     = raw[[col_map[k] for k in wanted]].copy()
        sub.columns = wanted
        sub     = sub.dropna(subset=["CODICE", "PREZZO"])
        sub["PREZZO"] = pd.to_numeric(sub["PREZZO"], errors="coerce").fillna(0)
        sub     = sub[sub["PREZZO"] > 0].copy()

        if not sub.empty:
            sub["FONTE"] = nome
            dfs.append(sub)

    if not dfs:
        return pd.DataFrame(columns=COLONNE_PREZZIARIO)
    return (
        pd.concat(dfs, ignore_index=True)
        .drop_duplicates(subset=["CODICE"])
        .reset_index(drop=True)
    )


# ══════════════════════════════════════════════════════════════════════════════
# AGGREGAZIONE PREZZIARI
# ══════════════════════════════════════════════════════════════════════════════

def get_all_voci(prezziari: dict[str, pd.DataFrame]) -> pd.DataFrame:
    dfs = [df for df in prezziari.values() if not df.empty]
    if not dfs:
        return pd.DataFrame(columns=COLONNE_PREZZIARIO)
    return (
        pd.concat(dfs, ignore_index=True)
        .drop_duplicates(subset=["CODICE"])
        .reset_index(drop=True)
    )


def cerca_voce(query: str, df_voci: pd.DataFrame, max_results: int = 30) -> pd.DataFrame:
    if df_voci.empty or not query.strip():
        return df_voci.head(max_results)
    q    = query.strip().upper()
    mask = (
        df_voci["CODICE"].str.upper().str.contains(q, na=False, regex=False)
        | df_voci["DESCRIZIONE"].str.upper().str.contains(q, na=False, regex=False)
    )
    return df_voci[mask].head(max_results)


def lookup_voce_by_codice(codice: str, df_voci: pd.DataFrame) -> dict | None:
    match = df_voci[df_voci["CODICE"].str.upper() == codice.strip().upper()]
    if match.empty:
        return None
    r = match.iloc[0]
    return {
        "codice":      r["CODICE"],
        "descrizione": r["DESCRIZIONE"],
        "um":          r.get("UM", ""),
        "prezzo":      float(r["PREZZO"]),
        "fonte":       r.get("FONTE", ""),
    }


# ══════════════════════════════════════════════════════════════════════════════
# IMPORTAZIONE COMPUTO DA XLSX ESISTENTE
# ══════════════════════════════════════════════════════════════════════════════

def import_computo_from_xlsx(
    xlsx_bytes:   bytes,
    sheet_name:   str,
    header_row:   int,
    col_cat:      int,
    col_sottocat: int,
    col_cod:      int,
    col_desc:     int,
    col_um:       int,
    col_q:        int,
    col_pu:       int,
    start_id:     int = 1,
) -> tuple[list[dict], int]:
    """Importa voci da XLSX esistente con mapping manuale colonne (0-based)."""
    raw        = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=sheet_name, header=None)
    dati       = raw.iloc[header_row + 1:].copy()
    voci:      list[dict] = []
    current_id = start_id

    for _, row in dati.iterrows():
        n        = len(row)
        cat      = _safe_str(row.iloc[col_cat])      if col_cat      < n else ""
        sottocat = _safe_str(row.iloc[col_sottocat]) if col_sottocat < n else ""
        cod      = _safe_str(row.iloc[col_cod])      if col_cod      < n else ""
        desc     = _safe_str(row.iloc[col_desc])     if col_desc     < n else ""
        um       = _safe_str(row.iloc[col_um])       if col_um       < n else ""
        q        = _safe_float(row.iloc[col_q])      if col_q        < n else 0.0
        pu       = _safe_float(row.iloc[col_pu])     if col_pu       < n else 0.0

        if not cod and not desc:
            continue

        mis = nuova_misurazione(descrizione="da XLSX", quantita=q)
        v   = nuova_voce(
            current_id,
            categoria=cat or "—", sottocategoria=sottocat,
            codice=cod, descrizione=desc, um=um, prezzo_unitario=pu,
            note="importato da XLSX", misurazioni=[mis],
        )
        v["quantita_totale"] = quantita_totale_voce(v)
        v["importo"]         = calcola_importo(v)
        voci.append(v)
        current_id += 1

    return voci, current_id


# ══════════════════════════════════════════════════════════════════════════════
# SERIALIZZAZIONE JSON
# ══════════════════════════════════════════════════════════════════════════════

def export_json(computo: list[dict], nomi_prezziari: list[str]) -> str:
    return json.dumps(
        {"computo": computo, "prezziari_caricati": nomi_prezziari},
        ensure_ascii=False, indent=2,
    )


def import_json(data: bytes | str) -> dict:
    try:
        obj = json.loads(data)
    except json.JSONDecodeError as exc:
        raise ValueError(f"File JSON non valido: {exc}") from exc
    if "computo" not in obj:
        raise ValueError("Il file JSON non contiene la chiave 'computo'.")
    return {
        "computo":            obj.get("computo", []),
        "prezziari_caricati": obj.get("prezziari_caricati", []),
    }


# ══════════════════════════════════════════════════════════════════════════════
# ESPORTAZIONE EXCEL  (formule vive + foglio riepilogo WBS)
# ══════════════════════════════════════════════════════════════════════════════

def export_excel(computo: list[dict], titolo_progetto: str = "Computo Metrico Estimativo") -> bytes:
    """
    Genera Excel professionale con:
    - Foglio «Computo»: voci + libretto misure con FORMULE EXCEL vive
    - Foglio «Riepilogo WBS»: totali per categoria
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()

    # ── Palette colori ────────────────────────────────────────────────────────
    C_DARK   = "1A1F36"
    C_MED    = "2D3561"
    C_ACCENT = "4A6CF7"
    C_LIGHT  = "E8ECF4"
    C_WHITE  = "FFFFFF"
    C_DGREEN = "1E7E34"
    C_BGROW  = "F5F7FF"

    def _font(bold=False, color=C_DARK, size=10, italic=False):
        return Font(name="Calibri", bold=bold, color=color, size=size, italic=italic)
    def _fill(color):
        return PatternFill("solid", fgColor=color)

    thin_s   = Side(style="thin",   color="CCCCCC")
    thick_s  = Side(style="medium", color="999999")
    brd_thin = Border(left=thin_s,  right=thin_s,  top=thin_s,  bottom=thin_s)
    brd_thk  = Border(left=thick_s, right=thick_s, top=thick_s, bottom=thick_s)
    al_c     = Alignment(horizontal="center", vertical="center")
    al_r     = Alignment(horizontal="right",  vertical="center")
    al_l     = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    FMT_EUR = '#,##0.00 "€"'
    FMT_N3  = '#,##0.000'
    FMT_N2  = '#,##0.00'

    # ══════════════════════════════════════════════════════════════════════════
    # FOGLIO 1 – COMPUTO METRICO
    # ══════════════════════════════════════════════════════════════════════════
    ws = wb.active
    ws.title        = "Computo Metrico"
    ws.freeze_panes = "A4"

    # Larghezze: A=N.  B=WBS  C=Codice  D=Descrizione  E=UM
    #            F=Parti  G=Lung  H=Larg  I=Alt  J=Qtà  K=P.U.  L=Importo
    widths = [("A",5),("B",8),("C",14),("D",55),("E",7),
              ("F",7),("G",9), ("H",9), ("I",9), ("J",12),("K",14),("L",15)]
    for col_l, w in widths:
        ws.column_dimensions[col_l].width = w

    # Riga 1 – Titolo
    ws.merge_cells("A1:L1")
    c = ws["A1"]
    c.value     = titolo_progetto.upper()
    c.font      = _font(bold=True, color=C_WHITE, size=14)
    c.fill      = _fill(C_DARK)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # Riga 2 – Intestazioni
    hdrs = ["N.", "WBS", "Codice", "Descrizione / Misurazione",
            "UM", "Parti", "Lung.", "Larg.", "Alt./Peso", "Quantità", "P.U. €", "Importo €"]
    for ci, h in enumerate(hdrs, 1):
        cell = ws.cell(row=2, column=ci, value=h)
        cell.font      = _font(bold=True, color=C_WHITE)
        cell.fill      = _fill(C_ACCENT)
        cell.alignment = al_c
        cell.border    = brd_thin
    ws.row_dimensions[2].height = 22

    data_row    = 3
    current_cat = None

    for voce in computo:
        cat = voce.get("categoria", "") or "—"

        # Separatore categoria
        if cat != current_cat:
            current_cat = cat
            ws.merge_cells(f"A{data_row}:L{data_row}")
            label = f"  {cat.upper()}"
            sc = voce.get("sottocategoria", "")
            if sc:
                label += f"  ›  {sc}"
            cell = ws.cell(row=data_row, column=1, value=label)
            cell.font      = _font(bold=True, color=C_DARK)
            cell.fill      = _fill(C_LIGHT)
            cell.border    = brd_thk
            cell.alignment = Alignment(vertical="center")
            ws.row_dimensions[data_row].height = 18
            data_row += 1

        voce_row    = data_row
        misurazioni = voce.get("misurazioni") or []

        # Riga voce – colonne A..E + K
        def _vc(col, val=None, **kw):
            cell = ws.cell(row=voce_row, column=col, value=val)
            cell.border = brd_thin
            return cell

        _vc(1, voce.get("id", "")).alignment     = al_c
        _vc(2, voce.get("wbs", "")).alignment     = al_c
        c3 = _vc(3, voce.get("codice", ""))
        c3.font      = _font(color=C_MED)
        c3.alignment = al_c

        c4 = _vc(4, voce.get("descrizione", ""))
        c4.font      = _font(bold=True)
        c4.alignment = al_l

        _vc(5, voce.get("um", "")).alignment      = al_c

        # F,G,H,I vuoti nella riga voce
        for ci in (6, 7, 8, 9):
            ws.cell(row=voce_row, column=ci).border = brd_thin

        # J (Quantità) e L (Importo) saranno formule → compilate dopo le misurazioni
        ws.cell(row=voce_row, column=10).border        = brd_thin
        ws.cell(row=voce_row, column=10).number_format = FMT_N3
        ws.cell(row=voce_row, column=10).alignment     = al_c

        ck = ws.cell(row=voce_row, column=11, value=voce.get("prezzo_unitario", 0))
        ck.font          = _font(bold=True, color=C_MED)
        ck.number_format = FMT_EUR
        ck.border        = brd_thin
        ck.alignment     = al_r

        cl = ws.cell(row=voce_row, column=12)
        cl.font          = _font(bold=True, color=C_DGREEN)
        cl.number_format = FMT_EUR
        cl.border        = brd_thin
        cl.alignment     = al_r
        cl.fill          = _fill("F0FFF4")

        ws.row_dimensions[voce_row].height = 22
        data_row += 1

        # Righe di misurazione
        mis_start = data_row
        for mis in misurazioni:
            def _mc(col, val=None):
                cell = ws.cell(row=data_row, column=col, value=val)
                cell.border = brd_thin
                return cell

            for ci in (1, 2, 3, 5):
                _mc(ci)

            cm4 = _mc(4, f"    {mis.get('descrizione', '')}")
            cm4.font      = _font(italic=True, color="666666")
            cm4.alignment = al_l

            _mc(6, mis.get("parti", 1) or 1).number_format  = FMT_N2
            _mc(7, mis.get("lung",  0) or 0).number_format  = FMT_N2
            _mc(8, mis.get("larg",  0) or 0).number_format  = FMT_N2
            _mc(9, mis.get("alt",   0) or 0).number_format  = FMT_N2
            for ci in (6, 7, 8, 9):
                ws.cell(row=data_row, column=ci).alignment = al_c

            # Quantità riga → formula Excel
            r  = data_row
            cF, cG, cH, cI = "F", "G", "H", "I"
            has_dim = (mis.get("lung") or 0) or (mis.get("larg") or 0) or (mis.get("alt") or 0)
            if has_dim:
                fq = (f"={cF}{r}"
                      f"*IF({cG}{r}=0,1,{cG}{r})"
                      f"*IF({cH}{r}=0,1,{cH}{r})"
                      f"*IF({cI}{r}=0,1,{cI}{r})")
            else:
                direct_q  = mis.get("quantita", 0) or 0
                fq        = None
                cq_direct = ws.cell(row=data_row, column=10,
                                    value=round(float(mis.get("parti", 1) or 1) * direct_q, 4))
                cq_direct.number_format = FMT_N3
                cq_direct.border        = brd_thin
                cq_direct.alignment     = al_c

            if fq:
                cq = ws.cell(row=data_row, column=10, value=fq)
                cq.number_format = FMT_N3
                cq.border        = brd_thin
                cq.alignment     = al_c

            # K e L vuoti nelle righe misurazioni
            for ci in (11, 12):
                ws.cell(row=data_row, column=ci).border = brd_thin

            ws.row_dimensions[data_row].height = 17
            data_row += 1

        mis_end = data_row - 1

        # Ora compila J e L della riga voce con formule
        j_range  = f"J{mis_start}:J{mis_end}" if misurazioni else None
        j_formula = f"=SUM({j_range})" if j_range else (voce.get("quantita_totale", 0))
        ws.cell(row=voce_row, column=10).value = j_formula
        ws.cell(row=voce_row, column=12).value = f"=J{voce_row}*K{voce_row}"

    # Riga totale
    ws.merge_cells(f"A{data_row}:K{data_row}")
    tl = ws.cell(row=data_row, column=1, value="TOTALE COMPLESSIVO")
    tl.font      = _font(bold=True, color=C_WHITE, size=11)
    tl.fill      = _fill(C_DARK)
    tl.alignment = al_r
    tl.border    = brd_thk

    tv = ws.cell(row=data_row, column=12,
                 value=f"=SUMIF(L3:L{data_row-1},\">0\")")
    tv.font          = _font(bold=True, color=C_WHITE, size=11)
    tv.fill          = _fill(C_DARK)
    tv.number_format = FMT_EUR
    tv.alignment     = al_r
    tv.border        = brd_thk
    ws.row_dimensions[data_row].height = 26

    # ══════════════════════════════════════════════════════════════════════════
    # FOGLIO 2 – RIEPILOGO WBS
    # ══════════════════════════════════════════════════════════════════════════
    ws2    = wb.create_sheet("Riepilogo WBS")
    df_wbs = riepilogo_wbs(computo)
    rhdrs  = ["WBS", "Categoria", "Sottocategoria", "N. voci", "Importo €"]

    ws2.merge_cells("A1:E1")
    th = ws2["A1"]
    th.value     = f"{titolo_progetto.upper()} – RIEPILOGO WBS"
    th.font      = _font(bold=True, color=C_WHITE, size=12)
    th.fill      = _fill(C_DARK)
    th.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 28

    for cl, w in [("A",8),("B",35),("C",30),("D",10),("E",16)]:
        ws2.column_dimensions[cl].width = w

    for ci, h in enumerate(rhdrs, 1):
        cell = ws2.cell(row=2, column=ci, value=h)
        cell.font      = _font(bold=True, color=C_WHITE)
        cell.fill      = _fill(C_ACCENT)
        cell.alignment = al_c
        cell.border    = brd_thin

    for ri, row in df_wbs.iterrows():
        for ci, col_name in enumerate(rhdrs, 1):
            val  = row[col_name]
            cell = ws2.cell(row=ri + 3, column=ci, value=val)
            cell.border = brd_thin
            if col_name == "Importo €":
                cell.number_format = FMT_EUR
                cell.alignment     = al_r
                cell.font          = _font(bold=True, color=C_DGREEN)
            elif col_name == "N. voci":
                cell.alignment = al_c
            else:
                cell.alignment = al_l

    tot_r = len(df_wbs) + 3
    ws2.merge_cells(f"A{tot_r}:D{tot_r}")
    tl2 = ws2.cell(row=tot_r, column=1, value="TOTALE COMPLESSIVO")
    tl2.font = _font(bold=True, color=C_WHITE)
    tl2.fill = _fill(C_DARK); tl2.alignment = al_r
    tv2 = ws2.cell(row=tot_r, column=5, value=totale_computo(computo))
    tv2.font = _font(bold=True, color=C_WHITE)
    tv2.fill = _fill(C_DARK)
    tv2.number_format = FMT_EUR; tv2.alignment = al_r

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# ESPORTAZIONE PDF  (ReportLab)
# ══════════════════════════════════════════════════════════════════════════════

def export_pdf(computo: list[dict], titolo_progetto: str = "Computo Metrico Estimativo") -> bytes:
    """
    Genera PDF del computo in formato A4 landscape con:
    - Intestazione, separatori categoria, libretto misure, totale, footer paginato.
    Richiede: pip install reportlab
    """
    try:
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.lib.units import cm, mm
        from reportlab.platypus import (
            Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
        )
    except ImportError as exc:
        raise ImportError("Installa reportlab: pip install reportlab") from exc

    PAGE_W, PAGE_H = landscape(A4)
    MARGIN = 1.5 * cm

    # Colori ReportLab
    RL_DARK   = colors.HexColor("#1A1F36")
    RL_ACCENT = colors.HexColor("#4A6CF7")
    RL_LIGHT  = colors.HexColor("#E8ECF4")
    RL_BGROW  = colors.HexColor("#F5F7FF")
    RL_GREEN  = colors.HexColor("#1E7E34")
    RL_GRAY   = colors.HexColor("#888888")
    RL_WHITE  = colors.white

    # Stili paragrafo
    def _ps(name, font="Helvetica", size=8, color=colors.black, align=TA_LEFT, leading=10, bold=False, italic=False):
        fn = font
        if bold and italic: fn += "-BoldOblique"
        elif bold:   fn += "-Bold"
        elif italic: fn += "-Oblique"
        return ParagraphStyle(name, fontName=fn, fontSize=size,
                              textColor=color, alignment=align, leading=leading)

    s_hdr   = _ps("hdr",  size=7.5, color=RL_WHITE, align=TA_CENTER, bold=True)
    s_desc  = _ps("desc", size=8,   color=colors.black, align=TA_LEFT, leading=10)
    s_mit   = _ps("mit",  size=7.5, color=RL_GRAY, align=TA_LEFT, leading=9, italic=True)
    s_num   = _ps("num",  size=8,   color=colors.black, align=TA_RIGHT)
    s_numg  = _ps("numg", size=8,   color=RL_GREEN,     align=TA_RIGHT, bold=True)
    s_code  = _ps("code", font="Courier", size=7.5, color=colors.HexColor("#2D3561"), align=TA_LEFT)
    s_cat   = _ps("cat",  size=9,   color=RL_DARK, align=TA_LEFT, bold=True)
    s_tot   = _ps("tot",  size=9,   color=RL_WHITE, align=TA_RIGHT, bold=True)

    def _p(text, style): return Paragraph(str(text), style)
    def _fmt(val, d=2):
        try:    return f"{float(val):,.{d}f}"
        except: return str(val)

    # Larghezze colonne (totale = PAGE_W - 2*MARGIN)
    COL_W = [0.7, 1.2, 2.2, 8.8, 1.0, 1.3, 1.6, 1.6, 1.6, 1.8, 2.3, 2.7]
    COL_W = [w * cm for w in COL_W]

    # Header tabella
    hdr_row = [[_p(h, s_hdr) for h in [
        "N.", "WBS", "Codice", "Descrizione / Misurazione",
        "UM", "Parti", "Lung.", "Larg.", "Alt.", "Quantità", "P.U. €", "Importo €"
    ]]]

    table_data = list(hdr_row)
    ts: list[tuple] = [
        ("BACKGROUND",     (0, 0), (-1, 0),  RL_ACCENT),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [RL_WHITE, RL_BGROW]),
        ("GRID",           (0, 0), (-1, -1), 0.3, colors.HexColor("#CCCCCC")),
        ("VALIGN",         (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",     (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING",  (0, 0), (-1, -1), 2),
    ]

    current_cat = None
    abs_r = 1

    for voce in computo:
        cat = voce.get("categoria", "") or "—"

        if cat != current_cat:
            current_cat = cat
            label = cat.upper()
            sc = voce.get("sottocategoria", "")
            if sc: label += f"  ›  {sc}"
            table_data.append([_p(f"  {label}", s_cat)] + [""] * 11)
            ts += [
                ("BACKGROUND",   (0, abs_r), (-1, abs_r), RL_LIGHT),
                ("SPAN",         (0, abs_r), (-1, abs_r)),
                ("LINEABOVE",    (0, abs_r), (-1, abs_r), 1, RL_ACCENT),
                ("TOPPADDING",   (0, abs_r), (-1, abs_r), 5),
                ("BOTTOMPADDING",(0, abs_r), (-1, abs_r), 5),
            ]
            abs_r += 1

        qt = voce.get("quantita_totale", 0)
        pu = voce.get("prezzo_unitario", 0)
        im = voce.get("importo", 0)
        table_data.append([
            _p(voce.get("id",          ""), s_num),
            _p(voce.get("wbs",         ""), s_num),
            _p(voce.get("codice",      ""), s_code),
            _p(voce.get("descrizione", ""), s_desc),
            _p(voce.get("um",          ""), s_num),
            "", "", "", "",
            _p(_fmt(qt, 3), s_numg),
            _p(_fmt(pu, 2), s_num),
            _p(_fmt(im, 2), s_numg),
        ])
        ts += [
            ("BACKGROUND", (0, abs_r), (-1, abs_r), colors.HexColor("#F0F3FF")),
            ("LINEBELOW",  (0, abs_r), (-1, abs_r), 0.5, RL_ACCENT),
            ("FONTNAME",   (0, abs_r), (-1, abs_r), "Helvetica-Bold"),
        ]
        abs_r += 1

        for mis in (voce.get("misurazioni") or []):
            q_mis = quantita_misurazione(mis)
            table_data.append([
                "", "", "",
                _p(f"    {mis.get('descrizione', '')}", s_mit),
                "",
                _p(_fmt(mis.get("parti", 1), 2), s_num),
                _p(_fmt(mis.get("lung",  0), 2), s_num),
                _p(_fmt(mis.get("larg",  0), 2), s_num),
                _p(_fmt(mis.get("alt",   0), 2), s_num),
                _p(_fmt(q_mis,              3), s_num),
                "", "",
            ])
            abs_r += 1

    # Totale
    total = totale_computo(computo)
    table_data.append(
        [_p("TOTALE COMPLESSIVO", s_tot)] + [""] * 10 +
        [_p(f"€ {total:,.2f}", s_tot)]
    )
    ts += [
        ("BACKGROUND",   (0, abs_r), (-1, abs_r), RL_DARK),
        ("SPAN",         (0, abs_r), (-2, abs_r)),
        ("TOPPADDING",   (0, abs_r), (-1, abs_r), 7),
        ("BOTTOMPADDING",(0, abs_r), (-1, abs_r), 7),
    ]

    main_table = Table(table_data, colWidths=COL_W, repeatRows=1)
    main_table.setStyle(TableStyle(ts))

    buf = io.BytesIO()

    def _on_page(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(RL_GRAY)
        canvas.drawRightString(PAGE_W - MARGIN, 8 * mm, f"Pagina {canvas.getPageNumber()}")
        canvas.drawString(MARGIN, 8 * mm, titolo_progetto)
        canvas.restoreState()

    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4),
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=MARGIN, bottomMargin=1.8 * cm,
    )

    # Titolo
    title_data  = [[_p(titolo_progetto.upper(), _ps("tt", size=14, color=RL_WHITE, align=TA_CENTER, bold=True))]]
    title_table = Table(title_data, colWidths=[PAGE_W - 2 * MARGIN])
    title_table.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), RL_DARK),
        ("TOPPADDING",    (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
    ]))

    doc.build([title_table, Spacer(1, 0.5 * cm), main_table],
              onFirstPage=_on_page, onLaterPages=_on_page)
    buf.seek(0)
    return buf.getvalue()