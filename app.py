"""
app.py  –  Computo Metrico Estimativo v2.1
==========================================
Tab 1  –  Editor Computo
    · Colonna sinistra  → Prezziario con ricerca istantanea
    · Colonna destra    → Tabella di Sintesi DINAMICA:
                          ogni riga è un expander che contiene le misurazioni
                          di dettaglio inline (commento/n°/lung/larg/alt/totale)
                          Tutto si aggiorna senza cambiare tab.

Tab 2  –  Computo Stampabile
    · Tabella HTML professionale con voci + misurazioni di dettaglio
    · Pulsante stampa (window.print) e download rapido CSV/Excel/PDF

Tab 3  –  Riepilogo WBS / Lotti
Tab 4  –  Importa XLSX
Tab 5  –  Esporta
"""

import copy
import io
from pathlib import Path

import pandas as pd
import streamlit as st

from prezziario_cache import PrezziarioCache, dataframe_info, md5_bytes
from src import (
    COLONNE_PREZZIARIO,
    TIPI_VOCE,
    TIPI_RIGA,
    aggiorna_importi,
    assegna_progressive,
    calcola_importo,
    cerca_voce,
    computo_to_dataframe,
    export_excel,
    export_json,
    export_pdf,
    extract_pdf_prezziario,
    extract_xlsx_prezziario,
    get_all_voci,
    import_computo_from_xlsx,
    import_json,
    nuova_misurazione,
    nuova_voce,
    quantita_misurazione,
    quantita_totale_voce,
    riepilogo_wbs,
    totale_computo,
    _normalizza_voce,
)

# ──────────────────────────────────────────────────────────────────────────────
# CONFIGURAZIONE PAGINA
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Computo Metrico",
    page_icon="📐",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────────────────────────────────────
# CSS GLOBALE
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

.main-header {
    background: linear-gradient(135deg, #1A1F36 0%, #2D3561 50%, #4A6CF7 100%);
    color:white; padding:1.2rem 2rem; border-radius:12px;
    margin-bottom:1rem; display:flex; align-items:center; gap:1rem;
}
.main-header h1 { margin:0; font-size:1.6rem; font-weight:700; letter-spacing:-.5px; }
.main-header p  { margin:0; opacity:.75; font-size:.82rem; }

.metric-card {
    background:white; border:1px solid #E8ECF4; border-radius:10px;
    padding:1rem 1.25rem; text-align:center;
    box-shadow:0 2px 8px rgba(26,31,54,.06);
}
.metric-card .value { font-size:1.5rem; font-weight:700; color:#2D3561; font-family:'DM Mono',monospace; }
.metric-card .label { font-size:.72rem; color:#999; text-transform:uppercase; letter-spacing:.6px; margin-top:.2rem; }

.tag-cod { display:inline-block; background:#E8F0FE; color:#2D3561; padding:1px 8px; border-radius:12px; font-size:.75rem; font-weight:600; font-family:'DM Mono',monospace; }

.total-row {
    background:linear-gradient(90deg,#1A1F36,#2D3561); color:white;
    padding:.9rem 1.4rem; border-radius:10px; font-family:'DM Mono',monospace;
    font-weight:600; font-size:1rem; display:flex; justify-content:space-between; margin-top:.8rem;
}
.sub-total-row {
    padding:.5rem 1rem; border-radius:8px; font-family:'DM Mono',monospace;
    font-weight:600; font-size:.85rem; display:flex; justify-content:space-between; margin-top:.3rem;
}

.sec-label { font-size:.68rem; text-transform:uppercase; letter-spacing:1.5px; color:#aaa; font-weight:600; margin:.8rem 0 .4rem 0; }
[data-testid="stSidebar"] { background:#F0F2F8; }
.stTabs [data-baseweb="tab-list"] { gap:.4rem; background:#F0F2F8; padding:.4rem; border-radius:10px; }
.stTabs [data-baseweb="tab"]      { border-radius:8px; font-weight:500; font-size:.88rem; }
.stButton > button { border-radius:8px; font-weight:600; }
[data-testid="stDataFrameResizable"] thead th { background:#1A1F36 !important; color:white !important; font-size:.78rem; }

/* Computo stampabile */
.print-btn {
    display:inline-block; background:#1A1F36; color:white;
    padding:.5rem 1.4rem; border-radius:8px; font-weight:600;
    cursor:pointer; font-size:.9rem; margin-bottom:.8rem; border:none;
    font-family:'DM Sans',sans-serif;
}
table.computo-tbl { width:100%; border-collapse:collapse; font-family:'DM Sans',sans-serif; font-size:8.5pt; }
table.computo-tbl th { background:#1A1F36; color:white; padding:5px 6px; text-align:center; border:1px solid #333; font-size:8pt; }
table.computo-tbl td { border:1px solid #ddd; padding:3px 5px; vertical-align:middle; }
table.computo-tbl .cat-row  td { background:#E8ECF4; font-weight:700; color:#1A1F36; padding:6px 8px; }
table.computo-tbl .voce-row td { background:#F0F3FF; font-weight:600; border-top:2px solid #4A6CF7; }
table.computo-tbl .mis-row  td { background:white;  color:#555; font-size:8pt; }
table.computo-tbl .mis-sub  td { background:#FFF0F3; color:#C62828; font-size:8pt; }
table.computo-tbl .mis-rif  td { background:#F0FFF4; color:#2E7D32; font-size:8pt; }
table.computo-tbl .total-r  td { background:#1A1F36; color:white; font-weight:700; font-size:9pt; }
table.computo-tbl .num  { text-align:right;  font-family:'DM Mono',monospace; }
table.computo-tbl .ctr  { text-align:center; }
table.computo-tbl .ind  { padding-left:18px; }

@media print {
    .no-print, [data-testid="stSidebar"], header, .main-header, .metric-card { display:none !important; }
}
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# HELPER FUNCTIONS
# ──────────────────────────────────────────────────────────────────────────────

def _build_mis_df(voce: dict, computo: list) -> pd.DataFrame:
    """Costruisce il DataFrame per il data_editor del libretto misure."""
    misi = voce.get("misurazioni") or []
    empty_cols = ["Tipo","Rif.#","Commento","N°","Lung.","Larg.","Alt.","Q.dir.","Totale"]
    if not misi:
        return pd.DataFrame(columns=empty_cols)
    return pd.DataFrame([
        {
            "Tipo":    m.get("tipo_riga", "standard"),
            "Rif.#":  m.get("rif_voce_id"),
            "Commento":m.get("commento", ""),
            "N°":      float(m.get("simili", 1)           or 1),
            "Lung.":   float(m.get("lung", 0)             or 0),
            "Larg.":   float(m.get("larg", 0)             or 0),
            "Alt.":    float(m.get("alt", 0)              or 0),
            "Q.dir.":  float(m.get("quantita_diretta", 0) or 0),
            "Totale":  round(quantita_misurazione(m, computo), 4),
        }
        for m in misi
    ])


def _apply_mis_df(voce: dict, edited: pd.DataFrame, computo: list) -> None:
    """Applica il DataFrame editato alle misurazioni della voce (sincronizzazione immediata)."""
    nuove = []
    for _, r in edited.iterrows():
        tipo  = str(r.get("Tipo", "standard") or "standard")
        r_raw = r.get("Rif.#")
        rif   = (int(r_raw) if r_raw is not None
                 and not (isinstance(r_raw, float) and pd.isna(r_raw))
                 else None)
        nuove.append(nuova_misurazione(
            commento         = str(r.get("Commento",  "") or ""),
            simili           = float(r.get("N°",   1)    or 1),
            lung             = float(r.get("Lung.", 0)   or 0),
            larg             = float(r.get("Larg.", 0)   or 0),
            alt              = float(r.get("Alt.",  0)   or 0),
            quantita_diretta = float(r.get("Q.dir.",0)   or 0),
            tipo_riga        = tipo,
            rif_voce_id      = rif,
        ))
    voce["misurazioni"]     = nuove
    voce["quantita_totale"] = quantita_totale_voce(voce, computo)
    voce["importo"]         = calcola_importo(voce, computo)


def _label_voce(v: dict) -> str:
    art = v.get("articolo") or v.get("codice", "—")
    brv = str(v.get("descrizione_breve") or v.get("descrizione", ""))[:35]
    prg = v.get("progressiva", "")
    return f"#{v['id']} [{prg}]  {art}  ·  {brv}"


def _tipo_icon(tipo: str) -> str:
    return {"sovrapprezzo_pct": "📈 ", "riferimento": "🔗 "}.get(tipo, "")


def _fmt_n(val, d=3) -> str:
    try:    return f"{float(val):,.{d}f}" if float(val) != 0 else ""
    except: return ""


def _fmt_e(val, d=2) -> str:
    try:    return f"€ {float(val):,.{d}f}" if float(val) != 0 else ""
    except: return ""


def _td(val, cls="", colspan=1) -> str:
    span = f' colspan="{colspan}"' if colspan > 1 else ""
    return f'<td class="{cls}"{span}>{val}</td>'


# ──────────────────────────────────────────────────────────────────────────────
# CACHE SINGLETON
# ──────────────────────────────────────────────────────────────────────────────
@st.cache_resource
def _get_cache() -> PrezziarioCache:
    return PrezziarioCache()

_cache = _get_cache()


# ──────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ──────────────────────────────────────────────────────────────────────────────
def _init() -> None:
    for k, v in {
        "prezziari":       {},
        "computo":         [],
        "next_id":         1,
        "titolo_progetto": "Computo Metrico Estimativo",
        "cache_caricata":  False,
    }.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init()

if not st.session_state.cache_caricata:
    cached = _cache.carica_tutti()
    if cached:
        st.session_state.prezziari.update(cached)
    st.session_state.cache_caricata = True


def _nuovo_id() -> int:
    nid = st.session_state.next_id
    st.session_state.next_id += 1
    return nid


def _voce_by_id(vid: int) -> dict | None:
    return next((v for v in st.session_state.computo if v["id"] == vid), None)


# ──────────────────────────────────────────────────────────────────────────────
# CALCOLI GLOBALI (ad ogni rerun, prima del rendering)
# ──────────────────────────────────────────────────────────────────────────────
for _v in st.session_state.computo:
    _normalizza_voce(_v)
aggiorna_importi(st.session_state.computo)
assegna_progressive(st.session_state.computo)

total_imp  = totale_computo(st.session_state.computo)
n_voci     = len(st.session_state.computo)
n_prez     = len(st.session_state.prezziari)
tot_v_prez = sum(len(d) for d in st.session_state.prezziari.values())


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:

    st.markdown('<p class="sec-label">📐 Progetto</p>', unsafe_allow_html=True)
    st.session_state.titolo_progetto = st.text_input(
        "Titolo progetto", value=st.session_state.titolo_progetto,
        placeholder="es. Ponte SR69 – Consolidamento soletta",
        label_visibility="collapsed",
    )

    st.markdown("---")
    st.markdown('<p class="sec-label">📚 Prezziari attivi</p>', unsafe_allow_html=True)

    if not st.session_state.prezziari:
        st.caption("Nessun prezziario caricato.")
    else:
        for nome, df in list(st.session_state.prezziari.items()):
            in_cache = nome in _cache
            c1, c2, c3 = st.columns([4, 1, 1])
            c1.markdown(f"**{nome}**  \n<small>{len(df):,} voci  {'💾' if in_cache else '⚠️'}</small>",
                        unsafe_allow_html=True)
            if not in_cache:
                if c2.button("💾", key=f"save_cache_{nome}"):
                    _cache.salva(nome, df); st.rerun()
            else:
                c2.markdown("✓")
            if c3.button("✕", key=f"del_prez_{nome}"):
                del st.session_state.prezziari[nome]; st.rerun()

    st.markdown("---")
    st.markdown('<p class="sec-label">➕ Gestione prezziari</p>', unsafe_allow_html=True)
    t_nuovo, t_cache_tab, t_parquet = st.tabs(["📄 Nuovo", "🗄️ Cache", "📦 Parquet"])

    with t_nuovo:
        st.caption("Carica PDF o XLSX. Salvato automaticamente in cache.")
        up_file   = st.file_uploader("PDF o XLSX", type=["pdf","xlsx","xls"],
                                      key="sidebar_up", label_visibility="collapsed")
        nome_prez = st.text_input("Nome prezziario", placeholder="es. NC-MP 2025", key="nome_prez")
        if st.button("📥 Analizza e carica", use_container_width=True, type="primary"):
            if not nome_prez:   st.warning("Inserisci un nome.")
            elif not up_file:   st.warning("Seleziona un file.")
            else:
                raw      = up_file.read()
                hash_src = md5_bytes(raw)
                if nome_prez in _cache:
                    meta = next((m for m in _cache.lista() if m["nome"] == nome_prez), {})
                    if meta.get("hash_sorgente") == hash_src:
                        st.session_state.prezziari[nome_prez] = _cache.carica(nome_prez)
                        st.rerun()
                with st.spinner(f"Analisi {up_file.name}…"):
                    try:
                        df_prez = (extract_pdf_prezziario(raw, nome_prez)
                                   if Path(up_file.name).suffix.lower() == ".pdf"
                                   else extract_xlsx_prezziario(raw, nome_prez))
                        if df_prez.empty:
                            st.warning("Nessuna voce estratta.")
                        else:
                            _cache.salva(nome_prez, df_prez, hash_sorgente=hash_src)
                            st.session_state.prezziari[nome_prez] = df_prez
                            st.rerun()
                    except ImportError as e:
                        st.error(str(e))

    with t_cache_tab:
        voci_cache = _cache.lista()
        if not voci_cache:
            st.caption("Cache vuota.")
        else:
            for meta in voci_cache:
                nome_c = meta["nome"]
                attivo = nome_c in st.session_state.prezziari
                with st.expander(f"**{nome_c}** — {meta['n_voci']:,} voci {'🟢' if attivo else '⚪'}"):
                    st.caption(f"Aggiornato: {meta['data_aggiornamento']}")
                    ca, cb, cc = st.columns(3)
                    if not attivo:
                        if ca.button("▶ Carica", key=f"load_{nome_c}", use_container_width=True, type="primary"):
                            st.session_state.prezziari[nome_c] = _cache.carica(nome_c); st.rerun()
                    else:
                        if ca.button("⏹ Scarica", key=f"unload_{nome_c}", use_container_width=True):
                            del st.session_state.prezziari[nome_c]; st.rerun()
                    try:
                        cb.download_button("📊 XLSX", data=_cache.esporta_xlsx(nome_c),
                            file_name=f"{nome_c}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_xlsx_{nome_c}", use_container_width=True)
                    except Exception: pass
                    try:
                        cc.download_button("📦 .parquet", data=_cache.esporta_parquet(nome_c),
                            file_name=f"{nome_c}.parquet", mime="application/octet-stream",
                            key=f"dl_pq_{nome_c}", use_container_width=True)
                    except Exception: pass
                    if st.button("🗑️ Elimina", key=f"del_cache_{nome_c}", use_container_width=True):
                        _cache.elimina(nome_c)
                        st.session_state.prezziari.pop(nome_c, None); st.rerun()

    with t_parquet:
        pq_up   = st.file_uploader(".parquet prezziario", type=["parquet"],
                                    key="pq_up", label_visibility="collapsed")
        pq_nome = st.text_input("Nome da assegnare", placeholder="es. NC-MP 2025", key="pq_nome")
        if st.button("📦 Importa parquet", use_container_width=True, type="primary"):
            if not pq_nome: st.warning("Inserisci un nome.")
            elif not pq_up: st.warning("Seleziona un file.")
            else:
                try:
                    st.session_state.prezziari[pq_nome] = _cache.importa_parquet(pq_nome, pq_up.read())
                    st.rerun()
                except Exception as e:
                    st.error(f"Errore: {e}")

    st.markdown("---")
    st.markdown('<p class="sec-label">💾 Gestione progetto</p>', unsafe_allow_html=True)
    if st.session_state.computo:
        json_str = export_json(st.session_state.computo, list(st.session_state.prezziari.keys()))
        st.download_button("💾 Salva progetto JSON", data=json_str.encode(),
                           file_name="computo.json", mime="application/json",
                           use_container_width=True, key="scarica_csv_computo_unico")
    json_up = st.file_uploader("Riapri progetto JSON", type=["json"], key="json_up")
    if json_up and st.button("📂 Carica progetto", use_container_width=True):
        try:
            res = import_json(json_up.read())
            st.session_state.computo = res["computo"]
            aggiorna_importi(st.session_state.computo); st.rerun()
        except ValueError as e:
            st.error(str(e))
    if st.button("🗑️ Azzera computo", use_container_width=True):
        st.session_state.computo = []; st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# HEADER + METRICHE
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="main-header">
    <div style="font-size:2.4rem">📐</div>
    <div>
        <h1>{st.session_state.titolo_progetto}</h1>
        <p>Computo Metrico Estimativo · Libretto Misure Inline · Sovrapprezzi · Riferimenti Voce</p>
    </div>
</div>
""", unsafe_allow_html=True)

m_cols = st.columns(4)
for col, (val, label) in zip(m_cols, [
    (str(n_prez),          "Prezziari"),
    (f"{tot_v_prez:,}",    "Voci disponibili"),
    (str(n_voci),          "Voci nel computo"),
    (f"€ {total_imp:,.2f}","Totale complessivo"),
]):
    col.markdown(
        f'<div class="metric-card"><div class="value">{val}</div>'
        f'<div class="label">{label}</div></div>',
        unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# TABS
# ──────────────────────────────────────────────────────────────────────────────
tab_editor, tab_stampa, tab_riepilogo, tab_import, tab_export = st.tabs([
    "📋 Editor Computo",
    "🖨️ Computo Stampabile",
    "📊 Riepilogo WBS",
    "📥 Importa XLSX",
    "📤 Esporta",
])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1  –  EDITOR COMPUTO
#   Colonna sx: Prezziario | Colonna dx: Tabella Sintesi Dinamica
# ══════════════════════════════════════════════════════════════════════════════
with tab_editor:
    col_prez, col_comp = st.columns([1, 1], gap="medium")

    # ─────────────────────────────────────────────────────────────────────────
    # COLONNA SINISTRA – PREZZIARIO
    # ─────────────────────────────────────────────────────────────────────────
    with col_prez:
        st.markdown("#### 🔍 Prezziario")
        all_voci_df = get_all_voci(st.session_state.prezziari)

        if all_voci_df.empty:
            st.info("👈 Carica un prezziario dalla barra laterale per iniziare.")
        else:
            query = st.text_input("Cerca", placeholder="es.  calcestruzzo  scavo  T.10",
                                   key="prez_search")
            cf1, cf2 = st.columns([3, 1])
            fonti      = ["Tutti"] + list(all_voci_df["FONTE"].unique())
            fonte_filt = cf1.selectbox("Fonte", fonti, key="fonte_filt", label_visibility="collapsed")
            max_r      = cf2.selectbox("Max", [30, 50, 100], key="max_r", label_visibility="collapsed")
            df_filt    = all_voci_df if fonte_filt == "Tutti" else all_voci_df[all_voci_df["FONTE"] == fonte_filt]
            risultati  = cerca_voce(query, df_filt, max_results=int(max_r))
            st.markdown(f"<small>**{len(risultati)}** risultati</small>", unsafe_allow_html=True)

            for _, riga in risultati.iterrows():
                r_cod = riga["CODICE"]
                r_desc = str(riga["DESCRIZIONE"])[:58]
                r_um   = riga.get("UM", "")
                r_pr   = float(riga["PREZZO"])

                cc, cd, cp, cb = st.columns([2, 5, 2, 1])
                cc.markdown(f"<span class='tag-cod'>{r_cod}</span>", unsafe_allow_html=True)
                cd.markdown(f"<small>{r_desc}</small>", unsafe_allow_html=True)
                cp.markdown(f"<small>**€ {r_pr:,.2f}**/{r_um}</small>", unsafe_allow_html=True)

                if cb.button("➕", key=f"add_{r_cod}", help=f"Aggiungi {r_cod} al computo"):
                    mis = nuova_misurazione(commento="Misura 1", quantita_diretta=1.0)
                    v   = nuova_voce(
                        _nuovo_id(),
                        articolo          = r_cod,
                        descrizione       = riga["DESCRIZIONE"],
                        descrizione_breve = str(riga["DESCRIZIONE"])[:40],
                        um                = r_um,
                        prezzo_unitario   = r_pr,
                        misurazioni       = [mis],
                    )
                    st.session_state.computo.append(v)
                    aggiorna_importi(st.session_state.computo)
                    assegna_progressive(st.session_state.computo)
                    st.rerun()
                st.divider()

    # ─────────────────────────────────────────────────────────────────────────
    # COLONNA DESTRA – TABELLA DI SINTESI DINAMICA
    # ─────────────────────────────────────────────────────────────────────────
    with col_comp:
        st.markdown("#### 📋 Tabella di Sintesi")

        if not st.session_state.computo:
            st.info("Il computo è vuoto. Clicca ➕ accanto a una voce del prezziario per iniziare.")
        else:
            # ── Toolbar ────────────────────────────────────────────────────────
            tb1, tb2 = st.columns([5, 2])
            tb1.caption(f"**{n_voci} voci**  ·  Espandi ogni riga per modificare le misurazioni inline.")
            if tb2.button("➕ Aggiungi voce vuota", use_container_width=True):
                st.session_state.computo.append(
                    nuova_voce(_nuovo_id(), articolo="NUOVO",
                               descrizione="Nuova voce",
                               descrizione_breve="Nuova voce",
                               categoria="Varie")
                )
                aggiorna_importi(st.session_state.computo)
                assegna_progressive(st.session_state.computo)
                st.rerun()

            st.markdown(" ")

            # ── Loop principale: ogni voce = expander ─────────────────────────
            voci_da_eliminare: list[int] = []

            for voce in st.session_state.computo:
                vid  = voce["id"]
                tipo = voce.get("tipo", "standard")
                prg  = voce.get("progressiva", "")
                art  = voce.get("articolo") or voce.get("codice", "—")
                brv  = str(voce.get("descrizione_breve") or voce.get("descrizione",""))[:30]
                um   = voce.get("um", "")
                qt   = voce.get("quantita_totale", 0)
                imp  = voce.get("importo", 0)
                icon = _tipo_icon(tipo)

                # Etichetta dell'expander: dati-chiave a colpo d'occhio
                lbl = (
                    f"{icon}**[{prg}]**  `{art}`  ·  {brv}  "
                    f"|  **{qt:.3f}** {um}  |  **€ {imp:,.2f}**"
                )

                with st.expander(lbl, expanded=False):

                    # ── A) CAMPI VOCE ────────────────────────────────────────
                    # Riga 1: Breve · Descrizione completa
                    a1, a2 = st.columns([2, 3])
                    new_brv  = a1.text_input("Descrizione breve",
                                              value=voce.get("descrizione_breve",""),
                                              key=f"brv_{vid}", max_chars=50)
                    new_desc = a2.text_input("Descrizione completa",
                                             value=voce.get("descrizione",""),
                                             key=f"desc_{vid}")

                    # Riga 2: Categoria · WBS · Lotto · UM
                    b1, b2, b3, b4 = st.columns([3, 2, 2, 1])
                    new_cat   = b1.text_input("Categoria",    value=voce.get("categoria",""),   key=f"cat_{vid}")
                    new_wbs   = b2.text_input("WBS",          value=voce.get("wbs",""),         key=f"wbs_{vid}")
                    new_lotto = b3.text_input("Lotto",        value=voce.get("lotto",""),       key=f"lotto_{vid}")
                    new_um    = b4.text_input("UM",           value=voce.get("um",""),          key=f"um_{vid}")

                    # Riga 3: P.U. · Tipo voce · Sottocategoria · Articolo
                    c1, c2, c3, c4 = st.columns([2, 2, 2, 2])
                    new_pu   = c1.number_input("P.U. (€)",
                                               value=float(voce.get("prezzo_unitario",0) or 0),
                                               min_value=0.0, format="%.4f", key=f"pu_{vid}")
                    new_tipo = c2.selectbox("Tipo voce", options=list(TIPI_VOCE),
                                            index=list(TIPI_VOCE).index(tipo),
                                            key=f"tipo_{vid}")
                    new_scat = c3.text_input("Sottocategoria", value=voce.get("sottocategoria",""), key=f"scat_{vid}")
                    new_art  = c4.text_input("Articolo / Tariffa", value=art, key=f"art_{vid}")

                    # Applica campi voce immediatamente
                    voce.update({
                        "descrizione_breve": new_brv,
                        "descrizione":       new_desc,
                        "categoria":         new_cat,
                        "wbs":               new_wbs,
                        "lotto":             new_lotto,
                        "um":                new_um,
                        "prezzo_unitario":   new_pu,
                        "tipo":              new_tipo,
                        "sottocategoria":    new_scat,
                        "articolo":          new_art,
                    })

                    # ── B) CONFIGURAZIONE TIPO SPECIALE ─────────────────────
                    if new_tipo in ("sovrapprezzo_pct", "riferimento"):
                        st.markdown("---")
                        altre = [v for v in st.session_state.computo if v["id"] != vid]
                        opts  = [None] + [v["id"] for v in altre]
                        fmt   = lambda i: ("— nessuno —" if i is None else _label_voce(_voce_by_id(i) or {}))
                        cur   = voce.get("rif_voce_id")
                        idx   = opts.index(cur) if cur in opts else 0

                        if new_tipo == "sovrapprezzo_pct":
                            st.markdown("**📈 Sovrapprezzo %**")
                            s1, s2 = st.columns([4, 2])
                            sel_rif = s1.selectbox("Voce base", opts, format_func=fmt, index=idx, key=f"sovr_rif_{vid}")
                            new_pct = s2.number_input("% sovr.", 0.0, 999.9,
                                                      float(voce.get("sovrapprezzo_pct") or 0),
                                                      step=0.5, format="%.2f", key=f"sovr_pct_{vid}")
                            voce["rif_voce_id"]      = sel_rif
                            voce["sovrapprezzo_pct"] = new_pct
                            if sel_rif and new_pct:
                                padre = _voce_by_id(sel_rif)
                                if padre:
                                    base = calcola_importo(padre, st.session_state.computo)
                                    st.success(f"Anteprima: {new_pct}% di € {base:,.2f} = **€ {base*new_pct/100:.2f}**")
                        else:
                            st.markdown("**🔗 Riferimento voce**")
                            sel_rif = st.selectbox("Voce padre (eredita quantità)", opts,
                                                   format_func=fmt, index=idx, key=f"rif_id_{vid}")
                            voce["rif_voce_id"] = sel_rif
                            if sel_rif:
                                padre = _voce_by_id(sel_rif)
                                if padre:
                                    qt_p = quantita_totale_voce(padre, st.session_state.computo)
                                    st.success(f"Quantità ereditata: **{qt_p:.4f} {padre.get('um','')}**")

                    # ── C) MISURAZIONI INLINE ────────────────────────────────
                    st.markdown("---")

                    if new_tipo == "riferimento" and voce.get("rif_voce_id"):
                        st.caption("ℹ️ Voce di tipo *riferimento*: la quantità è ereditata dalla voce padre.")
                    else:
                        st.markdown(
                            "<small>📏 **Libretto misure**  ·  "
                            "`standard` = simili×L×W×H  |  "
                            "`sottrazione` = valore negativo  |  "
                            "`riferimento_voce` = copia qta da voce #</small>",
                            unsafe_allow_html=True,
                        )

                        df_mis = _build_mis_df(voce, st.session_state.computo)
                        edited = st.data_editor(
                            df_mis,
                            use_container_width=True,
                            num_rows="dynamic",
                            hide_index=True,
                            disabled=["Totale"],
                            column_config={
                                "Tipo":    st.column_config.SelectboxColumn(
                                               "Tipo", options=list(TIPI_RIGA), width="small"),
                                "Rif.#":  st.column_config.NumberColumn(
                                               "Rif.#", width=55, min_value=0,
                                               help="ID voce da referenziare"),
                                "Commento":st.column_config.TextColumn(
                                               "Commento / Descrizione", width="large"),
                                "N°":     st.column_config.NumberColumn(
                                               "N° simili", format="%.2f", min_value=0.0, width=65),
                                "Lung.":  st.column_config.NumberColumn(
                                               "Lung.", format="%.3f", min_value=0.0, width=70),
                                "Larg.":  st.column_config.NumberColumn(
                                               "Larg.", format="%.3f", min_value=0.0, width=70),
                                "Alt.":   st.column_config.NumberColumn(
                                               "Alt.",  format="%.3f", min_value=0.0, width=70),
                                "Q.dir.": st.column_config.NumberColumn(
                                               "Q.dir.", format="%.4f", min_value=0.0, width=70,
                                               help="Quantità diretta (se senza dimensioni)"),
                                "Totale": st.column_config.NumberColumn(
                                               "Totale", format="%.4f", width=80),
                            },
                            key=f"mis_{vid}",
                        )
                        # Sync immediata: applica le misurazioni editate alla voce
                        _apply_mis_df(voce, edited, st.session_state.computo)

                    # ── D) TOTALI VOCE ───────────────────────────────────────
                    voce["quantita_totale"] = quantita_totale_voce(voce, st.session_state.computo)
                    voce["importo"]         = calcola_importo(voce, st.session_state.computo)

                    d1, d2 = st.columns(2)
                    d1.markdown(
                        f'<div class="sub-total-row" style="background:linear-gradient(90deg,#2D3561,#4A6CF7)">'
                        f'<span>Quantità totale</span>'
                        f'<span>{voce["quantita_totale"]:.4f} {voce.get("um","")}</span>'
                        f'</div>', unsafe_allow_html=True)
                    d2.markdown(
                        f'<div class="sub-total-row" style="background:linear-gradient(90deg,#1E7E34,#28A745)">'
                        f'<span>Importo voce</span>'
                        f'<span>€ {voce["importo"]:,.2f}</span>'
                        f'</div>', unsafe_allow_html=True)

                    # ── E) AZIONI VOCE ───────────────────────────────────────
                    st.markdown(" ")
                    e1, e2, e3, e4 = st.columns(4)

                    if e1.button("⬆️ Su",     key=f"up_{vid}",  use_container_width=True):
                        idx = next((i for i, v in enumerate(st.session_state.computo) if v["id"] == vid), None)
                        if idx and idx > 0:
                            c = st.session_state.computo
                            c[idx-1], c[idx] = c[idx], c[idx-1]
                        st.rerun()

                    if e2.button("⬇️ Giù",    key=f"dn_{vid}",  use_container_width=True):
                        idx = next((i for i, v in enumerate(st.session_state.computo) if v["id"] == vid), None)
                        c   = st.session_state.computo
                        if idx is not None and idx < len(c) - 1:
                            c[idx], c[idx+1] = c[idx+1], c[idx]
                        st.rerun()

                    if e3.button("📋 Duplica", key=f"dup_{vid}", use_container_width=True):
                        new_v = copy.deepcopy(voce)
                        new_v["id"] = _nuovo_id()
                        new_v["descrizione_breve"] = (new_v.get("descrizione_breve","") + " (copia)")[:50]
                        idx = next((i for i, v in enumerate(st.session_state.computo) if v["id"] == vid), None)
                        st.session_state.computo.insert(idx + 1, new_v)
                        st.rerun()

                    if e4.button("🗑️ Elimina", key=f"del_{vid}", use_container_width=True, type="secondary"):
                        voci_da_eliminare.append(vid)

            # Esegui le eliminazioni accumulate
            if voci_da_eliminare:
                st.session_state.computo = [
                    v for v in st.session_state.computo if v["id"] not in voci_da_eliminare
                ]
                st.rerun()

            # ── TOTALE COMPLESSIVO ─────────────────────────────────────────────
            aggiorna_importi(st.session_state.computo)
            st.markdown(
                f'<div class="total-row">'
                f'<span>TOTALE COMPLESSIVO  ({n_voci} voci)</span>'
                f'<span>€ {totale_computo(st.session_state.computo):,.2f}</span>'
                f'</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2  –  COMPUTO STAMPABILE
# ══════════════════════════════════════════════════════════════════════════════
with tab_stampa:
    st.markdown("### 🖨️ Computo Metrico – Vista Stampabile")

    if not st.session_state.computo:
        st.info("Il computo è vuoto. Aggiungi delle voci nella tab *Editor Computo*.")
    else:
        titolo = st.session_state.titolo_progetto

        st.markdown(
            '<button class="print-btn no-print" onclick="window.print()">🖨️ Stampa / Salva PDF</button>',
            unsafe_allow_html=True,
        )
        st.caption("Suggerimento: dal browser usa *Stampa → Salva come PDF* con Grafica di sfondo attiva.")

        # ── Genera tabella HTML professionale ─────────────────────────────────
        rows = []

        # Intestazione titolo
        rows.append(f"""
        <tr><td colspan="11" style="background:#1A1F36;color:white;font-size:11pt;
            font-weight:700;padding:10px 12px;text-align:center;border:none;">
            {titolo.upper()}
        </td></tr>""")

        # Intestazioni colonne
        rows.append("""
        <tr>
            <th style="width:4%">Prg.</th>
            <th style="width:8%">Articolo</th>
            <th style="width:27%">Descrizione / Commento</th>
            <th style="width:4%">UM</th>
            <th style="width:4%">N°</th>
            <th style="width:6%">Lung.</th>
            <th style="width:6%">Larg.</th>
            <th style="width:6%">Alt.</th>
            <th style="width:8%">Quantità</th>
            <th style="width:8%">P.U. €</th>
            <th style="width:9%">Importo €</th>
        </tr>""")

        current_cat = None
        aggiorna_importi(st.session_state.computo)

        for voce in st.session_state.computo:
            cat  = voce.get("categoria","") or "—"
            tipo = voce.get("tipo","standard")

            # Separatore categoria
            if cat != current_cat:
                current_cat = cat
                scat  = voce.get("sottocategoria","")
                label = f"{cat.upper()}{'  ›  '+scat if scat else ''}"
                rows.append(f'<tr class="cat-row"><td colspan="11">&nbsp;&nbsp;{label}</td></tr>')

            prg  = voce.get("progressiva","")
            art  = voce.get("articolo") or voce.get("codice","")
            desc = voce.get("descrizione","")
            um   = voce.get("um","")
            qt   = voce.get("quantita_totale",0)
            pu   = voce.get("prezzo_unitario",0)
            imp  = voce.get("importo",0)

            tipo_note = ""
            if tipo == "sovrapprezzo_pct":
                pct = voce.get("sovrapprezzo_pct",0)
                tipo_note = f' <span style="color:#E65100;font-size:7pt">[SOVR. {pct}%]</span>'
            elif tipo == "riferimento":
                tipo_note = f' <span style="color:#2E7D32;font-size:7pt">[RIF.→#{voce.get("rif_voce_id","")}]</span>'

            rows.append(f"""
            <tr class="voce-row">
                {_td(prg, "ctr")}
                {_td(f'<code style="font-size:7.5pt">{art}</code>')}
                {_td(f'<strong>{desc}</strong>{tipo_note}')}
                {_td(um, "ctr")}
                {_td("","ctr")} {_td("","num")} {_td("","num")} {_td("","num")}
                {_td(_fmt_n(qt,3), "num")}
                {_td(_fmt_e(pu,4), "num")}
                {_td(f'<strong>{_fmt_e(imp)}</strong>', "num")}
            </tr>""")

            for mis in (voce.get("misurazioni") or []):
                tipo_riga = mis.get("tipo_riga","standard")
                q_mis     = quantita_misurazione(mis, st.session_state.computo)
                commento  = mis.get("commento","")
                simili    = mis.get("simili",1)  or 1
                lung      = mis.get("lung",0)    or 0
                larg      = mis.get("larg",0)    or 0
                alt       = mis.get("alt",0)     or 0

                row_cls = {"sottrazione":"mis-sub","riferimento_voce":"mis-rif"}.get(tipo_riga,"mis-row")
                pfx = {"sottrazione":"(–) ","riferimento_voce":f"[→#{mis.get('rif_voce_id','')}] "}.get(tipo_riga,"")

                rows.append(f"""
                <tr class="{row_cls}">
                    {_td("","ctr")}
                    {_td("","ctr")}
                    {_td(f'<span class="ind">{pfx}{commento}</span>')}
                    {_td("","ctr")}
                    {_td(_fmt_n(simili,2), "num")}
                    {_td(_fmt_n(lung,3),  "num")}
                    {_td(_fmt_n(larg,3),  "num")}
                    {_td(_fmt_n(alt,3),   "num")}
                    {_td(_fmt_n(q_mis,4), "num")}
                    {_td("","num")}
                    {_td("","num")}
                </tr>""")

        # Totale
        total = totale_computo(st.session_state.computo)
        rows.append(f"""
        <tr class="total-r">
            {_td("TOTALE COMPLESSIVO", colspan=10)}
            {_td(f"<strong>€ {total:,.2f}</strong>", "num")}
        </tr>""")

        st.markdown(
            f'<div style="overflow-x:auto"><table class="computo-tbl">{"".join(rows)}</table></div>',
            unsafe_allow_html=True,
        )

        # ── Download rapido ────────────────────────────────────────────────────
        st.markdown("---")
        dl1, dl2, dl3 = st.columns(3)

        with dl1:
            try:
                xl = export_excel(st.session_state.computo, titolo)
                st.download_button("📊 Scarica Excel", data=xl, file_name="computo_metrico.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True, type="primary", key="1")
            except Exception as e:
                st.error(f"Excel: {e}")
        with dl2:
            try:
                pdf = export_pdf(st.session_state.computo, titolo)
                st.download_button("📄 Scarica PDF", data=pdf, file_name="computo_metrico.pdf",
                                   mime="application/pdf",
                                   use_container_width=True, type="primary", key="2")
            except ImportError:
                st.warning("⚠️ `pip install reportlab`")
            except Exception as e:
                st.error(f"PDF: {e}")
        with dl3:
            csv = computo_to_dataframe(st.session_state.computo).to_csv(index=False, sep=";", decimal=",")
            st.download_button("📃 Scarica CSV", data=csv, file_name="computo_metrico.csv",
                               mime="text/csv", use_container_width=True, key="3")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3  –  RIEPILOGO WBS
# ══════════════════════════════════════════════════════════════════════════════
with tab_riepilogo:
    st.markdown("### 📊 Riepilogo WBS / Lotti / Categorie")

    if not st.session_state.computo:
        st.info("Nessuna voce nel computo.")
    else:
        df_wbs = riepilogo_wbs(st.session_state.computo)
        lotti  = ["Tutti"] + sorted([x for x in df_wbs["Lotto"].unique() if x])
        lotto_filt  = st.selectbox("Filtra per Lotto", lotti, key="lotto_filt")
        df_wbs_show = df_wbs if lotto_filt == "Tutti" else df_wbs[df_wbs["Lotto"] == lotto_filt]

        cg, ct = st.columns([3, 2])
        with cg:
            chart_df = df_wbs_show.set_index("Categoria")[["Importo €"]].sort_values("Importo €")
            st.bar_chart(chart_df, height=320, color="#4A6CF7")
        with ct:
            d = df_wbs_show.copy()
            d["Importo €"] = d["Importo €"].map(lambda x: f"€ {x:,.2f}")
            st.dataframe(d, use_container_width=True, hide_index=True, height=320)

        st.markdown("---")
        st.markdown("#### Dettaglio completo")
        df_det = computo_to_dataframe(st.session_state.computo).copy()
        df_det["Importo €"] = df_det["Importo €"].map(lambda x: f"€ {x:,.2f}")
        df_det["P.U. €"]    = df_det["P.U. €"].map(lambda x: f"€ {x:,.4f}")
        df_det["Quantità"]  = df_det["Quantità"].map(lambda x: f"{x:,.3f}")
        st.dataframe(df_det, use_container_width=True, hide_index=True, height=300)
        st.markdown(
            f'<div class="total-row"><span>TOTALE COMPLESSIVO ({n_voci} voci)</span>'
            f'<span>€ {total_imp:,.2f}</span></div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4  –  IMPORTA XLSX
# ══════════════════════════════════════════════════════════════════════════════
with tab_import:
    st.markdown("### 📥 Importa Computo da XLSX")
    st.markdown("Carica un Excel con un computo esistente. Mappa le colonne tramite indici 0-based.")

    xlsx_up = st.file_uploader("XLSX computo", type=["xlsx","xls"], key="comp_xlsx")
    if xlsx_up:
        raw_bytes = xlsx_up.read()
        try:
            xls = pd.ExcelFile(io.BytesIO(raw_bytes))
            st.info(f"Fogli: {', '.join(xls.sheet_names)}")
            sheet_sel = st.selectbox("Seleziona foglio", xls.sheet_names)
            df_prev   = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=sheet_sel, header=None)
            st.dataframe(df_prev.head(12), use_container_width=True)

            st.markdown("**Mappa colonne (indici 0-based):**")
            ci = st.columns(8)
            h_row  = ci[0].number_input("Riga hdr",    0, value=0,  step=1, key="h_row")
            c_cat  = ci[1].number_input("Categoria",   0, value=1,  step=1, key="c_cat")
            c_scat = ci[2].number_input("Sottocateg.", 0, value=2,  step=1, key="c_scat")
            c_cod  = ci[3].number_input("Codice",      0, value=3,  step=1, key="c_cod")
            c_desc = ci[4].number_input("Descrizione", 0, value=4,  step=1, key="c_desc")
            c_um   = ci[5].number_input("UM",          0, value=5,  step=1, key="c_um")
            c_q    = ci[6].number_input("Quantità",    0, value=10, step=1, key="c_q")
            c_pu   = ci[7].number_input("Prezzo U.",   0, value=11, step=1, key="c_pu")

            if st.button("📥 Importa nel computo", type="primary"):
                voci_imp, new_id = import_computo_from_xlsx(
                    raw_bytes, sheet_sel,
                    int(h_row), int(c_cat), int(c_scat), int(c_cod),
                    int(c_desc), int(c_um), int(c_q), int(c_pu),
                    start_id=st.session_state.next_id,
                )
                st.session_state.computo.extend(voci_imp)
                st.session_state.next_id = new_id
                aggiorna_importi(st.session_state.computo)
                st.success(f"✅ {len(voci_imp)} voci importate!"); st.rerun()
        except Exception as e:
            st.error(f"Errore: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5  –  ESPORTA
# ══════════════════════════════════════════════════════════════════════════════
with tab_export:
    st.markdown("### 📤 Esporta Computo")

    if not st.session_state.computo:
        st.info("Il computo è vuoto.")
    else:
        titolo = st.session_state.titolo_progetto
        ex1, ex2, ex3 = st.columns(3)

        with ex1:
            st.markdown("#### 📊 Excel (.xlsx)")
            st.markdown("Formule vive: Simili×L×W×H, importi =L×M, totale SUMIF. "
                        "Fogli: *Computo* + *Riepilogo WBS*.")
            try:
                xl = export_excel(st.session_state.computo, titolo)
                st.download_button("📊 Scarica Excel", data=xl, file_name="computo_metrico.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True, type="primary", key="4")
            except Exception as e:
                st.error(f"Errore Excel: {e}")

        with ex2:
            st.markdown("#### 📄 PDF (A4 landscape)")
            st.markdown("PDF professionale con libretto misure, separatori categoria e footer paginato.")
            try:
                pdf = export_pdf(st.session_state.computo, titolo)
                st.download_button("📄 Scarica PDF", data=pdf, file_name="computo_metrico.pdf",
                                   mime="application/pdf",
                                   use_container_width=True, type="primary", key="5")
            except ImportError:
                st.warning("⚠️ `pip install reportlab`")
            except Exception as e:
                st.error(f"Errore PDF: {e}")

        with ex3:
            st.markdown("#### 📃 CSV (separatore ;)")
            st.markdown("Formato tabulare: Prg., Articolo, Breve, Lotto, WBS, Quantità, P.U., Importo.")
            csv = computo_to_dataframe(st.session_state.computo).to_csv(index=False, sep=";", decimal=",")
            st.download_button("📃 Scarica CSV", data=csv, file_name="computo_metrico.csv",
                               mime="text/csv", use_container_width=True, key="6")

        st.markdown("---")
        st.markdown("#### Anteprima dati esportati")
        df_prev = computo_to_dataframe(st.session_state.computo).copy()
        df_prev["Importo €"] = df_prev["Importo €"].map(lambda x: f"€ {x:,.2f}")
        df_prev["P.U. €"]    = df_prev["P.U. €"].map(lambda x: f"€ {x:,.2f}")
        df_prev["Quantità"]  = df_prev["Quantità"].map(lambda x: f"{x:,.3f}")
        st.dataframe(df_prev, use_container_width=True, hide_index=True)
        st.markdown(
            f'<div class="total-row"><span>TOTALE COMPLESSIVO</span>'
            f'<span>€ {total_imp:,.2f}</span></div>', unsafe_allow_html=True)
