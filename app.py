"""
app.py  –  Interfaccia Streamlit – Computo Metrico Estimativo v2.0
==================================================================
Layout ispirato a Primus:
  · Tab Editor    → Split: Prezziario (sx) | Tabella Sintesi v2 (dx)
  · Tab Libretto  → Misurazioni voce: commento/simili/lung/larg/alt/totale
                    + pannello tipo voce (standard|sovrapprezzo_pct|riferimento)
  · Tab Riepilogo → WBS/Lotto + grafici
  · Tab Importa   → Excel con mapping colonne
  · Tab Esporta   → Excel formule + PDF + CSV

Avvio:
    streamlit run app.py
"""

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
    lookup_voce_by_codice,
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
# CSS
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

.main-header {
    background: linear-gradient(135deg, #1A1F36 0%, #2D3561 50%, #4A6CF7 100%);
    color: white; padding: 1.2rem 2rem; border-radius: 12px;
    margin-bottom: 1rem; display: flex; align-items: center; gap: 1rem;
}
.main-header h1 { margin:0; font-size:1.6rem; font-weight:700; letter-spacing:-0.5px; }
.main-header p  { margin:0; opacity:0.75; font-size:0.82rem; }

.metric-card {
    background: white; border: 1px solid #E8ECF4; border-radius: 10px;
    padding: 1rem 1.25rem; text-align: center;
    box-shadow: 0 2px 8px rgba(26,31,54,0.06);
}
.metric-card .value { font-size:1.5rem; font-weight:700; color:#2D3561; font-family:'DM Mono',monospace; }
.metric-card .label { font-size:0.72rem; color:#999; text-transform:uppercase; letter-spacing:0.6px; margin-top:0.2rem; }

.tag-cod {
    display:inline-block; background:#E8F0FE; color:#2D3561;
    padding:1px 8px; border-radius:12px; font-size:0.75rem;
    font-weight:600; font-family:'DM Mono',monospace;
}
.tag-sovr {
    display:inline-block; background:#FFF3E0; color:#E65100;
    padding:1px 8px; border-radius:12px; font-size:0.72rem;
    font-weight:700; font-family:'DM Mono',monospace;
}
.tag-rif {
    display:inline-block; background:#E8F5E9; color:#2E7D32;
    padding:1px 8px; border-radius:12px; font-size:0.72rem;
    font-weight:700; font-family:'DM Mono',monospace;
}
.tag-sub {
    display:inline-block; background:#FCE4EC; color:#C62828;
    padding:1px 8px; border-radius:12px; font-size:0.72rem;
    font-weight:700; font-family:'DM Mono',monospace;
}

.badge-imp {
    background:#E6F4EA; color:#1E7E34; padding:2px 8px;
    border-radius:6px; font-weight:700; font-family:'DM Mono',monospace;
    font-size:0.82rem;
}

.total-row {
    background: linear-gradient(90deg,#1A1F36,#2D3561); color:white;
    padding:.9rem 1.4rem; border-radius:10px;
    font-family:'DM Mono',monospace; font-weight:600; font-size:1rem;
    display:flex; justify-content:space-between; margin-top:.8rem;
}

.sec-label {
    font-size:0.68rem; text-transform:uppercase; letter-spacing:1.5px;
    color:#aaa; font-weight:600; margin:.8rem 0 .4rem 0;
}

.voce-card {
    background:#F0F4FF; border-left:3px solid #4A6CF7;
    border-radius:0 8px 8px 0; padding:.6rem .9rem; margin:.3rem 0;
}
.voce-card-sovr {
    background:#FFF8F0; border-left:3px solid #FF9800;
    border-radius:0 8px 8px 0; padding:.6rem .9rem; margin:.3rem 0;
}
.voce-card-rif {
    background:#F0FFF4; border-left:3px solid #4CAF50;
    border-radius:0 8px 8px 0; padding:.6rem .9rem; margin:.3rem 0;
}

.mis-row-sub { background:#FFF0F3; border-radius:6px; }
.mis-row-rif { background:#F0FFF4; border-radius:6px; }

[data-testid="stSidebar"] { background:#F0F2F8; }

.stTabs [data-baseweb="tab-list"] { gap:.4rem; background:#F0F2F8; padding:.4rem; border-radius:10px; }
.stTabs [data-baseweb="tab"] { border-radius:8px; font-weight:500; font-size:0.88rem; }

.stButton > button { border-radius:8px; font-weight:600; }

[data-testid="stDataFrameResizable"] thead th {
    background:#1A1F36 !important; color:white !important; font-size:0.78rem;
}
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# CACHE
# ──────────────────────────────────────────────────────────────────────────────
@st.cache_resource
def _get_cache() -> PrezziarioCache:
    return PrezziarioCache()

_cache = _get_cache()


# ──────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ──────────────────────────────────────────────────────────────────────────────
def _init() -> None:
    defaults = {
        "prezziari":       {},
        "computo":         [],
        "next_id":         1,
        "voce_sel_id":     None,
        "titolo_progetto": "Computo Metrico Estimativo",
        "cache_caricata":  False,
    }
    for k, v in defaults.items():
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
    for v in st.session_state.computo:
        if v["id"] == vid:
            return v
    return None


def _label_voce(v: dict) -> str:
    art = v.get("articolo") or v.get("codice", "")
    brv = str(v.get("descrizione_breve") or v.get("descrizione", ""))[:35]
    prg = v.get("progressiva", "")
    return f"#{v['id']} [{prg}] {art} · {brv}"


# ──────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ──────────────────────────────────────────────────────────────────────────────
with st.sidebar:

    st.markdown('<p class="sec-label">📐 Progetto</p>', unsafe_allow_html=True)
    st.session_state.titolo_progetto = st.text_input(
        "Titolo progetto",
        value=st.session_state.titolo_progetto,
        placeholder="es. Ponte SR69 – Consolidamento soletta",
        label_visibility="collapsed",
    )

    st.markdown("---")
    st.markdown('<p class="sec-label">📚 Prezziari attivi</p>', unsafe_allow_html=True)

    if not st.session_state.prezziari:
        st.caption("Nessun prezziario caricato.")
    else:
        for nome, df in list(st.session_state.prezziari.items()):
            in_cache  = nome in _cache
            c1, c2, c3 = st.columns([4, 1, 1])
            cache_icon = "💾" if in_cache else "⚠️"
            c1.markdown(f"**{nome}**  \n<small>{len(df):,} voci  {cache_icon}</small>",
                        unsafe_allow_html=True)
            if not in_cache:
                if c2.button("💾", key=f"save_cache_{nome}",
                              help="Salva in cache"):
                    _cache.salva(nome, df)
                    st.success(f"'{nome}' salvato in cache!")
                    st.rerun()
            else:
                c2.markdown("✓", help="In cache")
            if c3.button("✕", key=f"del_prez_{nome}", help="Rimuovi dalla sessione"):
                del st.session_state.prezziari[nome]
                st.rerun()

    st.markdown("---")
    st.markdown('<p class="sec-label">➕ Gestione prezziari</p>', unsafe_allow_html=True)
    t_nuovo, t_cache, t_parquet = st.tabs(["📄 Nuovo", "🗄️ Cache", "📦 Parquet"])

    with t_nuovo:
        st.caption("Carica PDF o XLSX. Il risultato verrà salvato automaticamente in cache.")
        up_file   = st.file_uploader("PDF o XLSX", type=["pdf","xlsx","xls"],
                                      key="sidebar_up", label_visibility="collapsed")
        nome_prez = st.text_input("Nome prezziario", placeholder="es. NC-MP 2025", key="nome_prez")

        if st.button("📥 Analizza e carica", use_container_width=True, type="primary"):
            if not nome_prez:
                st.warning("Inserisci un nome.")
            elif not up_file:
                st.warning("Seleziona un file.")
            else:
                raw       = up_file.read()
                hash_src  = md5_bytes(raw)
                if nome_prez in _cache:
                    meta = next((m for m in _cache.lista() if m["nome"] == nome_prez), {})
                    if meta.get("hash_sorgente") == hash_src:
                        df_prez = _cache.carica(nome_prez)
                        st.session_state.prezziari[nome_prez] = df_prez
                        st.success(f"✅ '{nome_prez}' già in cache — {len(df_prez):,} voci!")
                        st.rerun()

                with st.spinner(f"Analisi {up_file.name}…"):
                    try:
                        if Path(up_file.name).suffix.lower() == ".pdf":
                            df_prez = extract_pdf_prezziario(raw, nome_prez)
                        else:
                            df_prez = extract_xlsx_prezziario(raw, nome_prez)
                        if df_prez.empty:
                            st.warning("Nessuna voce estratta. Verifica il formato.")
                        else:
                            _cache.salva(nome_prez, df_prez, hash_sorgente=hash_src)
                            st.session_state.prezziari[nome_prez] = df_prez
                            st.success(f"✅ {len(df_prez):,} voci caricate e salvate in cache!")
                            st.rerun()
                    except ImportError as e:
                        st.error(str(e))

    with t_cache:
        voci_cache = _cache.lista()
        if not voci_cache:
            st.caption("Cache vuota.")
        else:
            st.caption(f"**{len(voci_cache)}** prezziari in cache locale.")
            for meta in voci_cache:
                nome_c = meta["nome"]
                attivo = nome_c in st.session_state.prezziari
                stato  = "🟢 attivo" if attivo else "⚪ non caricato"
                with st.expander(f"**{nome_c}** — {meta['n_voci']:,} voci  {stato}"):
                    st.caption(f"Aggiornato: {meta['data_aggiornamento']}")
                    col_a, col_b, col_c = st.columns(3)
                    if not attivo:
                        if col_a.button("▶ Carica", key=f"load_cache_{nome_c}",
                                         use_container_width=True, type="primary"):
                            df_c = _cache.carica(nome_c)
                            st.session_state.prezziari[nome_c] = df_c
                            st.rerun()
                    else:
                        if col_a.button("⏹ Scarica", key=f"unload_cache_{nome_c}",
                                         use_container_width=True):
                            del st.session_state.prezziari[nome_c]
                            st.rerun()
                    try:
                        xl = _cache.esporta_xlsx(nome_c)
                        col_b.download_button("📊 XLSX", data=xl,
                            file_name=f"{nome_c}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_xlsx_{nome_c}", use_container_width=True)
                    except Exception:
                        pass
                    try:
                        pq = _cache.esporta_parquet(nome_c)
                        col_c.download_button("📦 .parquet", data=pq,
                            file_name=f"{nome_c}.parquet", mime="application/octet-stream",
                            key=f"dl_pq_{nome_c}", use_container_width=True)
                    except Exception:
                        pass
                    st.markdown(" ")
                    if st.button(f"🗑️ Elimina dalla cache", key=f"del_cache_{nome_c}",
                                  use_container_width=True):
                        _cache.elimina(nome_c)
                        if nome_c in st.session_state.prezziari:
                            del st.session_state.prezziari[nome_c]
                        st.rerun()

    with t_parquet:
        st.caption("Importa un .parquet esportato in precedenza.")
        pq_up   = st.file_uploader(".parquet prezziario", type=["parquet"],
                                    key="pq_up", label_visibility="collapsed")
        pq_nome = st.text_input("Nome da assegnare", placeholder="es. NC-MP 2025", key="pq_nome")
        if st.button("📦 Importa parquet", use_container_width=True, type="primary"):
            if not pq_nome:
                st.warning("Inserisci un nome.")
            elif not pq_up:
                st.warning("Seleziona un file .parquet.")
            else:
                try:
                    df_imp = _cache.importa_parquet(pq_nome, pq_up.read())
                    st.session_state.prezziari[pq_nome] = df_imp
                    info = dataframe_info(df_imp)
                    st.success(f"✅ '{pq_nome}' importato: {info['n_voci']:,} voci")
                    st.rerun()
                except Exception as e:
                    st.error(f"Errore importazione: {e}")

    st.markdown("---")
    st.markdown('<p class="sec-label">💾 Gestione progetto</p>', unsafe_allow_html=True)

    if st.session_state.computo:
        json_str = export_json(
            st.session_state.computo,
            list(st.session_state.prezziari.keys()),
        )
        st.download_button("💾 Salva progetto JSON", data=json_str.encode(),
                           file_name="computo.json", mime="application/json",
                           use_container_width=True)

    json_up = st.file_uploader("Riapri progetto JSON", type=["json"], key="json_up")
    if json_up and st.button("📂 Carica progetto", use_container_width=True):
        try:
            res = import_json(json_up.read())
            st.session_state.computo = res["computo"]
            aggiorna_importi(st.session_state.computo)
            st.success("Progetto caricato!")
            st.rerun()
        except ValueError as e:
            st.error(str(e))

    st.markdown(" ")
    if st.button("🗑️ Azzera computo", use_container_width=True):
        st.session_state.computo     = []
        st.session_state.voce_sel_id = None
        st.rerun()


# ──────────────────────────────────────────────────────────────────────────────
# HEADER + METRICHE
# ──────────────────────────────────────────────────────────────────────────────
aggiorna_importi(st.session_state.computo)
assegna_progressive(st.session_state.computo)

total_imp   = totale_computo(st.session_state.computo)
n_voci      = len(st.session_state.computo)
n_prez      = len(st.session_state.prezziari)
tot_v_prez  = sum(len(d) for d in st.session_state.prezziari.values())

st.markdown(f"""
<div class="main-header">
    <div style="font-size:2.4rem">📐</div>
    <div>
        <h1>{st.session_state.titolo_progetto}</h1>
        <p>Computo Metrico Estimativo · Sovrapprezzi · Riferimenti Voce · Libretto Misure</p>
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
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# TABS PRINCIPALI
# ──────────────────────────────────────────────────────────────────────────────
tab_editor, tab_libretto, tab_riepilogo, tab_import, tab_export = st.tabs([
    "📋 Editor Computo",
    "📏 Libretto Misure",
    "📊 Riepilogo WBS",
    "📥 Importa XLSX",
    "📤 Esporta",
])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 – EDITOR COMPUTO
# ══════════════════════════════════════════════════════════════════════════════
with tab_editor:
    col_prez, col_comp = st.columns([1, 1], gap="medium")

    # ── COLONNA SINISTRA – PREZZIARIO ─────────────────────────────────────────
    with col_prez:
        st.markdown("#### 🔍 Prezziario")

        all_voci_df = get_all_voci(st.session_state.prezziari)

        if all_voci_df.empty:
            st.info("👈 Carica un prezziario dalla barra laterale per iniziare.")
        else:
            query = st.text_input(
                "Cerca (codice o parola chiave)",
                placeholder="es.  T.10  calcestruzzo  scavo",
                key="prez_search",
            )
            col_f1, col_f2 = st.columns([3, 1])
            with col_f1:
                fonti      = ["Tutti"] + list(all_voci_df["FONTE"].unique())
                fonte_filt = st.selectbox("Fonte", fonti, key="fonte_filt",
                                          label_visibility="collapsed")
            with col_f2:
                max_r = st.selectbox("Max", [30, 50, 100, 200], key="max_r",
                                     label_visibility="collapsed")

            df_filt   = all_voci_df.copy()
            if fonte_filt != "Tutti":
                df_filt = df_filt[df_filt["FONTE"] == fonte_filt]

            risultati = cerca_voce(query, df_filt, max_results=int(max_r))
            st.markdown(f"<small>**{len(risultati)}** risultati</small>", unsafe_allow_html=True)

            for _, riga in risultati.iterrows():
                r_cod  = riga["CODICE"]
                r_desc = str(riga["DESCRIZIONE"])[:60]
                r_um   = riga.get("UM", "")
                r_pr   = float(riga["PREZZO"])

                c_cod, c_desc, c_pr, c_btn = st.columns([2, 5, 2, 1])
                c_cod.markdown(f"<span class='tag-cod'>{r_cod}</span>", unsafe_allow_html=True)
                c_desc.markdown(f"<small>{r_desc}</small>", unsafe_allow_html=True)
                c_pr.markdown(f"<small>**€ {r_pr:,.2f}**/{r_um}</small>", unsafe_allow_html=True)

                if c_btn.button("➕", key=f"add_{r_cod}", help=f"Aggiungi {r_cod}"):
                    mis = nuova_misurazione(commento="Misura 1", quantita_diretta=1.0)
                    v   = nuova_voce(
                        _nuovo_id(),
                        articolo=r_cod,
                        descrizione=riga["DESCRIZIONE"],
                        descrizione_breve=str(riga["DESCRIZIONE"])[:40],
                        um=r_um,
                        prezzo_unitario=r_pr,
                        misurazioni=[mis],
                    )
                    v["quantita_totale"] = quantita_totale_voce(v)
                    v["importo"]         = calcola_importo(v)
                    st.session_state.computo.append(v)
                    aggiorna_importi(st.session_state.computo)
                    assegna_progressive(st.session_state.computo)
                    st.success(f"➕ {r_cod} aggiunto al computo")
                    st.rerun()

                st.divider()

    # ── COLONNA DESTRA – TABELLA SINTESI ──────────────────────────────────────
    with col_comp:
        st.markdown("#### 📋 Tabella di Sintesi")

        if not st.session_state.computo:
            st.info("Il computo è vuoto. Clicca ➕ accanto a una voce del prezziario per iniziare.")
        else:
            # Prepara DataFrame editabile (Tabella Sintesi v2)
            df_edit = pd.DataFrame([
                {
                    "ID":          v["id"],
                    "Prg.":        v.get("progressiva", ""),
                    "Articolo":    v.get("articolo") or v.get("codice", ""),
                    "Breve":       str(v.get("descrizione_breve") or v.get("descrizione", ""))[:40],
                    "UM":          v.get("um", ""),
                    "Quantità":    round(v.get("quantita_totale", 0), 3),
                    "P.U. €":     round(v.get("prezzo_unitario", 0), 2),
                    "Importo €":  round(v.get("importo", 0), 2),
                    "Lotto":       v.get("lotto", ""),
                    "WBS":         v.get("wbs", ""),
                    "Categoria":   v.get("categoria", ""),
                    "Tipo":        v.get("tipo", "standard"),
                }
                for v in st.session_state.computo
            ])

            edited = st.data_editor(
                df_edit,
                use_container_width=True,
                height=400,
                hide_index=True,
                disabled=["ID", "Prg.", "Quantità", "Importo €"],
                column_config={
                    "ID":        st.column_config.NumberColumn("N.", width="small"),
                    "Prg.":      st.column_config.TextColumn("Prg.", width="small"),
                    "Articolo":  st.column_config.TextColumn("Articolo", width="small"),
                    "Breve":     st.column_config.TextColumn("Descrizione breve", width="large"),
                    "UM":        st.column_config.TextColumn("UM", width="small"),
                    "Quantità":  st.column_config.NumberColumn("Quantità", format="%.3f", width="small"),
                    "P.U. €":   st.column_config.NumberColumn("P.U. €", format="€ %.2f", width="small"),
                    "Importo €": st.column_config.NumberColumn("Importo €", format="€ %.2f", width="small"),
                    "Lotto":     st.column_config.TextColumn("Lotto", width="small"),
                    "WBS":       st.column_config.TextColumn("WBS", width="small"),
                    "Categoria": st.column_config.TextColumn("Categoria", width="medium"),
                    "Tipo":      st.column_config.SelectboxColumn(
                                    "Tipo", options=list(TIPI_VOCE), width="small"),
                },
                key="editor_computo",
            )

            # Sincronizza modifiche
            for _, row in edited.iterrows():
                vid  = int(row["ID"])
                voce = _voce_by_id(vid)
                if voce is None:
                    continue
                voce["articolo"]          = str(row.get("Articolo", "") or "")
                voce["descrizione_breve"] = str(row.get("Breve", "") or "")
                voce["prezzo_unitario"]   = float(row.get("P.U. €", 0) or 0)
                voce["lotto"]             = str(row.get("Lotto", "") or "")
                voce["wbs"]               = str(row.get("WBS", "") or "")
                voce["categoria"]         = str(row.get("Categoria", "") or "")
                voce["tipo"]              = str(row.get("Tipo", "standard") or "standard")
                voce["importo"]           = calcola_importo(voce, st.session_state.computo)

            aggiorna_importi(st.session_state.computo)
            assegna_progressive(st.session_state.computo)

            st.markdown(
                f'<div class="total-row">'
                f'<span>TOTALE COMPLESSIVO</span>'
                f'<span>€ {totale_computo(st.session_state.computo):,.2f}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )

            # Seleziona voce per libretto
            st.markdown(" ")
            ids_voci = [v["id"] for v in st.session_state.computo]
            sel_id   = st.selectbox(
                "📏 Seleziona voce per aprire il Libretto Misure",
                options=ids_voci,
                format_func=lambda i: _label_voce(_voce_by_id(i) or {}),
                key="sel_voce_libretto",
            )
            st.session_state.voce_sel_id = sel_id

            # Azioni riga
            c_del, c_up, c_dn = st.columns(3)
            if c_del.button("🗑️ Elimina voce", use_container_width=True):
                st.session_state.computo = [
                    v for v in st.session_state.computo if v["id"] != sel_id
                ]
                st.session_state.voce_sel_id = None
                st.rerun()
            if c_up.button("⬆️ Sposta su", use_container_width=True):
                idx = next((i for i, v in enumerate(st.session_state.computo) if v["id"] == sel_id), None)
                if idx and idx > 0:
                    c = st.session_state.computo
                    c[idx - 1], c[idx] = c[idx], c[idx - 1]
                    st.rerun()
            if c_dn.button("⬇️ Sposta giù", use_container_width=True):
                idx = next((i for i, v in enumerate(st.session_state.computo) if v["id"] == sel_id), None)
                c = st.session_state.computo
                if idx is not None and idx < len(c) - 1:
                    c[idx], c[idx + 1] = c[idx + 1], c[idx]
                    st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 – LIBRETTO DELLE MISURE v2
# ══════════════════════════════════════════════════════════════════════════════
with tab_libretto:
    st.markdown("### 📏 Libretto delle Misure")

    if not st.session_state.computo:
        st.info("Aggiungi prima delle voci nel computo.")
    else:
        vid = st.session_state.voce_sel_id
        if vid is None:
            vid = st.session_state.computo[0]["id"]

        voce = _voce_by_id(vid)
        if voce is None:
            st.warning("Voce non trovata.")
        else:
            _normalizza_voce(voce)
            tipo_voce = voce.get("tipo", "standard")

            # ── Intestazione voce ──────────────────────────────────────────────
            card_cls = {
                "standard":      "voce-card",
                "sovrapprezzo_pct": "voce-card-sovr",
                "riferimento":   "voce-card-rif",
            }.get(tipo_voce, "voce-card")

            tipo_badge = {
                "standard":         "",
                "sovrapprezzo_pct": f"<span class='tag-sovr'>SOVR. {voce.get('sovrapprezzo_pct',0)}%</span>&nbsp;",
                "riferimento":      f"<span class='tag-rif'>RIF.→#{voce.get('rif_voce_id','')}</span>&nbsp;",
            }.get(tipo_voce, "")

            art  = voce.get("articolo") or voce.get("codice", "—")
            desc = str(voce.get("descrizione", ""))[:90]
            brv  = str(voce.get("descrizione_breve", ""))[:50]
            prg  = voce.get("progressiva", "")
            um   = voce.get("um", "")
            pu   = voce.get("prezzo_unitario", 0)

            st.markdown(f"""
            <div class="{card_cls}">
                <div style="display:flex;align-items:center;gap:0.6rem;flex-wrap:wrap">
                    <span class="tag-cod">{art}</span>
                    {tipo_badge}
                    <strong>{prg}</strong>
                    <span style="color:#666;font-size:0.85rem">{desc}</span>
                </div>
                <div style="margin-top:0.35rem;display:flex;gap:1.2rem;font-size:0.82rem;color:#555">
                    <span>📋 Breve: <b>{brv or '—'}</b></span>
                    <span>📏 UM: <b>{um}</b></span>
                    <span>💶 P.U.: <b>€ {pu:,.4f}</b></span>
                </div>
            </div>
            """, unsafe_allow_html=True)

            st.markdown("---")

            # ── Pannello configurazione tipo voce ──────────────────────────────
            with st.expander("⚙️ Configurazione voce (tipo, descrizione, sovrapprezzo, riferimento)", expanded=(tipo_voce != "standard")):
                prop_c1, prop_c2 = st.columns([2, 3])

                with prop_c1:
                    nuovo_tipo = st.selectbox(
                        "Tipo voce",
                        options=list(TIPI_VOCE),
                        index=list(TIPI_VOCE).index(tipo_voce),
                        key=f"tipo_voce_{vid}",
                        help=(
                            "standard = calcolo normale  |  "
                            "sovrapprezzo_pct = importo% su voce rif.  |  "
                            "riferimento = copia quantità da voce padre"
                        ),
                    )
                    voce["tipo"] = nuovo_tipo

                    nuovo_art = st.text_input("Articolo / Codice tariffa",
                                              value=voce.get("articolo") or voce.get("codice", ""),
                                              key=f"art_{vid}")
                    voce["articolo"] = nuovo_art

                    nuovo_brv = st.text_input("Descrizione breve (max 50 car.)",
                                              value=voce.get("descrizione_breve", ""),
                                              max_chars=50, key=f"brv_{vid}")
                    voce["descrizione_breve"] = nuovo_brv

                    nuovo_lotto = st.text_input("Lotto", value=voce.get("lotto", ""),
                                                key=f"lotto_{vid}")
                    voce["lotto"] = nuovo_lotto

                with prop_c2:
                    nuovo_desc = st.text_area("Descrizione completa",
                                              value=voce.get("descrizione", ""),
                                              height=80, key=f"desc_{vid}")
                    voce["descrizione"] = nuovo_desc

                    if nuovo_tipo == "sovrapprezzo_pct":
                        # ── SOVRAPPREZZO ──────────────────────────────────────
                        st.markdown("##### 📈 Configurazione Sovrapprezzo")

                        st.info(
                            "**Opzione A – % applicata sull'importo totale della voce di riferimento.**\n\n"
                            "Seleziona la voce di riferimento e inserisci la percentuale. "
                            "L'importo sarà = Importo_voce_rif × %/100.\n\n"
                            "**Opzione B – Inserisci % come *Simili* nel libretto misure** "
                            "per applicare il sovrapprezzo sulla quantità direttamente."
                        )

                        # Voce di riferimento
                        altre_voci = [v for v in st.session_state.computo if v["id"] != vid]
                        opts_rif   = [None] + [v["id"] for v in altre_voci]
                        fmt_rif    = lambda i: ("— Nessun riferimento —" if i is None
                                                else _label_voce(_voce_by_id(i) or {}))
                        cur_rif    = voce.get("rif_voce_id")
                        idx_rif    = opts_rif.index(cur_rif) if cur_rif in opts_rif else 0
                        sel_rif    = st.selectbox("Voce di riferimento (base calcolo %)",
                                                  options=opts_rif,
                                                  format_func=fmt_rif,
                                                  index=idx_rif,
                                                  key=f"rif_art_{vid}")
                        voce["rif_voce_id"] = sel_rif

                        # Percentuale
                        pct_val = float(voce.get("sovrapprezzo_pct") or 0.0)
                        nuova_pct = st.number_input(
                            "Percentuale sovrapprezzo (%)",
                            min_value=0.0, max_value=999.9,
                            value=pct_val, step=0.5, format="%.2f",
                            key=f"sovr_pct_{vid}",
                            help="Es. 15.0 → sovrapprezzo del 15% sull'importo della voce rif."
                        )
                        voce["sovrapprezzo_pct"] = nuova_pct

                        # Anteprima calcolo
                        if sel_rif and nuova_pct:
                            padre = _voce_by_id(sel_rif)
                            if padre:
                                base = calcola_importo(padre, st.session_state.computo)
                                importo_sovr = round(base * nuova_pct / 100.0, 2)
                                st.success(
                                    f"**Anteprima:** {nuova_pct}% di € {base:,.2f} "
                                    f"= **€ {importo_sovr:,.2f}**"
                                )

                    elif nuovo_tipo == "riferimento":
                        # ── RIFERIMENTO VOCE ─────────────────────────────────
                        st.markdown("##### 🔗 Configurazione Riferimento Voce")
                        st.info(
                            "La quantità di questa voce sarà automaticamente uguale "
                            "alla quantità totale della voce selezionata come riferimento. "
                            "Se la voce padre viene modificata, questa si aggiorna automaticamente."
                        )

                        altre_voci = [v for v in st.session_state.computo if v["id"] != vid]
                        opts_rif   = [None] + [v["id"] for v in altre_voci]
                        fmt_rif    = lambda i: ("— Nessun riferimento —" if i is None
                                                else _label_voce(_voce_by_id(i) or {}))
                        cur_rif    = voce.get("rif_voce_id")
                        idx_rif    = opts_rif.index(cur_rif) if cur_rif in opts_rif else 0
                        sel_rif    = st.selectbox("Voce padre (fonte della quantità)",
                                                  options=opts_rif,
                                                  format_func=fmt_rif,
                                                  index=idx_rif,
                                                  key=f"rif_voce_{vid}")
                        voce["rif_voce_id"] = sel_rif

                        if sel_rif:
                            padre = _voce_by_id(sel_rif)
                            if padre:
                                qt_padre = quantita_totale_voce(padre, st.session_state.computo)
                                st.success(
                                    f"**Quantità ereditata:** {qt_padre:.4f} {padre.get('um','')}"
                                    f" (da voce #{sel_rif} – {padre.get('articolo','')}) "
                                )

            st.markdown("---")

            # ── TABELLA MISURAZIONI ────────────────────────────────────────────
            st.markdown("#### Misurazioni di dettaglio")

            # Se tipo riferimento, le misurazioni vengono dalla voce padre
            if tipo_voce == "riferimento" and voce.get("rif_voce_id"):
                st.caption("ℹ️ Questa voce eredita la quantità dalla voce padre. Le misurazioni non sono modificabili direttamente.")
                padre = _voce_by_id(voce["rif_voce_id"])
                if padre:
                    qt_ereditata = quantita_totale_voce(padre, st.session_state.computo)
                    st.info(f"Quantità ereditata: **{qt_ereditata:.4f} {voce.get('um','')}**")
            else:
                misurazioni = voce.get("misurazioni") or []

                # Costruisce DataFrame per data_editor con tipo_riga
                df_mis = pd.DataFrame([
                    {
                        "Tipo riga":    m.get("tipo_riga", "standard"),
                        "Rif. voce #":  m.get("rif_voce_id", None),
                        "Commento":     m.get("commento", ""),
                        "Simili (n°)":  float(m.get("simili", 1) or 1),
                        "Lunghezza":    float(m.get("lung",  0) or 0),
                        "Larghezza":    float(m.get("larg",  0) or 0),
                        "Altezza":      float(m.get("alt",   0) or 0),
                        "Q. diretta":   float(m.get("quantita_diretta", 0) or 0),
                        "Totale":       round(quantita_misurazione(m, st.session_state.computo), 4),
                    }
                    for m in misurazioni
                ]) if misurazioni else pd.DataFrame(columns=[
                    "Tipo riga", "Rif. voce #", "Commento",
                    "Simili (n°)", "Lunghezza", "Larghezza", "Altezza", "Q. diretta", "Totale",
                ])

                # Opzioni voce per selectbox riferimento
                ids_altre = {str(v["id"]): _label_voce(v)
                             for v in st.session_state.computo if v["id"] != vid}

                edited_mis = st.data_editor(
                    df_mis,
                    use_container_width=True,
                    num_rows="dynamic",
                    hide_index=False,
                    disabled=["Totale"],
                    column_config={
                        "Tipo riga":   st.column_config.SelectboxColumn(
                            "Tipo riga",
                            options=list(TIPI_RIGA),
                            width="small",
                            help="standard=normale | sottrazione=in detrazione | riferimento_voce=copia qta da voce",
                        ),
                        "Rif. voce #": st.column_config.NumberColumn(
                            "Rif. #",
                            width="small",
                            help="ID della voce da referenziare (solo se tipo=riferimento_voce)",
                            min_value=0,
                        ),
                        "Commento":    st.column_config.TextColumn(
                            "Commento / Descrizione misura", width="large"),
                        "Simili (n°)": st.column_config.NumberColumn(
                            "Simili (n°)", format="%.2f", min_value=0.0, width="small",
                            help="Numero di pezzi/tratte uguali (default 1)"),
                        "Lunghezza":   st.column_config.NumberColumn(
                            "Lung. (m)", format="%.3f", min_value=0.0, width="small"),
                        "Larghezza":   st.column_config.NumberColumn(
                            "Larg. (m)", format="%.3f", min_value=0.0, width="small"),
                        "Altezza":     st.column_config.NumberColumn(
                            "Alt. (m)", format="%.3f", min_value=0.0, width="small"),
                        "Q. diretta":  st.column_config.NumberColumn(
                            "Q. diretta", format="%.4f", min_value=0.0, width="small",
                            help="Usata solo se Lung/Larg/Alt sono tutti 0"),
                        "Totale":      st.column_config.NumberColumn(
                            "Totale", format="%.4f", width="small"),
                    },
                    key=f"mis_editor_{vid}",
                )

                # Pulsante salva misurazioni
                if st.button("💾 Salva misurazioni", type="primary",
                              use_container_width=True, key="salva_mis"):
                    nuove_mis = []
                    for _, r in edited_mis.iterrows():
                        tipo_riga = str(r.get("Tipo riga", "standard") or "standard")
                        rif_id    = r.get("Rif. voce #")
                        rif_id    = int(rif_id) if rif_id and not pd.isna(rif_id) else None

                        m_new = nuova_misurazione(
                            commento         = str(r.get("Commento", "") or ""),
                            simili           = float(r.get("Simili (n°)", 1) or 1),
                            lung             = float(r.get("Lunghezza",  0) or 0),
                            larg             = float(r.get("Larghezza",  0) or 0),
                            alt              = float(r.get("Altezza",    0) or 0),
                            quantita_diretta = float(r.get("Q. diretta", 0) or 0),
                            tipo_riga        = tipo_riga,
                            rif_voce_id      = rif_id,
                        )
                        nuove_mis.append(m_new)

                    voce["misurazioni"]     = nuove_mis
                    voce["quantita_totale"] = quantita_totale_voce(voce, st.session_state.computo)
                    voce["importo"]         = calcola_importo(voce, st.session_state.computo)
                    aggiorna_importi(st.session_state.computo)

                    st.success(
                        f"✅ {len(nuove_mis)} righe salvate  —  "
                        f"Quantità totale: **{voce['quantita_totale']:.4f} {voce.get('um','')}**  —  "
                        f"Importo: **€ {voce['importo']:,.2f}**"
                    )
                    st.rerun()

                # Legenda tipi riga
                with st.expander("ℹ️ Legenda tipi riga misurazioni"):
                    st.markdown("""
**`standard`** – Calcolo normale:
- Se Lung/Larg/Alt > 0 → `Simili × Lung × Larg × Alt` (i valori = 0 vengono ignorati)
- Altrimenti → `Simili × Q. diretta`

**`sottrazione`** – Come standard, ma il risultato è in **detrazione** (valore negativo).
Utile per sottrarre spazi vuoti (aperture, cavità, ecc.).

**`riferimento_voce`** – La riga assume automaticamente la quantità totale della voce
indicata in *Rif. #*. Utile per collegare dinamicamente misure tra voci diverse.
Si aggiorna in automatico se la voce padre cambia.
                    """)

            # ── RIEPILOGO VOCE ─────────────────────────────────────────────────
            qt_curr  = quantita_totale_voce(voce, st.session_state.computo)
            imp_curr = calcola_importo(voce, st.session_state.computo)

            st.markdown(" ")
            col_q, col_i = st.columns(2)
            col_q.markdown(
                f'<div class="total-row" style="margin-top:0">'
                f'<span>Quantità totale voce</span>'
                f'<span>{qt_curr:.4f} {voce.get("um","")}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )
            col_i.markdown(
                f'<div class="total-row" style="margin-top:0;background:linear-gradient(90deg,#1E7E34,#28A745)">'
                f'<span>Importo voce</span>'
                f'<span>€ {imp_curr:,.2f}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 – RIEPILOGO WBS
# ══════════════════════════════════════════════════════════════════════════════
with tab_riepilogo:
    st.markdown("### 📊 Riepilogo WBS / Lotti / Categorie")

    if not st.session_state.computo:
        st.info("Nessuna voce nel computo.")
    else:
        df_wbs = riepilogo_wbs(st.session_state.computo)

        # Filtro Lotto
        lotti = ["Tutti"] + sorted(df_wbs["Lotto"].unique().tolist())
        lotto_filt = st.selectbox("Filtra per Lotto", lotti, key="lotto_filt")
        df_wbs_show = df_wbs if lotto_filt == "Tutti" else df_wbs[df_wbs["Lotto"] == lotto_filt]

        col_g, col_t = st.columns([3, 2])
        with col_g:
            chart_df = (
                df_wbs_show.set_index("Categoria")[["Importo €"]]
                .sort_values("Importo €")
            )
            st.bar_chart(chart_df, height=320, color="#4A6CF7")

        with col_t:
            disp_wbs = df_wbs_show.copy()
            disp_wbs["Importo €"] = disp_wbs["Importo €"].map(lambda x: f"€ {x:,.2f}")
            st.dataframe(disp_wbs, use_container_width=True, hide_index=True, height=320)

        st.markdown("---")
        st.markdown("#### Dettaglio completo")

        df_det = computo_to_dataframe(st.session_state.computo).copy()
        df_det["Importo €"] = df_det["Importo €"].map(lambda x: f"€ {x:,.2f}")
        df_det["P.U. €"]    = df_det["P.U. €"].map(lambda x: f"€ {x:,.4f}")
        df_det["Quantità"]  = df_det["Quantità"].map(lambda x: f"{x:,.3f}")
        st.dataframe(df_det, use_container_width=True, hide_index=True, height=300)

        st.markdown(
            f'<div class="total-row">'
            f'<span>TOTALE COMPLESSIVO ({n_voci} voci)</span>'
            f'<span>€ {total_imp:,.2f}</span>'
            f'</div>',
            unsafe_allow_html=True,
        )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 – IMPORTA XLSX COMPUTO ESISTENTE
# ══════════════════════════════════════════════════════════════════════════════
with tab_import:
    st.markdown("### 📥 Importa Computo da XLSX")
    st.markdown(
        "Carica un file Excel con un computo esistente (formato libero). "
        "Mappa manualmente le colonne tramite gli indici 0-based."
    )

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
            h_row   = ci[0].number_input("Riga hdr",    min_value=0, value=0,  step=1, key="h_row")
            c_cat   = ci[1].number_input("Categoria",   min_value=0, value=1,  step=1, key="c_cat")
            c_scat  = ci[2].number_input("Sottocateg.", min_value=0, value=2,  step=1, key="c_scat")
            c_cod   = ci[3].number_input("Codice",      min_value=0, value=3,  step=1, key="c_cod")
            c_desc  = ci[4].number_input("Descrizione", min_value=0, value=4,  step=1, key="c_desc")
            c_um    = ci[5].number_input("UM",          min_value=0, value=5,  step=1, key="c_um")
            c_q     = ci[6].number_input("Quantità",    min_value=0, value=10, step=1, key="c_q")
            c_pu    = ci[7].number_input("Prezzo U.",   min_value=0, value=11, step=1, key="c_pu")

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
                st.success(f"✅ {len(voci_imp)} voci importate!")
                st.rerun()
        except Exception as e:
            st.error(f"Errore: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 – ESPORTA
# ══════════════════════════════════════════════════════════════════════════════
with tab_export:
    st.markdown("### 📤 Esporta Computo")

    if not st.session_state.computo:
        st.info("Il computo è vuoto. Aggiungi delle voci prima di esportare.")
    else:
        titolo = st.session_state.titolo_progetto

        col_xl, col_pdf, col_csv = st.columns(3)

        with col_xl:
            st.markdown("#### 📊 Excel (.xlsx)")
            st.markdown(
                "File con **formule Excel vive**: Simili×L×W×H, "
                "importi =L×M, totale con SUMIF. "
                "Due fogli: *Computo* e *Riepilogo WBS*. "
                "Colonne aggiornate: Prg. · Lotto · WBS · Articolo · Breve."
            )
            try:
                xl_bytes = export_excel(st.session_state.computo, titolo)
                st.download_button(
                    "📊 Scarica Excel", data=xl_bytes,
                    file_name="computo_metrico.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True, type="primary",
                )
            except Exception as e:
                st.error(f"Errore Excel: {e}")

        with col_pdf:
            st.markdown("#### 📄 PDF (A4 landscape)")
            st.markdown(
                "PDF professionale con libretto misure, "
                "separatori categoria, sovrapprezzi/riferimenti annotati, "
                "totale e footer paginato."
            )
            try:
                pdf_bytes = export_pdf(st.session_state.computo, titolo)
                st.download_button(
                    "📄 Scarica PDF", data=pdf_bytes,
                    file_name="computo_metrico.pdf",
                    mime="application/pdf",
                    use_container_width=True, type="primary",
                )
            except ImportError:
                st.warning("⚠️ reportlab non installato. Esegui:  `pip install reportlab`")
            except Exception as e:
                st.error(f"Errore PDF: {e}")

        with col_csv:
            st.markdown("#### 📃 CSV (separatore ;)")
            st.markdown(
                "Formato tabulare con tutte le colonne: "
                "Prg., Articolo, Breve, Lotto, WBS, Quantità, P.U., Importo, Tipo voce."
            )
            csv_data = computo_to_dataframe(st.session_state.computo).to_csv(
                index=False, sep=";", decimal=","
            )
            st.download_button(
                "📃 Scarica CSV", data=csv_data,
                file_name="computo_metrico.csv",
                mime="text/csv", use_container_width=True,
            )

        st.markdown("---")
        st.markdown("#### Anteprima dati esportati")
        df_prev = computo_to_dataframe(st.session_state.computo).copy()
        df_prev["Importo €"] = df_prev["Importo €"].map(lambda x: f"€ {x:,.2f}")
        df_prev["P.U. €"]    = df_prev["P.U. €"].map(lambda x: f"€ {x:,.2f}")
        df_prev["Quantità"]  = df_prev["Quantità"].map(lambda x: f"{x:,.3f}")
        st.dataframe(df_prev, use_container_width=True, hide_index=True)

        st.markdown(
            f'<div class="total-row">'
            f'<span>TOTALE COMPLESSIVO</span>'
            f'<span>€ {total_imp:,.2f}</span>'
            f'</div>',
            unsafe_allow_html=True,
        )
