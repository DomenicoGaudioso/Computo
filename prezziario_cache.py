"""
prezziario_cache.py  –  Cache persistente dei prezziari su disco
================================================================
I prezziari vengono parsati UNA SOLA VOLTA da PDF/XLSX, poi salvati
come file .parquet nella cartella  ./cache_prezziari/

Al riavvio dell'applicazione, la cache viene caricata automaticamente
senza bisogno di ricaricare i file originali.

Struttura della cartella cache:
    cache_prezziari/
        NC-MP_2025.parquet       ← DataFrame con CODICE,DESCRIZIONE,UM,PREZZO,FONTE
        MR_2025.parquet
        manifest.json            ← metadati: {nome: {file, n_voci, data, hash_src}}

API principale:
    cache = PreziarioCache()          # carica manifest al volo
    cache.salva(nome, df)             # persiste un prezziario
    cache.carica(nome) → DataFrame   # legge il parquet
    cache.lista() → list[dict]        # elenco prezziarii in cache con metadati
    cache.elimina(nome)               # rimuove dal disco
    cache.carica_tutti() → dict       # {nome: DataFrame} di tutti quelli in cache
    cache.esporta_xlsx(nome) → bytes  # DataFrame → XLSX scaricabile
    cache.importa_parquet(nome, bytes) → DataFrame  # carica da file .parquet esterno
"""

from __future__ import annotations

import hashlib
import io
import json
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd


# ──────────────────────────────────────────────────────────────────────────────
# CONFIGURAZIONE
# ──────────────────────────────────────────────────────────────────────────────

DEFAULT_CACHE_DIR = Path("cache_prezziari")
MANIFEST_FILE     = "manifest.json"
COLONNE           = ["CODICE", "DESCRIZIONE", "UM", "PREZZO", "FONTE"]


# ──────────────────────────────────────────────────────────────────────────────
# CLASSE CACHE
# ──────────────────────────────────────────────────────────────────────────────

class PrezziarioCache:
    """
    Gestisce la cache su disco dei prezziari in formato Parquet.

    Parametri
    ----------
    cache_dir : Path  –  cartella dove salvare i file (default: ./cache_prezziari)
    """

    def __init__(self, cache_dir: Path | str = DEFAULT_CACHE_DIR) -> None:
        self.cache_dir = Path(cache_dir)
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self._manifest: dict[str, dict] = self._load_manifest()

    # ── Manifest ──────────────────────────────────────────────────────────────

    def _manifest_path(self) -> Path:
        return self.cache_dir / MANIFEST_FILE

    def _load_manifest(self) -> dict[str, dict]:
        p = self._manifest_path()
        if p.exists():
            try:
                return json.loads(p.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, OSError):
                return {}
        return {}

    def _save_manifest(self) -> None:
        self._manifest_path().write_text(
            json.dumps(self._manifest, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    # ── Nome file ─────────────────────────────────────────────────────────────

    @staticmethod
    def _safe_filename(nome: str) -> str:
        """Converte il nome del prezziario in un nome file sicuro."""
        safe = "".join(c if c.isalnum() or c in "._- " else "_" for c in nome)
        return safe.strip().replace(" ", "_") + ".parquet"

    def _parquet_path(self, nome: str) -> Path:
        return self.cache_dir / self._safe_filename(nome)

    # ── CRUD ──────────────────────────────────────────────────────────────────

    def salva(self, nome: str, df: pd.DataFrame, hash_sorgente: str = "") -> None:
        """
        Persiste un DataFrame prezziario su disco come .parquet
        e aggiorna il manifest.

        Parametri
        ----------
        nome           : chiave/nome del prezziario
        df             : DataFrame con colonne CODICE,DESCRIZIONE,UM,PREZZO,FONTE
        hash_sorgente  : hash MD5 del file originale (opzionale, per rilevare aggiornamenti)
        """
        p = self._parquet_path(nome)
        df.to_parquet(p, index=False, engine="pyarrow")

        self._manifest[nome] = {
            "file":          p.name,
            "n_voci":        len(df),
            "data_aggiornamento": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "hash_sorgente": hash_sorgente,
        }
        self._save_manifest()

    def carica(self, nome: str) -> pd.DataFrame:
        """
        Legge il prezziario dalla cache.
        Lancia KeyError se il nome non è in cache,
        FileNotFoundError se il file .parquet manca.
        """
        if nome not in self._manifest:
            raise KeyError(f"Prezziario '{nome}' non trovato in cache.")
        p = self._parquet_path(nome)
        if not p.exists():
            # Rimuovi voce orfana dal manifest
            del self._manifest[nome]
            self._save_manifest()
            raise FileNotFoundError(f"File cache mancante per '{nome}': {p}")
        return pd.read_parquet(p, engine="pyarrow")

    def elimina(self, nome: str) -> None:
        """Rimuove il prezziario dalla cache (disco + manifest)."""
        p = self._parquet_path(nome)
        if p.exists():
            p.unlink()
        if nome in self._manifest:
            del self._manifest[nome]
            self._save_manifest()

    def rinomina(self, nome_vecchio: str, nome_nuovo: str) -> None:
        """
        Rinomina un prezziario in cache.
        Sposta il file .parquet e aggiorna il manifest.
        """
        if nome_vecchio not in self._manifest:
            raise KeyError(f"Prezziario '{nome_vecchio}' non trovato.")
        if nome_nuovo in self._manifest:
            raise ValueError(f"Un prezziario di nome '{nome_nuovo}' esiste già.")

        old_path = self._parquet_path(nome_vecchio)
        new_path = self._parquet_path(nome_nuovo)

        if old_path.exists():
            old_path.rename(new_path)

        entry = dict(self._manifest[nome_vecchio])
        entry["file"] = new_path.name
        del self._manifest[nome_vecchio]
        self._manifest[nome_nuovo] = entry
        self._save_manifest()

    # ── Accesso bulk ──────────────────────────────────────────────────────────

    def lista(self) -> list[dict]:
        """
        Restituisce la lista dei prezziari in cache con i loro metadati.

        Ogni dict ha le chiavi:
            nome, file, n_voci, data_aggiornamento, hash_sorgente
        """
        result = []
        for nome, meta in self._manifest.items():
            p = self._parquet_path(nome)
            result.append({
                "nome":               nome,
                "file":               meta.get("file", ""),
                "n_voci":             meta.get("n_voci", 0),
                "data_aggiornamento": meta.get("data_aggiornamento", ""),
                "hash_sorgente":      meta.get("hash_sorgente", ""),
                "presente":           p.exists(),
            })
        return result

    def nomi(self) -> list[str]:
        """Lista dei nomi dei prezziari in cache."""
        return list(self._manifest.keys())

    def carica_tutti(self) -> dict[str, pd.DataFrame]:
        """
        Carica tutti i prezziari dalla cache in un dict {nome: DataFrame}.
        Salta silenziosamente i file mancanti (mostra warning nel log).
        """
        result: dict[str, pd.DataFrame] = {}
        for nome in list(self._manifest.keys()):
            try:
                result[nome] = self.carica(nome)
            except FileNotFoundError:
                pass  # File orfano già rimosso da carica()
        return result

    def __contains__(self, nome: str) -> bool:
        return nome in self._manifest

    def __len__(self) -> int:
        return len(self._manifest)

    # ── Import / Export ───────────────────────────────────────────────────────

    def esporta_xlsx(self, nome: str) -> bytes:
        """
        Esporta un prezziario dalla cache come file Excel (.xlsx) scaricabile.
        Utile per condividere o fare backup del prezziario processato.
        """
        df = self.carica(nome)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Prezziario", index=False)
            # Formatta larghezze colonne
            ws = writer.sheets["Prezziario"]
            ws.column_dimensions["A"].width = 18   # CODICE
            ws.column_dimensions["B"].width = 70   # DESCRIZIONE
            ws.column_dimensions["C"].width = 8    # UM
            ws.column_dimensions["D"].width = 12   # PREZZO
            ws.column_dimensions["E"].width = 20   # FONTE
        buf.seek(0)
        return buf.getvalue()

    def esporta_parquet(self, nome: str) -> bytes:
        """
        Restituisce il contenuto grezzo del file .parquet per il download.
        L'utente può salvarlo e ricaricarlo la prossima volta.
        """
        p = self._parquet_path(nome)
        if not p.exists():
            raise FileNotFoundError(f"File cache mancante per '{nome}'.")
        return p.read_bytes()

    def importa_parquet(self, nome: str, parquet_bytes: bytes) -> pd.DataFrame:
        """
        Carica un .parquet esterno (precedentemente scaricato con esporta_parquet)
        e lo aggiunge alla cache locale.

        Parametri
        ----------
        nome           : nome da assegnare al prezziario
        parquet_bytes  : contenuto binario del file .parquet

        Restituisce
        -----------
        DataFrame caricato
        """
        df = pd.read_parquet(io.BytesIO(parquet_bytes), engine="pyarrow")

        # Normalizza colonne mancanti
        for col in COLONNE:
            if col not in df.columns:
                df[col] = "" if col != "PREZZO" else 0.0

        df = df[COLONNE].copy()
        self.salva(nome, df)
        return df


# ──────────────────────────────────────────────────────────────────────────────
# FUNZIONI DI UTILITÀ STANDALONE
# ──────────────────────────────────────────────────────────────────────────────

def md5_bytes(data: bytes) -> str:
    """Calcola l'hash MD5 di dati binari (per rilevare file identici)."""
    return hashlib.md5(data).hexdigest()


def dataframe_info(df: pd.DataFrame) -> dict:
    """
    Restituisce un dict di informazioni di sintesi su un DataFrame prezziario:
        n_voci, fonti, prezzo_min, prezzo_max, prezzo_medio
    """
    if df.empty:
        return {"n_voci": 0, "fonti": [], "prezzo_min": 0, "prezzo_max": 0, "prezzo_medio": 0}
    return {
        "n_voci":       len(df),
        "fonti":        list(df["FONTE"].unique()) if "FONTE" in df.columns else [],
        "prezzo_min":   float(df["PREZZO"].min()),
        "prezzo_max":   float(df["PREZZO"].max()),
        "prezzo_medio": float(df["PREZZO"].mean()),
    }
