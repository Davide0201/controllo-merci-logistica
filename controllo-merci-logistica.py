"""
Controllo Merci e DDT
---------------------
Script per il controllo automatico di un elenco merci in entrata.
Legge un file Excel con i dati di ricevimento merci e produce un report
che evidenzia: righe duplicate, dati mancanti e un riepilogo per fornitore.

Autore: Davide Puppin
Contesto: progetto personale ispirato alla mia esperienza come
Operatore Logistico presso CAL Srl (controllo conformità merce/DDT).
"""

import pandas as pd
from datetime import datetime


def carica_dati(percorso_file: str) -> pd.DataFrame:
    """Carica il file Excel con i dati delle merci ricevute."""
    df = pd.read_excel(percorso_file)
    return df


def controlla_duplicati(df: pd.DataFrame, colonna_chiave: str = "Codice DDT") -> pd.DataFrame:
    """Restituisce le righe che hanno un codice DDT duplicato."""
    duplicati = df[df.duplicated(subset=[colonna_chiave], keep=False)]
    return duplicati


def controlla_dati_mancanti(df: pd.DataFrame) -> pd.DataFrame:
    """Restituisce le righe con almeno un campo obbligatorio vuoto."""
    righe_incomplete = df[df.isnull().any(axis=1)]
    return righe_incomplete


def riepilogo_per_fornitore(df: pd.DataFrame) -> pd.DataFrame:
    """Conta quanti articoli sono arrivati per ciascun fornitore."""
    riepilogo = df.groupby("Fornitore").size().reset_index(name="Numero articoli")
    return riepilogo


def genera_report(df: pd.DataFrame, percorso_output: str = "report_controllo_merci.xlsx"):
    """Genera un file Excel con tre fogli: duplicati, dati mancanti, riepilogo."""
    duplicati = controlla_duplicati(df)
    mancanti = controlla_dati_mancanti(df)
    riepilogo = riepilogo_per_fornitore(df)

    with pd.ExcelWriter(percorso_output, engine="openpyxl") as writer:
        duplicati.to_excel(writer, sheet_name="Duplicati", index=False)
        mancanti.to_excel(writer, sheet_name="Dati mancanti", index=False)
        riepilogo.to_excel(writer, sheet_name="Riepilogo fornitori", index=False)

    print(f"Report generato: {percorso_output}")
    print(f"- Righe duplicate trovate: {len(duplicati)}")
    print(f"- Righe con dati mancanti: {len(mancanti)}")
    print(f"- Fornitori distinti: {len(riepilogo)}")


def crea_file_di_esempio(percorso_file: str = "merci_esempio.xlsx"):
    """Crea un file Excel di esempio, utile per testare lo script subito."""
    dati_esempio = {
        "Codice DDT": ["DDT001", "DDT002", "DDT002", "DDT003", "DDT004"],
        "Fornitore": ["Fornitore A", "Fornitore B", "Fornitore B", "Fornitore A", None],
        "Descrizione Materiale": [
            "Componente aeronautico X",
            "Componente aeronautico Y",
            "Componente aeronautico Y",
            "Componente aeronautico Z",
            "Componente aeronautico W",
        ],
        "Quantita": [10, 5, 5, None, 8],
        "Data Ricevimento": [
            datetime(2026, 1, 10),
            datetime(2026, 1, 11),
            datetime(2026, 1, 11),
            datetime(2026, 1, 12),
            datetime(2026, 1, 13),
        ],
    }
    df = pd.DataFrame(dati_esempio)
    df.to_excel(percorso_file, index=False)
    print(f"File di esempio creato: {percorso_file}")


if __name__ == "__main__":
    # 1. Se non hai ancora un file vero, genera prima un esempio per testare:
    crea_file_di_esempio()

    # 2. Carica il file (sostituisci con il tuo file reale quando lo hai)
    dati = carica_dati("merci_esempio.xlsx")

    # 3. Genera il report di controllo
    genera_report(dati)
