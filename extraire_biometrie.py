"""
Extraction de longueur axiale depuis PDFs de biométrie
=======================================================
Lit tous les PDFs du dossier configuré et extrait :
  - Nom du patient
  - Date de naissance
  - Date de mesure (depuis le nom du fichier ET depuis le texte)
  - Longueur axiale composite OD et OG (Comp. AL)

Le résultat est sauvegardé en CSV pour jointure ultérieure
avec la base ophtalmologique.

Dépendances : pip install pdfplumber pandas
documents non inclus dans le dépôt (données sensibles)
"""

import re
import pdfplumber
import pandas as pd
from pathlib import Path
from datetime import datetime

from config import DOSSIER_PDF, FICHIER_CSV_BIOMETRIE as FICHIER_CSV, OUTPUT_CSV

# DOSSIER_PDF  → data/pdf/          (PDFs biométrie MMT-Full_*.pdf)
# FICHIER_CSV  → output/csv/biometrie_extraite.csv

RE_AL        = re.compile(r"Comp\.\s*AL\s*:\s*([\d.]+)\s*mm")
RE_NOM       = re.compile(r"Nom\s*:\s*(.+)")
RE_DOB       = re.compile(r"Date de naissance\s*:\s*(\d{2}/\d{2}/\d{4})")
RE_DATE_PDF  = re.compile(r"Date de mesure\s*[^:]*:\s*(\d{2}/\d{2}/\d{4})")

def date_depuis_nom_fichier(nom: str) -> str | None:
    """
    Extrait la date depuis le nom de fichier.
    Ex : MMT-Full_20260416_004404_1533.pdf → 2026-04-16
    """
    m = re.search(r"_(\d{8})_", nom)
    if m:
        s = m.group(1)  # "20260416"
        try:
            return datetime.strptime(s, "%Y%m%d").strftime("%Y-%m-%d")
        except ValueError:
            return None
    return None

def extraire_pdf(pdf_path: Path) -> dict:
    """
    Extrait toutes les informations utiles d'un PDF de biométrie.
    """
    with pdfplumber.open(str(pdf_path)) as pdf:
        texte = "\n".join(
            page.extract_text() for page in pdf.pages
            if page.extract_text()
        )

    def first(pattern):
        m = pattern.search(texte)
        return m.group(1).strip() if m else None

    # Longueurs axiales : 1ère occurrence = OD, 2ème = OG
    valeurs_al = RE_AL.findall(texte)

    # Date de mesure : priorité au nom de fichier (fiable), sinon texte PDF
    date_mesure = (
        date_depuis_nom_fichier(pdf_path.name)
        or first(RE_DATE_PDF)
    )

    return {
        "fichier":       pdf_path.name,
        "NOM":           first(RE_NOM),
        "DateNaissance": first(RE_DOB),
        "DateMesure":    date_mesure,
        "AL_OD":         float(valeurs_al[0]) if len(valeurs_al) > 0 else None,
        "AL_OG":         float(valeurs_al[1]) if len(valeurs_al) > 1 else None,
    }

def extraire_dossier(dossier: str) -> pd.DataFrame:
    """
    Parcourt tous les PDFs du dossier et retourne un DataFrame.
    """
    pdfs = sorted(Path(dossier).glob("MMT-Full_*.pdf"))

    if not pdfs:
        print(f"⚠  Aucun PDF trouvé dans : {dossier}")
        return pd.DataFrame()

    print(f"▶ {len(pdfs)} PDF(s) trouvé(s) dans {dossier}\n")

    resultats = []
    for pdf in pdfs:
        try:
            data = extraire_pdf(pdf)
            resultats.append(data)
            print(
                f"  ✓  {pdf.name}\n"
                f"     Nom         : {data['NOM']}\n"
                f"     Naissance   : {data['DateNaissance']}\n"
                f"     Mesure      : {data['DateMesure']}\n"
                f"     AL OD       : {data['AL_OD']} mm\n"
                f"     AL OG       : {data['AL_OG']} mm\n"
            )
        except Exception as e:
            print(f"  ✗  {pdf.name} : {e}\n")

    df = pd.DataFrame(resultats)

    # Convertir les dates en datetime
    df["DateMesure"]    = pd.to_datetime(df["DateMesure"],    errors="coerce")
    df["DateNaissance"] = pd.to_datetime(df["DateNaissance"], format="%d/%m/%Y", errors="coerce")

    return df

def main():
    df = extraire_dossier(DOSSIER_PDF)

    if df.empty:
        return

    print("─" * 50)
    print(f"  Total extraits  : {len(df)}")
    print(f"  AL OD manquant  : {df['AL_OD'].isna().sum()}")
    print(f"  AL OG manquant  : {df['AL_OG'].isna().sum()}")
    print(f"  OD moyen        : {df['AL_OD'].mean():.2f} mm")
    print(f"  OG moyen        : {df['AL_OG'].mean():.2f} mm")
    print("─" * 50)

    OUTPUT_CSV.mkdir(parents=True, exist_ok=True)
    df.to_csv(FICHIER_CSV, index=False, encoding="utf-8-sig")
    print(f"\n✓ Résultats sauvegardés : {FICHIER_CSV}")
    print(df.to_string(index=False))

if __name__ == "__main__":
    main()
