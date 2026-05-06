"""
Suivi longitudinal de la longueur axiale
==========================================
- Lit la table Documents depuis PUBLIC.MDB
- Résout le dossier patient via find_patient_folder (inspiré de patient.py)
- Extrait CodePatient depuis le chemin, date depuis le nom de fichier
- Extrait AL_OD / AL_OG depuis chaque PDF MMT
- Affiche nom/prénom du patient (table Patient de PUBLIC.MDB)

Dépendances : pip install pandas matplotlib pdfplumber pyodbc
Driver requis : Microsoft Access Database Engine 2016 (64 bits)
"""

import logging
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pdfplumber
import pyodbc

# ═════════════════════════════════════════════════════════════════════════════
# ▶▶  CONFIGURATION
# ═════════════════════════════════════════════════════════════════════════════
PATIENT_ID = "66844742"  # ex: "1758507609" pour aller directement, None = menu interactif
FICHIER_MDB   = Path(r"C:\Stage\database\baseSQL\PUBLIC.MDB")
DEST_PHOTOS   = Path(r"c:\Stage\database\donnés_pdf")   # racine réseau des dossiers photos
SAUVEGARDER   = False        # True = PNG sur disque, False = fenêtre interactive

# ═════════════════════════════════════════════════════════════════════════════

log = logging.getLogger(__name__)
RE_AL = re.compile(r"Comp\.\s*AL\s*:\s*([\d.]+)\s*mm")


# ─────────────────────────────────────────────────────────────────────────────
# 1. CONNEXION MDB
# ─────────────────────────────────────────────────────────────────────────────

def _db_connect() -> pyodbc.Connection:
    if not FICHIER_MDB.exists():
        raise FileNotFoundError(f"PUBLIC.MDB introuvable : {FICHIER_MDB}")
    return pyodbc.connect(
        f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={FICHIER_MDB};"
    )


# ─────────────────────────────────────────────────────────────────────────────
# 2. RÉSOLUTION DU DOSSIER PATIENT (adapté de find_patient_folder.py)
# ─────────────────────────────────────────────────────────────────────────────

def find_patient_folder(patient_code: str, conn: pyodbc.Connection) -> Path | None:
    try:
        cursor = conn.cursor()
        # Priorité aux fichiers MMT
        cursor.execute(
            "SELECT TOP 1 [Photo externe] FROM Documents "
            "WHERE [code patient] = ? "
            "AND [Photo externe] IS NOT NULL "
            "AND [Photo externe] LIKE '%MMT%'",
            (int(patient_code),)
        )
        row = cursor.fetchone()

        # Fallback : n'importe quel document
        if not row:
            cursor.execute(
                "SELECT TOP 1 [Photo externe] FROM Documents "
                "WHERE [code patient] = ? AND [Photo externe] IS NOT NULL",
                (int(patient_code),)
            )
            row = cursor.fetchone()

        

        if not row or not row[0]:
            log.warning(f"Aucun document trouvé pour le patient {patient_code}")
            return None

        
        chemin = Path(row[0].strip())
        dossier = DEST_PHOTOS / Path(*chemin.parts[1:-1])  # exclure racine et fichier

        if not dossier.is_dir():
            log.warning(f"Dossier absent sur le disque : {dossier}")
            return None

        return dossier

    except Exception as e:
        log.error(f"Erreur lors de la recherche en base : {e}")
        return None


# ─────────────────────────────────────────────────────────────────────────────
# 3. EXTRACTION PDF
# ─────────────────────────────────────────────────────────────────────────────

def code_patient_depuis_chemin(chemin: str) -> str | None:
    m = re.search(r"\\(\d{7,})[^\\]*\\[^\\]+\.pdf$", chemin)
    return m.group(1) if m else None


def date_depuis_nom_fichier(nom: str) -> pd.Timestamp | None:
    """
    MMT-Full_20260227_023237_2921.pdf  →  Timestamp("2026-02-27")
    """
    m = re.search(r"_(\d{8})_", nom)
    if m:
        try:
            return pd.Timestamp(datetime.strptime(m.group(1), "%Y%m%d"))
        except ValueError:
            return None
    return None


def extraire_al_pdf(pdf_path: Path) -> tuple[float | None, float | None]:
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            texte = "\n".join(p.extract_text() for p in pdf.pages if p.extract_text())
        valeurs = RE_AL.findall(texte)
        return (
            float(valeurs[0]) if len(valeurs) > 0 else None,
            float(valeurs[1]) if len(valeurs) > 1 else None,
        )
    except Exception as e:
        print(f"    ⚠  {pdf_path.name} : {e}")
        return None, None


# ─────────────────────────────────────────────────────────────────────────────
# 4. CHARGEMENT BIOMÉTRIE 
# ─────────────────────────────────────────────────────────────────────────────

def _extraire_depuis_mdb(patient_code: str | None = None) -> pd.DataFrame:
    conn = _db_connect()

    # Filtre SQL si patient connu → on ne lit que ses documents
    if patient_code:
        docs = pd.read_sql(
            "SELECT * FROM [Documents] WHERE [code patient] = ?",
            conn, params=(int(patient_code),)
        )
    else:
        docs = pd.read_sql("SELECT * FROM [Documents]", conn)

    patients = pd.read_sql(
        "SELECT [Code patient], [NOM], [Prénom] FROM [Patients]", conn
    )
    

    col = "Photo externe"
    if col not in docs.columns:
        raise KeyError(f"Colonne '{col}' absente de la table Documents")

    mmt = docs[docs[col].astype(str).str.contains("MMT", na=False)].copy()
    print(f"  {len(mmt)} PDF(s) MMT référencé(s) dans Documents")

    resultats = []
    for _, row in mmt.iterrows():
        chemin      = str(row[col])
        code_pat    = code_patient_depuis_chemin(chemin)
        nom_pdf     = Path(chemin).name
        date_mesure = date_depuis_nom_fichier(nom_pdf)

        # Résolution du chemin physique via find_patient_folder
        if code_pat:
            dossier = find_patient_folder(code_pat,conn)
            pdf_path = dossier / nom_pdf if dossier else None
        else:
            pdf_path = None

        if pdf_path is None or not pdf_path.exists():
            print(f"    ⚠  PDF introuvable : {nom_pdf}")
            al_od, al_og = None, None
        else:
            al_od, al_og = extraire_al_pdf(pdf_path)
            print(f"    ✓  {nom_pdf}  OD={al_od}  OG={al_og}")

        resultats.append({
            "CodePatient": str(code_pat) if code_pat else None,
            "DateMesure":  date_mesure,
            "AL_OD":       al_od,
            "AL_OG":       al_og,
            "fichier":     nom_pdf,
        })
    conn.close()

    df = pd.DataFrame(resultats).sort_values("DateMesure").reset_index(drop=True)

    # Jointure nom/prénom
    patients["CodePatient"] = patients["Code patient"].astype(str)
    df = df.merge(patients[["CodePatient", "NOM", "Prénom"]], on="CodePatient", how="left")
    
    print(f"  {df['AL_OD'].notna().sum()} AL OD extraite(s), "
          f"{df['AL_OG'].notna().sum()} AL OG extraite(s)")
    return df



def load_biometrie(patient_code: str | None = None) -> pd.DataFrame:
    print("  Extraction depuis PUBLIC.MDB et PDFs…")
    return _extraire_depuis_mdb(patient_code)


# ─────────────────────────────────────────────────────────────────────────────
# 5. SÉLECTION INTERACTIVE
# ─────────────────────────────────────────────────────────────────────────────

def choisir_patient(df: pd.DataFrame) -> str:
    patients = sorted(df["CodePatient"].dropna().unique())
    total    = len(patients)

    print()
    print("┌─────────────────────────────────────────────────┐")
    print("│        SÉLECTION DU PATIENT (BIOMÉTRIE)         │")
    print("├─────────────────────────────────────────────────┤")
    print(f"│  {total} patient(s) disponible(s)                    │")
    print("│                                                 │")
    print("│  Numéro, ID, ou nom/prénom                      │")
    print("└─────────────────────────────────────────────────┘")
    print()

    for i, pid in enumerate(patients):
        sub    = df[df["CodePatient"] == pid]
        nom    = sub["NOM"].iloc[0]    if "NOM"    in sub.columns else "?"
        prenom = sub["Prénom"].iloc[0] if "Prénom" in sub.columns else ""
        n      = sub["DateMesure"].notna().sum()
        print(f"  [{i}]  {prenom} {nom}  (ID {pid}, {n} mesure(s))")
    print()

    while True:
        saisie = input("  Votre choix : ").strip()

        # Par numéro
        if saisie.isdigit() and int(saisie) < total:
            return patients[int(saisie)]

        # Par ID direct
        if saisie in [str(p) for p in patients]:
            return saisie

        # Par nom/prénom
        terme = saisie.upper()
        correspondances = [
            pid for pid in patients
            if terme in str(df[df["CodePatient"] == pid]["NOM"].iloc[0]).upper()
            or terme in str(df[df["CodePatient"] == pid]["Prénom"].iloc[0]).upper()
        ]
        if correspondances:
            print(f"\n  {len(correspondances)} patient(s) trouvé(s) :")
            for pid in correspondances:
                sub = df[df["CodePatient"] == pid]
                print(f"    {sub['Prénom'].iloc[0]} {sub['NOM'].iloc[0]}  (ID {pid})")
            print()
            confirm = input("  Tracer ces patient(s) ? [Entrée=oui / n=non] : ").strip().lower()
            if confirm != "n":
                return correspondances[0] if len(correspondances) == 1 else correspondances
            print()
            continue

        print(f"  ⚠  Entrée non reconnue — réessayez")


# ─────────────────────────────────────────────────────────────────────────────
# 6. VISUALISATION
# ─────────────────────────────────────────────────────────────────────────────

def plot_al(df: pd.DataFrame, code_patient: str) -> None:
    pat = df[df["CodePatient"] == code_patient].dropna(
        subset=["DateMesure"]
    ).sort_values("DateMesure")

    if pat.empty:
        print(f"  Aucune donnée pour {code_patient}")
        return

    nom    = pat["NOM"].iloc[0]    if "NOM"    in pat.columns else code_patient
    prenom = pat["Prénom"].iloc[0] if "Prénom" in pat.columns else ""
    n      = len(pat)
    od     = pat.dropna(subset=["AL_OD"])
    og     = pat.dropna(subset=["AL_OG"])
    all_al = pd.concat([od["AL_OD"], og["AL_OG"]]).dropna()

    fig, ax = plt.subplots(figsize=(12, 5))
    fig.patch.set_facecolor("#0f172a")
    ax.set_facecolor("#111827")
    ax.grid(True, color="#1e293b", linewidth=0.8, zorder=0)
    ax.set_axisbelow(True)

    # Zone normale
    ax.axhspan(22.0, 24.0, color="#10b981", alpha=0.05, zorder=0,
               label="Globe normal (22–24 mm)")
    ax.axhline(23.0, color="#10b981", linestyle=":", linewidth=1.0, alpha=0.4)

    # Courbes
    ax.plot(od["DateMesure"], od["AL_OD"], marker="^", color="#a78bfa",
            linewidth=2.0, markersize=7, label="Longueur axiale OD (mm)")
    ax.plot(og["DateMesure"], og["AL_OG"], marker="v", color="#f472b6",
            linewidth=2.0, markersize=7, label="Longueur axiale OG (mm)")

    # Annotations pour tous les points (alternance d'offset pour limiter les chevauchements)
    for col, color, base_offset in [("AL_OD", "#a78bfa", 8), ("AL_OG", "#f472b6", 8)]:
        sub = pat.dropna(subset=[col])
        if sub.empty:
            continue
        for idx, (_, row) in enumerate(sub.iterrows()):
            val = row[col]
            date = row["DateMesure"]
            v_offset = base_offset if idx % 2 == 0 else -base_offset
            ax.annotate(
                f"{val:.2f} mm",
                xy=(date, val),
                xytext=(6, v_offset), textcoords="offset points",
                fontsize=8, color=color,
                bbox=dict(boxstyle="round,pad=0.2", fc="#0f172a", ec=color, alpha=0.85),
                arrowprops=dict(arrowstyle="-", color=color, alpha=0.5),
            )

    # Axe X
    if n > 1:
        date_range = (pat["DateMesure"].max() - pat["DateMesure"].min()).days
        if date_range > 365 * 10:
            ax.xaxis.set_major_locator(mdates.YearLocator(5))
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
        elif date_range > 365 * 3:
            ax.xaxis.set_major_locator(mdates.YearLocator())
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
        elif date_range > 180:
            ax.xaxis.set_major_locator(mdates.MonthLocator(bymonth=[1, 4, 7, 10]))
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
        else:
            ax.xaxis.set_major_locator(mdates.MonthLocator())
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=35, ha="right", fontsize=9)

    # Axe Y
    if not all_al.empty:
        ax.set_ylim(all_al.min() - 0.3, all_al.max() + 0.3)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:.2f} mm"))
    ax.tick_params(colors="#64748b", labelsize=9)
    for spine in ax.spines.values():
        spine.set_edgecolor("#1e293b")

    ax.set_xlabel("Date de mesure", color="#94a3b8", fontsize=10, labelpad=8)
    ax.set_ylabel("Longueur axiale (mm)", color="#94a3b8", fontsize=10, labelpad=8)
    ax.set_title(
        f"Évolution de la longueur axiale — {prenom} {nom}".strip(),
        color="#e2e8f0", fontsize=13, fontweight="bold", pad=14
    )
    ax.legend(fontsize=8, framealpha=0.25, facecolor="#0f172a",
              edgecolor="#1e293b", labelcolor="#94a3b8", loc="upper left")

    
    if not od.empty and len(od) > 1:
        delta = od["AL_OD"].iloc[-1] - od["AL_OD"].iloc[0]
        duree = (od["DateMesure"].max() - od["DateMesure"].min()).days / 365.25
        ax.text(
            0.99, 0.97,
            f"n={n} mesure(s)  |  durée {duree:.1f} ans\n"
            f"OD : {od['AL_OD'].min():.2f} → {od['AL_OD'].max():.2f} mm  (Δ {delta:+.2f} mm)",
            transform=ax.transAxes, ha="right", va="top",
            fontsize=8, color="#64748b", fontfamily="monospace",
            bbox=dict(boxstyle="round,pad=0.4", fc="#0f172a", ec="#1e293b", alpha=0.7),
        )

    plt.tight_layout()
    if SAUVEGARDER:
        outpath = Path(f"al_{prenom}_{nom}.png".replace(" ", "_"))
        fig.savefig(outpath, dpi=150, bbox_inches="tight", facecolor=fig.get_facecolor())
        print(f"  Sauvegardé : {outpath}")
    else:
        plt.show()
    plt.close(fig)


# ─────────────────────────────────────────────────────────────────────────────
# 7. POINT D'ENTRÉE
# ─────────────────────────────────────────────────────────────────────────────

def main():
    logging.basicConfig(level=logging.WARNING, format="%(levelname)-8s %(message)s")

    # Si PATIENT_ID configuré → extraction ciblée, pas de menu
    if PATIENT_ID:
        print(f"▶ Chargement biométrie pour patient {PATIENT_ID}…")
        df  = load_biometrie(PATIENT_ID)
        pid = PATIENT_ID
    else:
        print("▶ Chargement de toute la biométrie…")
        df    = load_biometrie()
        choix = choisir_patient(df)
        pids  = [choix] if isinstance(choix, str) else choix
        pid   = pids[0]

    sub    = df[df["CodePatient"] == str(pid)]
    nom    = sub["NOM"].iloc[0]    if not sub.empty and "NOM"    in sub.columns else pid
    prenom = sub["Prénom"].iloc[0] if not sub.empty and "Prénom" in sub.columns else ""
    print(f"\n▶ Tracé pour {prenom} {nom}…")
    plot_al(df, str(pid))
    print("✓ Terminé.")


if __name__ == "__main__":
    main()