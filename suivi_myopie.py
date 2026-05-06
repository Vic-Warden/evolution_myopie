"""
Suivi longitudinal de la myopie
================================
Lecture depuis PUBLIC.MDB + biométrie depuis PDFs liés via Documents.json

Dépendances : pip install pandas matplotlib pyodbc pdfplumber
Driver requis : Microsoft Access Database Engine 2016 (64 bits)
https://www.microsoft.com/en-us/download/details.aspx?id=54920
"""

import json
import re
from pathlib import Path

# ═════════════════════════════════════════════════════════════════════════════
# ▶▶  CONFIGURATION
# ═════════════════════════════════════════════════════════════════════════════

FICHIER_MDB       = r"C:\Stage\database\baseSQL\PUBLIC.MDB"
DOSSIER_PDF       = r"c:\Stage\database\donnés_pdf"

PATIENT_IDS  = [66844742]  # [1758507609] = forcer 
TYPEREF      = 7      # 7 = Réfraction subjective 6 = Réfraction automatique, None = tous types
MODE_COHORTE = False
OEIL_COHORTE = "D"
SAUVEGARDER  = False

# ═════════════════════════════════════════════════════════════════════════════

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.dates as mdates
import pyodbc
import pdfplumber

# ─────────────────────────────────────────────────────────────────────────────
# 1. LECTURE BASE MDB
# ─────────────────────────────────────────────────────────────────────────────

def connect_mdb(chemin: str) -> pyodbc.Connection:
    conn_str = (
        r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
        rf"DBQ={chemin};"
    )
    try:
        return pyodbc.connect(conn_str)
    except pyodbc.Error as e:
        raise ConnectionError(
            f"Impossible de se connecter à {chemin}\n"
            f"Vérifiez que le driver Access 64 bits est installé.\n{e}"
        )


def load_table(conn: pyodbc.Connection, table: str) -> pd.DataFrame:
    return pd.read_sql(f"SELECT * FROM [{table}]", conn)


def load_all_tables(patient_ids: list[str] | None = None):
    print(f"▶ Connexion à {FICHIER_MDB}…")
    conn = connect_mdb(FICHIER_MDB)
    df_pat = load_table(conn, "Patients")

    if patient_ids:
        ids_str = ", ".join(patient_ids)
        df_con = pd.read_sql(
            f"SELECT * FROM [Consultation] WHERE [Code patient] IN ({ids_str})", conn
        )
        df_ref = pd.read_sql(
            f"SELECT r.* FROM [tREFRACTION] r "
            f"INNER JOIN [Consultation] c ON r.[NumConsult] = c.[N° consultation] "
            f"WHERE c.[Code patient] IN ({ids_str})", conn
        )
    else:
        df_con = load_table(conn, "Consultation")
        df_ref = load_table(conn, "tREFRACTION")

    conn.close()
    print(f"  patients={len(df_pat)}  consultations={len(df_con)}  réfractions={len(df_ref)}")
    return df_pat, df_con, df_ref

# ─────────────────────────────────────────────────────────────────────────────
# 2. UTILITAIRES
# ─────────────────────────────────────────────────────────────────────────────

def parse_fr_float(series: pd.Series) -> pd.Series:
    return (
        series.astype(str).str.strip()
        .str.replace(",", ".", regex=False)
        .apply(pd.to_numeric, errors="coerce")
    )


def calc_se(sph: pd.Series, cyl: pd.Series) -> pd.Series:
    return sph + cyl / 2.0


def _parse_dob(val) -> "pd.Timestamp":
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return pd.NaT
    s = str(val).strip().replace("\\", "")
    if not s:
        return pd.NaT
    m = re.match(r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})(?:\s+[\d:]+)?$", s)
    if m:
        month, day, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if year < 100:
            year += 1900 if year >= 25 else 2000
            if year > pd.Timestamp.now().year:
                year -= 100
        try:
            return pd.Timestamp(year=year, month=month, day=day)
        except ValueError:
            return pd.NaT
    ts = pd.to_datetime(s, errors="coerce", dayfirst=False)
    return ts if pd.notna(ts) else pd.NaT


# ─────────────────────────────────────────────────────────────────────────────
# 3. CONSTRUCTION DE L'HISTORIQUE
# ─────────────────────────────────────────────────────────────────────────────

def build_history(
    df_patients: pd.DataFrame,
    df_consult: pd.DataFrame,
    df_refrac: pd.DataFrame,
    typeref: int | None = None,
) -> pd.DataFrame:

    # Étape 1 : Jointure tREFRACTION ↔ Consultation
    consult_slim = df_consult[["N° consultation", "Code patient", "Date"]].copy()
    consult_slim.columns = ["NumConsult", "CodePatient", "Date"]
    consult_slim["NumConsult"]  = consult_slim["NumConsult"].astype(str)
    consult_slim["CodePatient"] = consult_slim["CodePatient"].astype(str)

    df_refrac = df_refrac.copy()
    df_refrac["NumConsult"] = df_refrac["NumConsult"].astype(str)
    merged = df_refrac.merge(consult_slim, on="NumConsult", how="left")

    # Étape 2 : Filtre TypeRef
    if typeref is not None:
        merged = merged[merged["TypeRef"].astype(str) == str(typeref)]
        print(f"  [TypeRef={typeref}] {len(merged)} lignes retenues")

    # Étape 3 : Jointure ↔ Patient
    pat_slim = df_patients[["Code patient", "NOM", "Prénom", "Date de Naissance"]].copy()
    pat_slim.columns = ["CodePatient", "NOM", "Prenom", "DateNaissance"]
    pat_slim["CodePatient"] = pat_slim["CodePatient"].astype(str)
    merged["CodePatient"] = merged["CodePatient"].astype(str)
    full = merged.merge(pat_slim, on="CodePatient", how="left")

    # Étape 4 : Types
    full["Date"] = full["Date"].apply(_parse_dob)
    for col in ["SphD", "CylD", "SphG", "CylG"]:
        full[col] = parse_fr_float(full[col])

    # Étape 5 : SE
    full["SE_D"] = calc_se(full["SphD"], full["CylD"])
    full["SE_G"] = calc_se(full["SphG"], full["CylG"])

    # Étape 6 : Âge
    full["DateNaissance"] = full["DateNaissance"].apply(_parse_dob)
    full["Age"] = ((full["Date"] - full["DateNaissance"]).dt.days / 365.25).round(1)

    full = full.sort_values(["CodePatient", "Date"]).reset_index(drop=True)
    print(f"  Historique : {len(full)} lignes, {full['CodePatient'].nunique()} patients")
    return full


# ─────────────────────────────────────────────────────────────────────────────
# 4. BIOMÉTRIE — extraction depuis Documents.json + PDFs
# ─────────────────────────────────────────────────────────────────────────────

RE_AL        = re.compile(r"Comp\.\s*AL\s*:\s*([\d.]+)\s*mm")
RE_DATE_PDF  = re.compile(r"Date de mesure\s*[^:]*:\s*(\d{2}/\d{2}/\d{4})")


def code_patient_depuis_chemin(chemin: str) -> str | None:
    """
    Extrait le CodePatient depuis le chemin.
    Ex : \\17.000\\1758507609dav.mat\\MMT-Full_...pdf  →  "1758507609"
    """
    m = re.search(r"\\(\d{7,})[^\\]*\\[^\\]+\.pdf$", chemin)
    return m.group(1) if m else None


def date_depuis_nom_fichier(nom: str) -> str | None:
    """
    MMT-Full_20260227_023237_2921.pdf  →  "2026-02-27"
    """
    m = re.search(r"_(\d{8})_", nom)
    if m:
        try:
            from datetime import datetime
            return datetime.strptime(m.group(1), "%Y%m%d").strftime("%Y-%m-%d")
        except ValueError:
            return None
    return None


def extraire_al_pdf(pdf_path: str) -> tuple[float | None, float | None]:
    """Retourne (AL_OD, AL_OG) depuis un PDF biométrie."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            texte = "\n".join(
                p.extract_text() for p in pdf.pages if p.extract_text()
            )
        valeurs = RE_AL.findall(texte)
        od = float(valeurs[0]) if len(valeurs) > 0 else None
        og = float(valeurs[1]) if len(valeurs) > 1 else None
        return od, og
    except Exception as e:
        print(f"    ⚠  {Path(pdf_path).name} : {e}")
        return None, None


def load_biometrie(patient_ids: list[str] | None = None) -> pd.DataFrame:
    print(f"▶ Chargement de la biométrie depuis {FICHIER_MDB}…")
    conn = connect_mdb(FICHIER_MDB)

    if patient_ids:
        ids_str = ", ".join(patient_ids)
        docs = pd.read_sql(
            f"SELECT * FROM [Documents] WHERE [code patient] IN ({ids_str})", conn
        )
    else:
        docs = pd.read_sql("SELECT * FROM [Documents]", conn)

    conn.close()

    col = "Photo externe"
    if col not in docs.columns:
        print(f"  ⚠  Colonne '{col}' absente de la table Documents")
        return pd.DataFrame()

    mmt = docs[docs[col].astype(str).str.contains("MMT", na=False)].copy()

    if mmt.empty:
        print("  ⚠  Aucun fichier MMT trouvé dans Documents.json")
        return pd.DataFrame()

    print(f"  {len(mmt)} PDF(s) MMT référencé(s) dans Documents.json")

    resultats = []
    for _, row in mmt.iterrows():
        chemin_relatif = str(row[col])
        code_patient   = code_patient_depuis_chemin(chemin_relatif)
        nom_pdf        = Path(chemin_relatif).name
        date_mesure    = date_depuis_nom_fichier(nom_pdf)

        # Construire le chemin absolu
        pdf_path = Path(DOSSIER_PDF) / chemin_relatif.lstrip("\\")

        if not pdf_path.exists():
            print(f"    ⚠  PDF introuvable : {pdf_path}")
            al_od, al_og = None, None
        else:
            al_od, al_og = extraire_al_pdf(str(pdf_path))
            print(f"    ✓  {nom_pdf}  →  OD={al_od}  OG={al_og}")

        resultats.append({
            "CodePatient": str(code_patient) if code_patient else None,
            "DateMesure":  pd.to_datetime(date_mesure, errors="coerce"),
            "AL_OD":       al_od,
            "AL_OG":       al_og,
            "fichier":     nom_pdf,
        })

    df_al = pd.DataFrame(resultats)
    n_match = df_al["CodePatient"].notna().sum()
    n_al    = df_al["AL_OD"].notna().sum()
    print(f"  Biométrie : {n_match} patient(s) identifié(s), {n_al} AL extraite(s)")
    return df_al


# ─────────────────────────────────────────────────────────────────────────────
# 5. VISUALISATION
# ─────────────────────────────────────────────────────────────────────────────

SEVERITY_ZONES = [
    (0,   -3,  "#f59e0b", "Myopie faible (0→−3 D)"),
    (-3,  -6,  "#f97316", "Myopie moyenne (−3→−6 D)"),
    (-6,  -20, "#ef4444", "Myopie forte (< −6 D)"),
]

STYLE = {
    "D": dict(color="#2563eb", marker="o", linewidth=2.2, markersize=6, label="Œil droit (SE_D)"),
    "G": dict(color="#dc2626", marker="s", linewidth=2.2, markersize=6, label="Œil gauche (SE_G)"),
}


def _add_severity_zones(ax, ymin, ymax):
    for top, bot, color, _ in SEVERITY_ZONES:
        ax.axhspan(max(bot, ymin-1), min(top, ymax+1), color=color, alpha=0.07, zorder=0)
    ax.axhline(0, color="#94a3b8", linestyle="--", linewidth=1.0, alpha=0.7, zorder=1)


def _severity_label(se):
    if se is None or pd.isna(se): return ""
    if se >  0.5: return "Hypermétrope"
    if se >= -0.5: return "Emmétrope"
    if se >= -3:  return "Myopie faible"
    if se >= -6:  return "Myopie moyenne"
    return "Myopie forte"


def _axe_x(ax, date_min, date_max):
    date_range = (date_max - date_min).days
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


def plot_patient(df, code_patient, df_al=None, save=False, output_dir="."):
    pat = df[df["CodePatient"] == str(code_patient)].dropna(subset=["Date"]).copy()
    pat = pat[pat[["SE_D", "SE_G"]].notna().any(axis=1)].sort_values("Date")

    if pat.empty:
        print(f"  Aucune donnée valide pour le patient {code_patient}")
        return

    nom    = pat["NOM"].iloc[0]    if "NOM"    in pat.columns else "?"
    prenom = pat["Prenom"].iloc[0] if "Prenom" in pat.columns else ""
    n_pts  = len(pat)

    # Biométrie
    al_pat = pd.DataFrame()
    if df_al is not None and not df_al.empty:
        al_pat = df_al[df_al["CodePatient"] == str(code_patient)].dropna(subset=["DateMesure"])
    has_al = not al_pat.empty

    # Figure
    fig, ax = plt.subplots(figsize=(13, 6))
    ax2 = ax.twinx() if has_al else None
    fig.patch.set_facecolor("#0f172a")
    ax.set_facecolor("#111827")
    ax.grid(True, color="#1e293b", linewidth=0.8, zorder=0)
    ax.set_axisbelow(True)

    all_se = pd.concat([pat["SE_D"].dropna(), pat["SE_G"].dropna()])
    ymin = min(-0.5, all_se.min()) - 0.5
    ymax = max(0.5,  all_se.max()) + 0.5
    _add_severity_zones(ax, ymin, ymax)

    # Courbes SE
    for eye in ("D", "G"):
        col = f"SE_{eye}"
        sub = pat.dropna(subset=[col])
        if sub.empty: continue
        ax.plot(sub["Date"], sub[col], zorder=3, **STYLE[eye])
        last = sub.iloc[-1]
        ax.annotate(
            f"{last[col]:+.2f} D\n{_severity_label(last[col])}",
            xy=(last["Date"], last[col]),
            xytext=(10, 8 if eye == "D" else -22),
            textcoords="offset points", fontsize=8,
            color=STYLE[eye]["color"],
            bbox=dict(boxstyle="round,pad=0.3", fc="#0f172a", ec=STYLE[eye]["color"], alpha=0.85),
            arrowprops=dict(arrowstyle="-", color=STYLE[eye]["color"], alpha=0.5),
        )

    zone_patches = [
        mpatches.Patch(color=c, alpha=0.35, label=lbl)
        for _, _, c, lbl in SEVERITY_ZONES
    ]

    # Courbes AL
    if has_al:
        ax2.plot(al_pat["DateMesure"], al_pat["AL_OD"],
                 marker="^", color="#a78bfa", linewidth=1.8,
                 markersize=7, linestyle="--", label="Longueur axiale OD (mm)")
        ax2.plot(al_pat["DateMesure"], al_pat["AL_OG"],
                 marker="v", color="#f472b6", linewidth=1.8,
                 markersize=7, linestyle="--", label="Longueur axiale OG (mm)")
        ax2.set_ylabel("Longueur axiale (mm)", color="#94a3b8", fontsize=10)
        ax2.tick_params(colors="#64748b", labelsize=9)
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:.2f} mm"))
        for spine in ax2.spines.values(): spine.set_edgecolor("#1e293b")

    # Axes
    _axe_x(ax, pat["Date"].min(), pat["Date"].max())
    ax.set_ylim(ymin, ymax)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:+.1f} D"))
    ax.tick_params(colors="#64748b", labelsize=9)
    for spine in ax.spines.values(): spine.set_edgecolor("#1e293b")

    ax.set_xlabel("Date de consultation", color="#94a3b8", fontsize=10, labelpad=8)
    ax.set_ylabel("Équivalent sphérique (dioptries)", color="#94a3b8", fontsize=10, labelpad=8)

    title = f"Évolution de la myopie — {prenom} {nom}".strip()
    if "Age" in pat.columns:
        ages = pat["Age"].dropna()
        ages = ages[ages > 0]
        if not ages.empty:
            title += f"  ({ages.min():.0f}→{ages.max():.0f} ans)"
    ax.set_title(title, color="#e2e8f0", fontsize=13, fontweight="bold", pad=14)

    # Légende
    handles, labels_leg = ax.get_legend_handles_labels()
    if has_al:
        h2, l2 = ax2.get_legend_handles_labels()
        ax.legend(handles + zone_patches + h2,
                  labels_leg + [p.get_label() for p in zone_patches] + l2,
                  loc="lower left", fontsize=8, framealpha=0.25,
                  facecolor="#0f172a", edgecolor="#1e293b", labelcolor="#94a3b8")
    else:
        ax.legend(handles + zone_patches,
                  labels_leg + [p.get_label() for p in zone_patches],
                  loc="lower left", fontsize=8, framealpha=0.25,
                  facecolor="#0f172a", edgecolor="#1e293b", labelcolor="#94a3b8")

    # Stats
    seD    = pat["SE_D"].dropna()
    seG    = pat["SE_G"].dropna()
    prog_D = (seD.iloc[-1] - seD.iloc[0]) if len(seD) > 1 else float("nan")
    duree  = (pat["Date"].max() - pat["Date"].min()).days / 365.25
    ax.text(0.99, 0.97,
            f"n={n_pts} mesures  |  durée {duree:.1f} ans\n"
            f"OD : {seD.min():+.2f} D → {seD.max():+.2f} D  (Δ {prog_D:+.2f} D)",
            transform=ax.transAxes, ha="right", va="top",
            fontsize=8, color="#64748b", fontfamily="monospace",
            bbox=dict(boxstyle="round,pad=0.4", fc="#0f172a", ec="#1e293b", alpha=0.7))

    plt.tight_layout()
    if save:
        outpath = Path(output_dir) / f"myopie_patient_{code_patient}.png"
        fig.savefig(outpath, dpi=150, bbox_inches="tight", facecolor=fig.get_facecolor())
        print(f"  Sauvegardé : {outpath}")
    else:
        plt.show()
    plt.close(fig)


def plot_cohort(df, patient_ids=None, eye="D", save=False, output_dir="."):
    ids    = patient_ids or sorted(df["CodePatient"].unique())
    se_col = f"SE_{eye}"

    fig, ax = plt.subplots(figsize=(14, 7))
    fig.patch.set_facecolor("#0f172a")
    ax.set_facecolor("#111827")
    ax.grid(True, color="#1e293b", linewidth=0.8, zorder=0)
    ax.set_axisbelow(True)

    all_se = df[se_col].dropna()
    if not all_se.empty:
        _add_severity_zones(ax, all_se.min() - 0.5, all_se.max() + 0.5)

    cmap = plt.colormaps["tab20"](range(len(ids)))
    for color, pid in zip(cmap, ids):
        sub = df[df["CodePatient"] == str(pid)].dropna(subset=["Date", se_col]).sort_values("Date")
        if sub.empty: continue
        nom = sub["NOM"].iloc[0]    if "NOM"    in sub.columns else str(pid)
        prn = sub["Prenom"].iloc[0] if "Prenom" in sub.columns else ""
        ax.plot(sub["Date"], sub[se_col], marker="o", markersize=4,
                linewidth=1.6, color=color, alpha=0.85, label=f"{prn} {nom}".strip())

    _axe_x(ax, df["Date"].min(), df["Date"].max())
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:+.1f} D"))
    ax.tick_params(colors="#64748b", labelsize=9)
    for spine in ax.spines.values(): spine.set_edgecolor("#1e293b")
    ax.set_xlabel("Date", color="#94a3b8", fontsize=10)
    ax.set_ylabel(f"SE œil {'droit' if eye=='D' else 'gauche'} (D)", color="#94a3b8", fontsize=10)
    ax.set_title(f"Évolution comparative — Cohorte ({len(ids)} patients)",
                 color="#e2e8f0", fontsize=13, fontweight="bold")
    ax.legend(fontsize=8, framealpha=0.2, facecolor="#0f172a",
              edgecolor="#1e293b", labelcolor="#e2e8f0", ncol=2)
    plt.tight_layout()
    if save:
        outpath = Path(output_dir) / "myopie_cohorte.png"
        fig.savefig(outpath, dpi=150, bbox_inches="tight", facecolor=fig.get_facecolor())
        print(f"  Sauvegardé : {outpath}")
    else:
        plt.show()
    plt.close(fig)


# ─────────────────────────────────────────────────────────────────────────────
# 6. SÉLECTION INTERACTIVE
# ─────────────────────────────────────────────────────────────────────────────

def choisir_patients(df: pd.DataFrame) -> list[str]:
    tous  = sorted(df["CodePatient"].unique())
    total = len(tous)

    ref = (
        df[["CodePatient", "NOM", "Prenom"]]
        .drop_duplicates("CodePatient")
        .set_index("CodePatient")
    )

    print()
    print("┌─────────────────────────────────────────────────┐")
    print("│           SÉLECTION DU PATIENT                  │")
    print("├─────────────────────────────────────────────────┤")
    print(f"│  {total} patient(s) disponible(s) dans la base      │")
    print("│                                                 │")
    print("│  Entrez un nom/prénom    →  ex: dupont          │")
    print("│  Entrez un ID patient    →  ex: 1758510666      │")
    print("│  Plusieurs IDs (virgule) →  ex: 101, 102        │")
    print("│  Appuyez Entrée seul     →  tous les patients   │")
    print("└─────────────────────────────────────────────────┘")
    print()

    while True:
        saisie = input("  Votre choix : ").strip()

        # Entrée vide → tous
        if saisie == "":
            print(f"  → Tous les patients ({total})")
            return [str(p) for p in tous]

        # ── Recherche par nom/prénom ──────────────────────────────────────────
        if not saisie.replace(",", "").replace(";", "").replace(" ", "").isdigit():
            terme = saisie.upper()
            correspondances = [
                pid for pid in tous
                if pid in ref.index and (
                    terme in str(ref.loc[pid, "NOM"]).upper() or
                    terme in str(ref.loc[pid, "Prenom"]).upper()
                )
            ]
            if correspondances:
                print(f"\n  {len(correspondances)} patient(s) trouvé(s) :")
                for pid in correspondances:
                    nom = ref.loc[pid, "NOM"]
                    prn = ref.loc[pid, "Prenom"]
                    print(f"    {pid}  {prn} {nom}")
                print()
                confirm = input("  Tracer ces patient(s) ? [Entrée=oui / n=non] : ").strip().lower()
                if confirm != "n":
                    return [str(p) for p in correspondances]
                print()
                continue
            else:
                print(f"  ⚠  Aucun patient trouvé pour '{saisie}'")
                continue

        # ── Saisie par ID ────────────────────────────────────────────────────
        morceaux = [m.strip() for m in saisie.replace(";", ",").split(",") if m.strip()]
        ids_valides, erreurs = [], []
        for m in morceaux:
            if m in [str(p) for p in tous]:
                ids_valides.append(m)
            else:
                erreurs.append(m)

        if erreurs:
            print(f"  ⚠  ID(s) inconnu(s) : {', '.join(erreurs)} — réessayez.")
            continue

        print(f"  → {len(ids_valides)} patient(s) sélectionné(s) : {', '.join(ids_valides)}")
        return ids_valides


# ─────────────────────────────────────────────────────────────────────────────
# 7. POINT D'ENTRÉE
# ─────────────────────────────────────────────────────────────────────────────

def main():
    # Chargement depuis MDB
    ids_config = [str(i) for i in PATIENT_IDS] if PATIENT_IDS else None

    df_pat, df_con, df_ref = load_all_tables(patient_ids=ids_config)

    print(f"▶ Construction de l'historique…")
    df = build_history(df_pat, df_con, df_ref, typeref=TYPEREF)

    # Biométrie filtrée sur les mêmes patients
    print("▶ Chargement de la biométrie…")
    df_al = load_biometrie(patient_ids=ids_config)

    # Sélection
    if PATIENT_IDS:
        ids = ids_config
        print(f"▶ Patient(s) configuré(s) : {', '.join(ids)}")
    else:
        ids = choisir_patients(df)
        # Si sélection interactive → recharger la bio pour ces patients seulement
        df_al = load_biometrie(patient_ids=ids)

    # Visualisation
    print()
    if MODE_COHORTE:
        print(f"▶ Courbe cohorte ({len(ids)} patients, œil {OEIL_COHORTE})…")
        plot_cohort(df, patient_ids=ids, eye=OEIL_COHORTE, save=SAUVEGARDER)
    else:
        print(f"▶ Tracé individuel de {len(ids)} patient(s)…")
        for pid in ids:
            plot_patient(df, pid, df_al=df_al, save=SAUVEGARDER)

    print("✓ Terminé.")


if __name__ == "__main__":
    main()
