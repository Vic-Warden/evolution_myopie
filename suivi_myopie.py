"""
Suivi longitudinal de la myopie
================================
Lecture depuis PUBLIC.MDB + biométrie depuis PDFs liés via Documents.json

Dépendances : pip install pandas matplotlib pyodbc pdfplumber pywin32
Driver requis : Microsoft Access Database Engine 2016 (64 bits)
https://www.microsoft.com/en-us/download/details.aspx?id=54920
"""

import json
import re
import sys
import time
import logging
import os
from pathlib import Path

# ▶▶  CONFIGURATION

FICHIER_MDB       = r"C:\Stage\database\baseSQL\PUBLIC.MDB"
DOSSIER_PDF       = r"c:\Stage\database\donnés_pdf"

PATIENT_IDS  = None  # [1758507609] = forcer 
TYPEREF      = 7      # 7 = Réfraction subjective 6 = Réfraction automatique, None = tous types
MODE_COHORTE = False
OEIL_COHORTE = "D"
SAUVEGARDER  = False


# Write logs to both the console and ~/evolution_myopie/suivi.log
_LOG_DIR  = os.path.join(os.path.expanduser("~"), "evolution_myopie")
os.makedirs(_LOG_DIR, exist_ok=True)
_LOG_FILE = os.path.join(_LOG_DIR, "suivi.log")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(_LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("suivi_myopie")


import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.dates as mdates
import pyodbc
import pdfplumber


# Try to import win32com to detect the active patient in Access via COM Interop.
# If pywin32 is not installed, COM auto-detection is simply disabled.
_ACCESS_FIELD_CODE   = "Code patient"
_ACCESS_FIELD_NOM    = "NOM"
_ACCESS_FIELD_PRENOM = "Prénom"

try:
    import win32com.client as _win32
    _WIN32_AVAILABLE = True
except ImportError:
    _WIN32_AVAILABLE = False
    log.warning("pywin32 not available — COM auto-detection disabled.")


def get_active_patient() -> dict | None:
    """
    Attempts to read the currently open patient from Access via COM Interop.
    Returns {"code": str, "nom": str, "prenom": str} or None if Access is
    closed, no form is active, or win32com is unavailable.
    """
    if not _WIN32_AVAILABLE:
        return None
    try:
        access = _win32.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            return None

        target = {_ACCESS_FIELD_CODE, _ACCESS_FIELD_NOM, _ACCESS_FIELD_PRENOM}
        data: dict = {}

        for i in range(form.Controls.Count):
            ctrl = form.Controls(i)
            try:
                if str(ctrl.Name) in target:
                    data[ctrl.Name] = ctrl.Value
            except Exception:
                pass

        if not target.issubset(data.keys()):
            log.debug("COM: active form found but required fields missing.")
            return None

        return {
            "code":   str(data[_ACCESS_FIELD_CODE]),
            "nom":    str(data[_ACCESS_FIELD_NOM]),
            "prenom": str(data[_ACCESS_FIELD_PRENOM]),
        }

    except Exception as e:
        log.debug(f"COM error while reading active patient: {e}")
        return None


# Poll a path until it becomes reachable or the timeout expires.
# Useful when the database / PDF folder lives on a NAS that may not be
# immediately mounted on Windows startup.
_PATH_POLL_INTERVAL = 5   # seconds between retries
_PATH_TIMEOUT       = 120 # total seconds before giving up


def wait_for_path(path: str, label: str = "") -> bool:
    """
    Waits until *path* is accessible on the filesystem (local or UNC/NAS).
    Returns True as soon as the path exists, False if *_PATH_TIMEOUT* is
    exceeded.  Logs a warning on the first failed attempt so the user knows
    the script is waiting rather than hanging silently.
    """
    p           = Path(path)
    label_str   = f"[{label}] " if label else ""
    first_check = True
    deadline    = time.monotonic() + _PATH_TIMEOUT

    while True:
        try:
            if p.exists():
                if not first_check:
                    log.info(f"{label_str}Path is now reachable: {path}")
                return True
        except Exception:
            pass  # OSError on unreachable UNC share — treat as not yet ready

        if time.monotonic() >= deadline:
            log.error(
                f"{label_str}Path unreachable after {_PATH_TIMEOUT} s — "
                f"aborting. Check NAS / network mount: {path}"
            )
            return False

        if first_check:
            log.warning(
                f"{label_str}Path not yet reachable, retrying every "
                f"{_PATH_POLL_INTERVAL} s (timeout {_PATH_TIMEOUT} s): {path}"
            )
            first_check = False

        time.sleep(_PATH_POLL_INTERVAL)


# 1. LECTURE BASE MDB

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
    log.info(f"▶ Connexion à {FICHIER_MDB}…")
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
    log.info(f"  patients={len(df_pat)}  consultations={len(df_con)}  réfractions={len(df_ref)}")
    return df_pat, df_con, df_ref

# 2. UTILITAIRES

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


# 3. CONSTRUCTION DE L'HISTORIQUE

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
        log.info(f"  [TypeRef={typeref}] {len(merged)} lignes retenues")

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

    # Keep only visits where the patient was under 25 years old at the time of consultation
    full = full[full["Age"] < 25.0]

    # Count the number of distinct visit dates per patient
    measurements_per_patient = full.groupby("CodePatient")["Date"].nunique()

    # Retain only patients who have at least 2 distinct measurement dates
    valid_patients = measurements_per_patient[measurements_per_patient >= 2].index

    # Filter the dataset to keep only those valid patients
    full = full[full["CodePatient"].isin(valid_patients)]

    full = full.sort_values(["CodePatient", "Date"]).reset_index(drop=True)

    # Build a NumConsult → CodePatient lookup from the consultation table
    consult_for_glasses = df_consult[["N° consultation", "Code patient"]].copy()
    consult_for_glasses.columns = ["NumConsult", "CodePatient"]
    consult_for_glasses["NumConsult"]  = consult_for_glasses["NumConsult"].astype(str)
    consult_for_glasses["CodePatient"] = consult_for_glasses["CodePatient"].astype(str)

    # Join the full unfiltered refraction table with the lookup above
    ref_full = df_refrac.copy()
    ref_full["NumConsult"] = ref_full["NumConsult"].astype(str)
    ref_full = ref_full.merge(consult_for_glasses, on="NumConsult", how="inner")

    # Compute the spherical equivalent for the right eye
    ref_full["SphD"] = parse_fr_float(ref_full["SphD"])
    ref_full["CylD"] = parse_fr_float(ref_full["CylD"])
    ref_full["SE_D_raw"] = calc_se(ref_full["SphD"], ref_full["CylD"])

    # Flag patients with at least one subjective refraction (TypeRef=7) and SE_D ≤ -0.5 D
    glasses_mask = (ref_full["TypeRef"].astype(str) == "7") & (ref_full["SE_D_raw"] <= -0.5)
    glasses_patients = set(ref_full[glasses_mask]["CodePatient"].astype(str).unique())

    # Print the list of retained patients
    log.info("  PATIENTS UNDER 25 WITH AT LEAST 2 MEASUREMENTS")

    for pid in valid_patients:

        # Retrieve the first row for this patient to get name info
        pat_info = full[full["CodePatient"] == str(pid)].iloc[0]

        # Get the number of distinct measurement dates for this patient
        nb_measurements = measurements_per_patient[pid]

        # Check whether this patient is inferred as a glasses wearer
        wears_glasses = "Yes" if str(pid) in glasses_patients else "Unknown/No"

        log.info(f"  ID: {pid:<10} | {pat_info['Prenom']:<10} {pat_info['NOM']:<15} | {nb_measurements} measurements | Glasses: {wears_glasses}")

    log.info(f"  Historique : {len(full)} lignes, {full['CodePatient'].nunique()} patients")
    return full


# 4. BIOMÉTRIE — extraction depuis Documents.json + PDFs

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
        log.warning(f"    ⚠  {Path(pdf_path).name} : {e}")
        return None, None


def load_biometrie(patient_ids: list[str] | None = None) -> pd.DataFrame:
    log.info(f"▶ Chargement de la biométrie depuis {FICHIER_MDB}…")
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
        log.warning(f"  ⚠  Colonne '{col}' absente de la table Documents")
        return pd.DataFrame()

    mmt = docs[docs[col].astype(str).str.contains("MMT", na=False)].copy()

    if mmt.empty:
        log.warning("  ⚠  Aucun fichier MMT trouvé dans Documents.json")
        return pd.DataFrame()

    log.info(f"  {len(mmt)} PDF(s) MMT référencé(s) dans Documents.json")

    resultats = []
    for _, row in mmt.iterrows():
        chemin_relatif = str(row[col])
        code_patient   = code_patient_depuis_chemin(chemin_relatif)
        nom_pdf        = Path(chemin_relatif).name
        date_mesure    = date_depuis_nom_fichier(nom_pdf)

        # Construire le chemin absolu
        pdf_path = Path(DOSSIER_PDF) / chemin_relatif.lstrip("\\")

        if not pdf_path.exists():
            log.warning(f"    ⚠  PDF introuvable : {pdf_path}")
            al_od, al_og = None, None
        else:
            al_od, al_og = extraire_al_pdf(str(pdf_path))
            log.info(f"    ✓  {nom_pdf}  →  OD={al_od}  OG={al_og}")

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
    log.info(f"  Biométrie : {n_match} patient(s) identifié(s), {n_al} AL extraite(s)")
    return df_al


# 5. VISUALISATION

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



# Age-stratified annual progression rates (SE in D/yr, AL in mm/yr) from the
# Correction of Myopia Evaluation Trial (COMET).
# SE: Jones-Jordan et al., Invest Ophthalmol Vis Sci 2010;51:3875-3884.
# AL: Jones-Jordan et al., Invest Ophthalmol Vis Sci 2018 (PMC6013843).
# Each tuple: (age_lower, age_upper, se_rate_D_per_yr, al_rate_mm_per_yr).
_COMET_RATES = [
    ( 6,  7, -0.75, 0.38),
    ( 8, 10, -0.55, 0.33),
    (11, 12, -0.45, 0.22),
    (13, 14, -0.30, 0.15),
    (15, 16, -0.18, 0.08),
    (17, 18, -0.10, 0.04),
    (19, 25, -0.05, 0.01),
]

# Default rates applied beyond age 25 (post-stabilisation residual drift).
_COMET_SE_DEFAULT  = -0.05
_COMET_AL_DEFAULT  =  0.01


def _comet_rates_at_age(age: float) -> tuple[float, float]:
    """Return (se_rate, al_rate) for the given age from _COMET_RATES."""
    for age_lo, age_hi, se_r, al_r in _COMET_RATES:
        if age_lo <= age <= age_hi:
            return se_r, al_r
    return _COMET_SE_DEFAULT, _COMET_AL_DEFAULT


def _build_se_reference_curve(
    date_start: "pd.Timestamp",
    age_start: float,
    se_start: float,
    date_end: "pd.Timestamp",
) -> tuple[list, list]:
    """
    Build the COMET normative SE progression curve anchored on the patient's
    first visit. Steps quarterly until date_end or age 25.
    Returns (dates, SE_values) ready for Matplotlib.
    """
    from datetime import timedelta

    dates, values = [date_start], [se_start]
    current_date, current_age, current_val = date_start, age_start, se_start

    while current_date < date_end and current_age < 25:
        se_rate, _ = _comet_rates_at_age(current_age)
        step = 0.25
        current_val  += se_rate * step
        current_age  += step
        current_date += timedelta(days=round(365.25 * step))
        dates.append(current_date)
        values.append(current_val)

    return dates, values


def _build_al_reference_curve(
    date_start: "pd.Timestamp",
    age_start: float,
    al_start: float,
    date_end: "pd.Timestamp",
) -> tuple[list, list]:
    """
    Build the COMET normative AL elongation curve anchored on the patient's
    first AL measurement. Steps quarterly until date_end or age 25.
    Returns (dates, AL_values) ready for Matplotlib.
    """
    from datetime import timedelta

    dates, values = [date_start], [al_start]
    current_date, current_age, current_val = date_start, age_start, al_start

    while current_date < date_end and current_age < 25:
        _, al_rate = _comet_rates_at_age(current_age)
        step = 0.25
        current_val  += al_rate * step
        current_age  += step
        current_date += timedelta(days=round(365.25 * step))
        dates.append(current_date)
        values.append(current_val)

    return dates, values


def plot_patient(df, code_patient, df_al=None, save=False, output_dir="."):

    pat = df[df["CodePatient"] == str(code_patient)].dropna(subset=["Date"]).copy()
    pat = pat[pat[["SE_D", "SE_G"]].notna().any(axis=1)].sort_values("Date")

    if pat.empty:
        log.warning(f"  Aucune donnée valide pour le patient {code_patient}")
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
    ymin = all_se.min() - 0.5
    ymax = all_se.max() + 0.5
    _add_severity_zones(ax, ymin, ymax)

    # Courbes SE
    for eye in ("D", "G"):
        col = f"SE_{eye}"
        sub = pat.dropna(subset=[col])
        if sub.empty:
            continue
        ax.plot(sub["Date"], sub[col], zorder=3, **STYLE[eye])

        # Label each intermediate point
        for _, row in sub.iloc[:-1].iterrows():
            ax.annotate(
                f"{row[col]:+.2f}",
                xy=(row["Date"], row[col]),
                xytext=(0, 9 if eye == "D" else -13),
                textcoords="offset points", fontsize=7,
                color=STYLE[eye]["color"], ha="center", zorder=4,
            )

        # Last point: OD label above, OG label below to avoid overlap
        last = sub.iloc[-1]
        y_offset = 22 if eye == "D" else -38
        ax.annotate(
            f"{last[col]:+.2f} D\n{_severity_label(last[col])}",
            xy=(last["Date"], last[col]),
            xytext=(10, y_offset),
            textcoords="offset points", fontsize=8,
            color=STYLE[eye]["color"],
            bbox=dict(boxstyle="round,pad=0.3", fc="#0f172a", ec=STYLE[eye]["color"], alpha=0.85),
            arrowprops=dict(arrowstyle="-", color=STYLE[eye]["color"], alpha=0.5),
        )

    # SE reference curve — COMET (Jones-Jordan et al. 2010 / 2018)
    ref_patch = None
    if "Age" in pat.columns and "DateNaissance" in pat.columns:
        first = pat.dropna(subset=["SE_D"]).iloc[0] if not pat.dropna(subset=["SE_D"]).empty else None
        if first is not None and pd.notna(first["Age"]) and pd.notna(first["SE_D"]):
            ref_dates, ref_vals = _build_se_reference_curve(
                date_start=first["Date"],
                age_start=float(first["Age"]),
                se_start=float(first["SE_D"]),
                date_end=pat["Date"].max(),
            )
            ax.plot(ref_dates, ref_vals,
                    color="#34d399", linewidth=1.5, linestyle="--",
                    alpha=0.75, zorder=2)
            ref_patch = plt.Line2D([], [], color="#34d399", linewidth=1.5,
                                   linestyle="--", alpha=0.75,
                                   label="Patient type équivalent sphérique (COMET)")

    # AL curves — connecting lines + value label on every point + COMET reference
    al_ref_patch = None
    if has_al:
        for al_col, marker, color, label in [
            ("AL_OD", "^", "#a78bfa", "Longueur axiale OD (mm)"),
            ("AL_OG", "v", "#f472b6", "Longueur axiale OG (mm)"),
        ]:
            al_sub = al_pat.dropna(subset=[al_col])
            if al_sub.empty:
                continue
            ax2.plot(al_sub["DateMesure"], al_sub[al_col],
                     marker=marker, color=color, linewidth=1.6,
                     markersize=7, linestyle="--", label=label)

            # Value label on every AL point
            for _, row in al_sub.iterrows():
                y_off = 9 if al_col == "AL_OD" else -13
                ax2.annotate(
                    f"{row[al_col]:.2f}",
                    xy=(row["DateMesure"], row[al_col]),
                    xytext=(0, y_off),
                    textcoords="offset points", fontsize=7,
                    color=color, ha="center", zorder=4,
                )

        # AL reference curve — COMET (Jones-Jordan et al. 2018)
        al_first = al_pat.dropna(subset=["AL_OD"]).sort_values("DateMesure")
        if not al_first.empty and "Age" in pat.columns:
            first_al_date = al_first.iloc[0]["DateMesure"]
            pat_with_age = pat.dropna(subset=["Age"])
            if not pat_with_age.empty:
                closest_idx = (pat_with_age["Date"] - first_al_date).abs().idxmin()
                age_at_first_al = float(pat_with_age.loc[closest_idx, "Age"])
                al_ref_dates, al_ref_vals = _build_al_reference_curve(
                    date_start=first_al_date,
                    age_start=age_at_first_al,
                    al_start=float(al_first.iloc[0]["AL_OD"]),
                    date_end=al_pat["DateMesure"].max(),
                )
                ax2.plot(al_ref_dates, al_ref_vals,
                         color="#86efac", linewidth=1.5, linestyle=":",
                         alpha=0.75, zorder=2)
                al_ref_patch = plt.Line2D([], [], color="#86efac", linewidth=1.5,
                                          linestyle=":", alpha=0.75,
                                          label="Patient type longueur axiale (COMET)")

        ax2.set_ylabel("Longueur axiale (mm)", color="#94a3b8", fontsize=10)
        ax2.tick_params(colors="#64748b", labelsize=9)
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:.2f} mm"))
        for spine in ax2.spines.values():
            spine.set_edgecolor("#1e293b")

    # Axes
    _axe_x(ax, pat["Date"].min(), pat["Date"].max())
    ax.set_ylim(ymin, ymax)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:+.1f} D"))
    ax.tick_params(colors="#64748b", labelsize=9)
    for spine in ax.spines.values():
        spine.set_edgecolor("#1e293b")

    ax.set_xlabel("Date de consultation", color="#94a3b8", fontsize=10, labelpad=8)
    ax.set_ylabel("Équivalent sphérique (dioptries)", color="#94a3b8", fontsize=10, labelpad=8)

    title = f"Évolution de la myopie — {prenom} {nom}".strip()
    if "Age" in pat.columns:
        ages = pat["Age"].dropna()
        ages = ages[ages > 0]
        if not ages.empty:
            title += f"  ({ages.min():.0f}→{ages.max():.0f} ans)"
    ax.set_title(title, color="#e2e8f0", fontsize=13, fontweight="bold", pad=14)

    # Légende — ordre fixe
    se_handles = [
        plt.Line2D([], [], **{k: v for k, v in STYLE["D"].items() if k != "label"},
                   label="Œil droit (SE_D)"),
        plt.Line2D([], [], **{k: v for k, v in STYLE["G"].items() if k != "label"},
                   label="Œil gauche (SE_G)"),
    ]
    zone_handles = [
        mpatches.Patch(color=c, alpha=0.35, label=lbl)
        for _, _, c, lbl in SEVERITY_ZONES
    ]
    al_handles = []
    if has_al:
        for al_col, marker, color, label in [
            ("AL_OD", "^", "#a78bfa", "Longueur axiale OD (mm)"),
            ("AL_OG", "v", "#f472b6", "Longueur axiale OG (mm)"),
        ]:
            if not al_pat.dropna(subset=[al_col]).empty:
                al_handles.append(plt.Line2D([], [], color=color, marker=marker,
                                             linewidth=1.6, markersize=7,
                                             linestyle="--", label=label))
    ref_handles = []
    if ref_patch:
        ref_handles.append(ref_patch)
    if al_ref_patch:
        ref_handles.append(al_ref_patch)

    all_handles = se_handles + zone_handles + al_handles + ref_handles
    ax.legend(all_handles, [h.get_label() for h in all_handles],
              loc="lower left", fontsize=8, framealpha=0.25,
              facecolor="#0f172a", edgecolor="#1e293b", labelcolor="#94a3b8")

    # Stats
    seD    = pat["SE_D"].dropna()
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
        log.info(f"  Sauvegardé : {outpath}")
    else:
        plt.show()
    plt.close(fig)

    pat = df[df["CodePatient"] == str(code_patient)].dropna(subset=["Date"]).copy()
    pat = pat[pat[["SE_D", "SE_G"]].notna().any(axis=1)].sort_values("Date")

    if pat.empty:
        log.warning(f"  Aucune donnée valide pour le patient {code_patient}")
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
    ymin = all_se.min() - 0.5
    ymax = all_se.max() + 0.5
    _add_severity_zones(ax, ymin, ymax)

    # Courbes SE
    for eye in ("D", "G"):
        col = f"SE_{eye}"
        sub = pat.dropna(subset=[col])
        if sub.empty: continue
        ax.plot(sub["Date"], sub[col], zorder=3, **STYLE[eye])

        # Label each intermediate point
        for _, row in sub.iloc[:-1].iterrows():
            ax.annotate(
                f"{row[col]:+.2f}",
                xy=(row["Date"], row[col]),
                xytext=(0, 9 if eye == "D" else -13),
                textcoords="offset points", fontsize=7,
                color=STYLE[eye]["color"], ha="center", zorder=4,
            )

        # Last point: OD label above, OG label below to avoid overlap
        last = sub.iloc[-1]
        y_offset = 22 if eye == "D" else -38
        ax.annotate(
            f"{last[col]:+.2f} D\n{_severity_label(last[col])}",
            xy=(last["Date"], last[col]),
            xytext=(10, y_offset),
            textcoords="offset points", fontsize=8,
            color=STYLE[eye]["color"],
            bbox=dict(boxstyle="round,pad=0.3", fc="#0f172a", ec=STYLE[eye]["color"], alpha=0.85),
            arrowprops=dict(arrowstyle="-", color=STYLE[eye]["color"], alpha=0.5),
        )

    zone_patches = [
        mpatches.Patch(color=c, alpha=0.35, label=lbl)
        for _, _, c, lbl in SEVERITY_ZONES
    ]

    # SE reference curve — COMET (Jones-Jordan et al. 2010 / 2018)
    ref_patch = None
    if "Age" in pat.columns and "DateNaissance" in pat.columns:
        first = pat.dropna(subset=["SE_D"]).iloc[0] if not pat.dropna(subset=["SE_D"]).empty else None
        if first is not None and pd.notna(first["Age"]) and pd.notna(first["SE_D"]):
            ref_dates, ref_vals = _build_se_reference_curve(
                date_start=first["Date"],
                age_start=float(first["Age"]),
                se_start=float(first["SE_D"]),
                date_end=pat["Date"].max(),
            )
            ax.plot(ref_dates, ref_vals,
                    color="#34d399", linewidth=1.5, linestyle="--",
                    alpha=0.75, zorder=2, label="Patient type équivalent sphérique (COMET)")
            ref_patch = plt.Line2D([], [], color="#34d399", linewidth=1.5,
                                   linestyle="--", alpha=0.75,
                                   label="Patient type équivalent sphérique (COMET)")

    # AL curves — connecting lines + value label on every point + COMET reference
    al_ref_patch = None
    if has_al:
        for al_col, marker, color, label in [
            ("AL_OD", "^", "#a78bfa", "Longueur axiale OD (mm)"),
            ("AL_OG", "v", "#f472b6", "Longueur axiale OG (mm)"),
        ]:
            al_sub = al_pat.dropna(subset=[al_col])
            if al_sub.empty:
                continue
            ax2.plot(al_sub["DateMesure"], al_sub[al_col],
                     marker=marker, color=color, linewidth=1.6,
                     markersize=7, linestyle="--", label=label)

            # Value label on every AL point
            for _, row in al_sub.iterrows():
                y_off = 9 if al_col == "AL_OD" else -13
                ax2.annotate(
                    f"{row[al_col]:.2f}",
                    xy=(row["DateMesure"], row[al_col]),
                    xytext=(0, y_off),
                    textcoords="offset points", fontsize=7,
                    color=color, ha="center", zorder=4,
                )

        # AL reference curve — COMET (Jones-Jordan et al. 2018)
        # Age is read from the SE dataframe (pat) matched on the closest date,
        # so no Age column is required in df_al.
        al_first = al_pat.dropna(subset=["AL_OD"]).sort_values("DateMesure")
        if not al_first.empty and "Age" in pat.columns:
            first_al_date = al_first.iloc[0]["DateMesure"]
            # Find the SE visit closest to the first AL measurement to get age.
            pat_with_age = pat.dropna(subset=["Age"])
            if not pat_with_age.empty:
                closest_idx = (pat_with_age["Date"] - first_al_date).abs().idxmin()
                age_at_first_al = float(pat_with_age.loc[closest_idx, "Age"])
                al_ref_dates, al_ref_vals = _build_al_reference_curve(
                    date_start=first_al_date,
                    age_start=age_at_first_al,
                    al_start=float(al_first.iloc[0]["AL_OD"]),
                    date_end=al_pat["DateMesure"].max(),
                )
                ax2.plot(al_ref_dates, al_ref_vals,
                         color="#86efac", linewidth=1.5, linestyle=":",
                         alpha=0.75, zorder=2,
                         label="Patient type longueur axiale (COMET)")
                al_ref_patch = plt.Line2D([], [], color="#86efac", linewidth=1.5,
                                          linestyle=":", alpha=0.75,
                                          label="Patient type longueur axiale (COMET)")

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
    extra = zone_patches + ([ref_patch] if ref_patch else [])
    extra_labels = [p.get_label() for p in extra]
    if has_al:
        h2, l2 = ax2.get_legend_handles_labels()
        al_extra = [al_ref_patch] if al_ref_patch else []
        ax.legend(handles + extra + h2 + al_extra,
                  labels_leg + extra_labels + l2 + [p.get_label() for p in al_extra],
                  loc="lower left", fontsize=8, framealpha=0.25,
                  facecolor="#0f172a", edgecolor="#1e293b", labelcolor="#94a3b8")
    else:
        ax.legend(handles + extra,
                  labels_leg + extra_labels,
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
        log.info(f"  Sauvegardé : {outpath}")
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
        log.info(f"  Sauvegardé : {outpath}")
    else:
        plt.show()
    plt.close(fig)


# 6. SÉLECTION INTERACTIVE

def choisir_patients(df: pd.DataFrame) -> list[str]:
    tous  = sorted(df["CodePatient"].unique())
    total = len(tous)

    ref = (
        df[["CodePatient", "NOM", "Prenom"]]
        .drop_duplicates("CodePatient")
        .set_index("CodePatient")
    )

    while True:
        saisie = input("  Votre choix : ").strip()

        # Entrée vide → tous
        if saisie == "":
            print(f"  → Tous les patients ({total})")
            return [str(p) for p in tous]

        # Recherche par nom/prénom
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

        # Saisie par ID
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


# 7. POINT D'ENTRÉE

def main():
    log.info("=" * 70)
    log.info("  Suivi myopie — démarrage")
    log.info(f"  Log : {_LOG_FILE}")
    log.info("=" * 70)

    # -------------------------------------------------------------------------
    # PRODUCTION PATTERN 3 — Path availability guard
    # Check that FICHIER_MDB and DOSSIER_PDF are reachable before doing anything
    # else.  Useful when the database or PDF folder lives on a NAS that may take
    # a few seconds to mount on Windows startup.
    # -------------------------------------------------------------------------
    if not wait_for_path(FICHIER_MDB, label="MDB"):
        log.error("Cannot reach FICHIER_MDB — aborting.")
        sys.exit(1)

    if not wait_for_path(DOSSIER_PDF, label="PDF"):
        log.error("Cannot reach DOSSIER_PDF — aborting.")
        sys.exit(1)
    # -------------------------------------------------------------------------

    # -------------------------------------------------------------------------
    # PRODUCTION PATTERN 2 — COM auto-detection of the active patient
    # If Access is open and a patient record is displayed, bypass the
    # interactive menu and generate the chart directly for that patient.
    # Falls back to choisir_patients() if Access is closed or no patient is
    # open.
    # -------------------------------------------------------------------------
    ids_config = [str(i) for i in PATIENT_IDS] if PATIENT_IDS else None
    active_patient = get_active_patient()

    if active_patient:
        log.info(
            f"COM auto-detection: patient actif → "
            f"{active_patient['prenom']} {active_patient['nom']} "
            f"(code {active_patient['code']})"
        )
        ids_config  = [active_patient["code"]]
        auto_detect = True
    else:
        if _WIN32_AVAILABLE:
            log.info("COM auto-detection: aucun patient ouvert dans Access — menu interactif.")
        auto_detect = False
    # -------------------------------------------------------------------------

    # Chargement depuis MDB
    df_pat, df_con, df_ref = load_all_tables(patient_ids=ids_config)

    log.info("▶ Construction de l'historique…")
    df = build_history(df_pat, df_con, df_ref, typeref=TYPEREF)

    # Biométrie filtrée sur les mêmes patients
    log.info("▶ Chargement de la biométrie…")
    df_al = load_biometrie(patient_ids=ids_config)

    # Sélection
    if auto_detect:
        # Patient résolu automatiquement via COM — on garde ids_config tel quel
        ids = ids_config
        log.info(f"▶ Patient détecté automatiquement : {', '.join(ids)}")
    elif PATIENT_IDS and not auto_detect:
        ids = ids_config
        log.info(f"▶ Patient(s) configuré(s) : {', '.join(ids)}")
    else:
        ids = choisir_patients(df)
        # Si sélection interactive → recharger la bio pour ces patients seulement
        df_al = load_biometrie(patient_ids=ids)

    # Visualisation
    print()
    if MODE_COHORTE:
        log.info(f"▶ Courbe cohorte ({len(ids)} patients, œil {OEIL_COHORTE})…")
        plot_cohort(df, patient_ids=ids, eye=OEIL_COHORTE, save=SAUVEGARDER)
    else:
        log.info(f"▶ Tracé individuel de {len(ids)} patient(s)…")
        for pid in ids:
            plot_patient(df, pid, df_al=df_al, save=SAUVEGARDER)

    log.info("✓ Terminé.")


if __name__ == "__main__":
    main()