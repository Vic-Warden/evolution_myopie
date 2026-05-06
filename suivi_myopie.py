"""
Suivi longitudinal de la myopie
================================
Visualisation de l'équivalent sphérique à partir de fichiers JSON
issus d'une base ophtalmologique (Patient / Consultation / tREFRACTION).

Dépendances : pip install pandas matplotlib
"""

import json
from pathlib import Path

from config import (
    FICHIER_PATIENTS, FICHIER_CONSULTATIONS, FICHIER_REFRACTION,
    FICHIER_JSON_BIOMETRIE, OUTPUT_GRAPHS, SAUVEGARDER,
    PATIENT_IDS, TYPEREF, MODE_COHORTE, OEIL_COHORTE,
)

# FICHIER_PATIENTS      → data/json/Patients.json
# FICHIER_CONSULTATIONS → data/json/Consultation.json
# FICHIER_REFRACTION    → data/json/tREFRACTION.json
# FICHIER_BIOMETRIE     → output/csv/biometrie_extraite.json  (optionnel)

FICHIER_BIOMETRIE = FICHIER_JSON_BIOMETRIE

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.dates as mdates

# 1. UTILITAIRES

def parse_fr_float(series: pd.Series) -> pd.Series:
    """Convertit les nombres au format français ('-3,50') en float."""
    return (
        series.astype(str)
        .str.strip()
        .str.replace(",", ".", regex=False)
        .apply(pd.to_numeric, errors="coerce")
    )

def calc_se(sph: pd.Series, cyl: pd.Series) -> pd.Series:
    """Équivalent sphérique : SE = Sph + Cyl / 2"""
    return sph + cyl / 2.0

# 2. CHARGEMENT & JOINTURES

def load_json(path: str) -> pd.DataFrame:
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    return pd.DataFrame(data)

def _parse_dob(val) -> "pd.Timestamp":
    import re
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

def build_history(
    df_patients: pd.DataFrame,
    df_consult: pd.DataFrame,
    df_refrac: pd.DataFrame,
    typeref: int | None = None,
) -> pd.DataFrame:
    """
    Construit le DataFrame historique réfractionnel complet.
    Jointures : tREFRACTION → Consultation → Patient
    """
    consult_slim = df_consult[["N° consultation", "Code patient", "Date"]].copy()
    consult_slim.columns = ["NumConsult", "CodePatient", "Date"]
    consult_slim["NumConsult"]  = consult_slim["NumConsult"].astype(str)
    consult_slim["CodePatient"] = consult_slim["CodePatient"].astype(str)

    df_refrac = df_refrac.copy()
    df_refrac["NumConsult"] = df_refrac["NumConsult"].astype(str)

    merged = df_refrac.merge(consult_slim, on="NumConsult", how="left")

    if typeref is not None:
        merged = merged[merged["TypeRef"].astype(str) == str(typeref)]
        print(f"  [TypeRef={typeref}] {len(merged)} lignes retenues")

    pat_slim = df_patients[["Code patient", "NOM", "Prénom", "Date de Naissance"]].copy()
    pat_slim.columns = ["CodePatient", "NOM", "Prenom", "DateNaissance"]
    pat_slim["CodePatient"] = pat_slim["CodePatient"].astype(str)

    merged["CodePatient"] = merged["CodePatient"].astype(str)
    full = merged.merge(pat_slim, on="CodePatient", how="left")

    full["Date"] = full["Date"].apply(_parse_dob)
    for col in ["SphD", "CylD", "SphG", "CylG"]:
        full[col] = parse_fr_float(full[col])

    full["SE_D"] = calc_se(full["SphD"], full["CylD"])
    full["SE_G"] = calc_se(full["SphG"], full["CylG"])

    full["DateNaissance"] = full["DateNaissance"].apply(_parse_dob)
    full["Age"] = ((full["Date"] - full["DateNaissance"]).dt.days / 365.25).round(1)

    full = full.sort_values(["CodePatient", "Date"]).reset_index(drop=True)
    print(f"  Historique construit : {len(full)} lignes, {full['CodePatient'].nunique()} patients")
    return full

def load_biometrie(df_history: pd.DataFrame) -> pd.DataFrame:
    """
    Charge le JSON de biométrie et fait le lien avec l'historique
    via NOM + DateNaissance.
    """
    if not FICHIER_BIOMETRIE or not Path(FICHIER_BIOMETRIE).exists():
        print("  ⚠  Pas de fichier biométrie — courbe AL désactivée")
        return pd.DataFrame()

    df_al = pd.DataFrame(json.load(open(FICHIER_BIOMETRIE, encoding="utf-8")))
    df_al["DateMesure"]    = pd.to_datetime(df_al["DateMesure"],    errors="coerce")
    df_al["DateNaissance"] = pd.to_datetime(df_al["DateNaissance"], errors="coerce")

    # Normaliser le NOM : "Doe, John" → "DOE" pour matcher "NOM" de la base
    df_al["NOM_norm"] = (
        df_al["NOM"]
        .str.split(",").str[0]
        .str.strip()
        .str.upper()
    )

    # Côté historique : un CodePatient → une DateNaissance + NOM
    ref = (
        df_history[["CodePatient", "NOM", "DateNaissance"]]
        .drop_duplicates("CodePatient")
        .copy()
    )
    ref["NOM_norm"] = ref["NOM"].str.strip().str.upper()

    # Jointure sur NOM_norm + DateNaissance
    df_al = df_al.merge(
        ref[["CodePatient", "NOM_norm", "DateNaissance"]],
        on=["NOM_norm", "DateNaissance"],
        how="left"
    )

    n_match = df_al["CodePatient"].notna().sum()
    print(f"  Biométrie : {len(df_al)} mesure(s), {n_match} liée(s) à un patient")
    return df_al

# 3. VISUALISATION

SEVERITY_ZONES = [
    (0,   -3,  "#f59e0b", "Myopie faible (0→−3 D)"),
    (-3,  -6,  "#f97316", "Myopie moyenne (−3→−6 D)"),
    (-6,  -20, "#ef4444", "Myopie forte (< −6 D)"),
]

STYLE = {
    "D": dict(color="#2563eb", marker="o", linewidth=2.2, markersize=6, label="Œil droit (SE_D)"),
    "G": dict(color="#dc2626", marker="s", linewidth=2.2, markersize=6, label="Œil gauche (SE_G)"),
}

def _add_severity_zones(ax: plt.Axes, ymin: float, ymax: float) -> None:
    for top, bot, color, _ in SEVERITY_ZONES:
        ax.axhspan(
            max(bot, ymin - 1), min(top, ymax + 1),
            color=color, alpha=0.07, zorder=0
        )
    ax.axhline(0, color="#94a3b8", linestyle="--", linewidth=1.0, alpha=0.7, zorder=1)

def _severity_label(se: float) -> str:
    if se is None or pd.isna(se): return ""
    if se > 0.5:   return "Hypermétrope"
    if se >= -0.5: return "Emmétrope"
    if se >= -3:   return "Myopie faible"
    if se >= -6:   return "Myopie moyenne"
    return "Myopie forte"

def plot_patient(
    df: pd.DataFrame,
    code_patient: str,
    df_al: pd.DataFrame = None,
    save: bool = False,
) -> None:
    """Trace l'évolution de l'équivalent sphérique pour un patient."""
    pat = df[df["CodePatient"] == str(code_patient)].dropna(subset=["Date"]).copy()
    pat = pat[pat[["SE_D", "SE_G"]].notna().any(axis=1)].sort_values("Date")

    if pat.empty:
        print(f"  Aucune donnée valide pour le patient {code_patient}")
        return

    nom    = pat["NOM"].iloc[0] if "NOM" in pat.columns else "?"
    prenom = pat["Prenom"].iloc[0] if "Prenom" in pat.columns else ""
    n_pts  = len(pat)

    if df_al is not None and not df_al.empty:
        al_pat = df_al[df_al["CodePatient"] == str(code_patient)].dropna(subset=["DateMesure"])
    else:
        al_pat = pd.DataFrame()

    has_al = not al_pat.empty

    fig, ax = plt.subplots(figsize=(13, 6))
    ax2 = ax.twinx() if has_al else None
    fig.patch.set_facecolor("#0f172a")
    ax.set_facecolor("#111827")

    ax.grid(True, color="#1e293b", linewidth=0.8, zorder=0)
    ax.set_axisbelow(True)

    # Zones de sévérité
    all_se = pd.concat([pat["SE_D"].dropna(), pat["SE_G"].dropna()])
    ymin = min(-0.5, all_se.min()) - 0.5
    ymax = max(0.5,  all_se.max()) + 0.5
    _add_severity_zones(ax, ymin, ymax)

    # Courbes SE
    for eye in ("D", "G"):
        col = f"SE_{eye}"
        sub = pat.dropna(subset=[col])
        if sub.empty:
            continue
        ax.plot(sub["Date"], sub[col], zorder=3, **STYLE[eye])
        last = sub.iloc[-1]
        label_txt = f"{last[col]:+.2f} D\n{_severity_label(last[col])}"
        ax.annotate(
            label_txt,
            xy=(last["Date"], last[col]),
            xytext=(10, 8 if eye == "D" else -22),
            textcoords="offset points",
            fontsize=8,
            color=STYLE[eye]["color"],
            bbox=dict(boxstyle="round,pad=0.3", fc="#0f172a", ec=STYLE[eye]["color"], alpha=0.85),
            arrowprops=dict(arrowstyle="-", color=STYLE[eye]["color"], alpha=0.5),
        )

    # Zone patches — défini avant le bloc AL pour être disponible partout
    zone_patches = [
        mpatches.Patch(color=c, alpha=0.35, label=lbl)
        for _, _, c, lbl in SEVERITY_ZONES
    ]

    # Courbes AL sur axe droit
    if has_al:
        ax2.plot(
            al_pat["DateMesure"], al_pat["AL_OD"],
            marker="^", color="#a78bfa", linewidth=1.8,
            markersize=7, linestyle="--", label="Longueur axiale OD (mm)"
        )
        ax2.plot(
            al_pat["DateMesure"], al_pat["AL_OG"],
            marker="v", color="#f472b6", linewidth=1.8,
            markersize=7, linestyle="--", label="Longueur axiale OG (mm)"
        )
        ax2.set_ylabel("Longueur axiale (mm)", color="#94a3b8", fontsize=10)
        ax2.tick_params(colors="#64748b", labelsize=9)
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:.2f} mm"))
        for spine in ax2.spines.values():
            spine.set_edgecolor("#1e293b")

    date_range = (pat["Date"].max() - pat["Date"].min()).days

    if date_range > 365 * 10:
        ax.xaxis.set_major_locator(mdates.YearLocator(5))   # une graduation tous les 5 ans
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    elif date_range > 365 * 3:
        ax.xaxis.set_major_locator(mdates.YearLocator())    # une par an
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    elif date_range > 180:
        ax.xaxis.set_major_locator(mdates.MonthLocator(bymonth=[1, 4, 7, 10]))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    else:
        ax.xaxis.set_major_locator(mdates.MonthLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))

    ax.xaxis.set_minor_locator(mdates.AutoDateLocator())
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=35, ha="right", fontsize=9)

    ax.set_ylim(ymin, ymax)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:+.1f} D"))
    ax.tick_params(colors="#64748b", labelsize=9)
    for spine in ax.spines.values():
        spine.set_edgecolor("#1e293b")

    ax.set_xlabel("Date de consultation", color="#94a3b8", fontsize=10, labelpad=8)
    ax.set_ylabel("Équivalent sphérique (dioptries)", color="#94a3b8", fontsize=10, labelpad=8)

    title = f"Évolution de la myopie — {prenom} {nom}".strip()
    if "Age" in pat.columns:
        ages_valides = pat["Age"].dropna()
        ages_valides = ages_valides[ages_valides > 0]
        if not ages_valides.empty:
            title += f"  ({ages_valides.min():.0f}→{ages_valides.max():.0f} ans)"
    ax.set_title(title, color="#e2e8f0", fontsize=13, fontweight="bold", pad=14)

    handles, labels_leg = ax.get_legend_handles_labels()

    if has_al:
        h2, l2 = ax2.get_legend_handles_labels()
        ax.legend(
            handles + zone_patches + h2,
            labels_leg + [p.get_label() for p in zone_patches] + l2,
            loc="lower left", fontsize=8,
            framealpha=0.25, facecolor="#0f172a",
            edgecolor="#1e293b", labelcolor="#94a3b8",
        )
    else:
        ax.legend(
            handles + zone_patches,
            labels_leg + [p.get_label() for p in zone_patches],
            loc="lower left", fontsize=8,
            framealpha=0.25, facecolor="#0f172a",
            edgecolor="#1e293b", labelcolor="#94a3b8",
        )

    seD = pat["SE_D"].dropna()
    seG = pat["SE_G"].dropna()
    prog_D = (seD.iloc[-1] - seD.iloc[0]) if len(seD) > 1 else float("nan")
    duree  = (pat["Date"].max() - pat["Date"].min()).days / 365.25

    stats_txt = (
        f"n={n_pts} mesures  |  durée {duree:.1f} ans\n"
        f"OD : {seD.min():+.2f} D → {seD.max():+.2f} D  (Δ {prog_D:+.2f} D)"
    )
    ax.text(
        0.99, 0.97, stats_txt,
        transform=ax.transAxes,
        ha="right", va="top",
        fontsize=8, color="#64748b",
        fontfamily="monospace",
        bbox=dict(boxstyle="round,pad=0.4", fc="#0f172a", ec="#1e293b", alpha=0.7),
    )

    plt.tight_layout()

    if save:
        OUTPUT_GRAPHS.mkdir(parents=True, exist_ok=True)
        outpath = OUTPUT_GRAPHS / f"myopie_patient_{code_patient}.png"
        fig.savefig(outpath, dpi=150, bbox_inches="tight", facecolor=fig.get_facecolor())
        print(f"  Sauvegardé : {outpath}")
    else:
        plt.show()

    plt.close(fig)

def plot_cohort(
    df: pd.DataFrame,
    patient_ids: list[str] | None = None,
    eye: str = "D",
    save: bool = False,
) -> None:
    """Superpose les courbes de plusieurs patients (cohorte)."""
    ids = patient_ids or sorted(df["CodePatient"].unique())
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
        sub = df[(df["CodePatient"] == str(pid))].dropna(subset=["Date", se_col]).sort_values("Date")
        if sub.empty:
            continue
        nom = sub["NOM"].iloc[0] if "NOM" in sub.columns else str(pid)
        prn = sub["Prenom"].iloc[0] if "Prenom" in sub.columns else ""
        ax.plot(
            sub["Date"], sub[se_col],
            marker="o", markersize=4, linewidth=1.6,
            color=color, alpha=0.85,
            label=f"{prn} {nom}".strip()
        )

    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.xaxis.set_major_locator(mdates.YearLocator())
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=35, ha="right", fontsize=9)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:+.1f} D"))
    ax.tick_params(colors="#64748b", labelsize=9)
    for spine in ax.spines.values():
        spine.set_edgecolor("#1e293b")

    ax.set_xlabel("Date", color="#94a3b8", fontsize=10)
    ax.set_ylabel(f"Équivalent sphérique œil {'droit' if eye=='D' else 'gauche'} (D)", color="#94a3b8", fontsize=10)
    ax.set_title(f"Évolution comparative de la myopie — Cohorte ({len(ids)} patients)", color="#e2e8f0", fontsize=13, fontweight="bold")
    ax.legend(fontsize=8, framealpha=0.2, facecolor="#0f172a", edgecolor="#1e293b", labelcolor="#e2e8f0", ncol=2)

    plt.tight_layout()

    if save:
        OUTPUT_GRAPHS.mkdir(parents=True, exist_ok=True)
        outpath = OUTPUT_GRAPHS / "myopie_cohorte.png"
        fig.savefig(outpath, dpi=150, bbox_inches="tight", facecolor=fig.get_facecolor())
        print(f"  Sauvegardé : {outpath}")
    else:
        plt.show()

    plt.close(fig)

# 4. POINT D'ENTRÉE

def choisir_patients(df: pd.DataFrame) -> list[str]:
    tous = sorted(df["CodePatient"].unique())
    total = len(tous)

    # Table de référence nom/prénom par patient
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
                print(f"  → {len(correspondances)} patient(s) sélectionné(s)")
                return [str(p) for p in correspondances]
            print()
            continue

        morceaux = [m.strip() for m in saisie.replace(";", ",").split(",") if m.strip()]
        ids_valides, erreurs = [], []

        for m in morceaux:
            if m in [str(p) for p in tous]:
                ids_valides.append(m)
            else:
                erreurs.append(m)

        if erreurs:
            print(f"  ⚠  Aucun résultat pour : {', '.join(erreurs)} — réessayez.")
            continue

        print(f"  → {len(ids_valides)} patient(s) sélectionné(s) : {', '.join(ids_valides)}")
        return ids_valides

def main():
    missing_files = [
        path for path in (FICHIER_PATIENTS, FICHIER_CONSULTATIONS, FICHIER_REFRACTION)
        if not Path(path).exists()
    ]
    if missing_files:
        raise FileNotFoundError(
            f"Fichiers JSON manquants : {', '.join(missing_files)}"
        )

    print("▶ Chargement des fichiers JSON…")
    df_pat = load_json(FICHIER_PATIENTS)
    df_con = load_json(FICHIER_CONSULTATIONS)
    df_ref = load_json(FICHIER_REFRACTION)
    print(f"  patients={len(df_pat)}  consultations={len(df_con)}  réfractions={len(df_ref)}")

    print(f"▶ Construction de l'historique (TypeRef={'tous' if TYPEREF is None else TYPEREF})…")
    df = build_history(df_pat, df_con, df_ref, typeref=TYPEREF)
    df_al = load_biometrie(df)

    if PATIENT_IDS:
        ids = [str(i) for i in PATIENT_IDS]
        print(f"▶ Patient(s) configuré(s) : {', '.join(ids)}")
    else:
        ids = choisir_patients(df)

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