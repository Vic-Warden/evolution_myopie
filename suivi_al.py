"""
Suivi longitudinal de la longueur axiale
==========================================
Visualisation de l'AL (Comp. AL) à partir de biometrie_extraite.json

Dépendances : pip install pandas matplotlib
"""

import json
from pathlib import Path

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# ═════════════════════════════════════════════════════════════════════════════
# ▶▶  CONFIGURATION
# ═════════════════════════════════════════════════════════════════════════════

FICHIER_BIOMETRIE = "biometrie_extraite.json"
SAUVEGARDER       = False   # True = PNG sur disque, False = fenêtre interactive

# ═════════════════════════════════════════════════════════════════════════════


def load_biometrie() -> pd.DataFrame:
    with open(FICHIER_BIOMETRIE, encoding="utf-8") as f:
        df = pd.DataFrame(json.load(f))
    df["DateMesure"]    = pd.to_datetime(df["DateMesure"],    errors="coerce")
    df["DateNaissance"] = pd.to_datetime(df["DateNaissance"], errors="coerce")
    df["NOM_norm"]    = df["NOM"].str.split(",").str[0].str.strip().str.upper()
    df["PRENOM_norm"] = df["NOM"].str.split(",").str[1].str.strip().str.upper().fillna("")
    return df.sort_values("DateMesure").reset_index(drop=True)


def choisir_patient(df: pd.DataFrame) -> tuple[str, str, pd.Timestamp]:
    patients = (
        df[["NOM_norm", "PRENOM_norm", "DateNaissance"]]
        .drop_duplicates()
        .sort_values(["NOM_norm", "PRENOM_norm"])
        .reset_index(drop=True)
    )
    total = len(patients)

    print()
    print("┌─────────────────────────────────────────────────┐")
    print("│        SÉLECTION DU PATIENT (BIOMÉTRIE)         │")
    print("├─────────────────────────────────────────────────┤")
    print(f"│  {total} patient(s) disponible(s)                    │")
    print("└─────────────────────────────────────────────────┘")
    print()

    for i, row in patients.iterrows():
        dob = row["DateNaissance"].strftime("%d/%m/%Y") if pd.notna(row["DateNaissance"]) else "?"
        print(f"  [{i}]  {row['PRENOM_norm']} {row['NOM_norm']}  (né le {dob})")
    print()

    while True:
        saisie = input("  Votre choix (numéro) : ").strip()
        try:
            idx = int(saisie)
            if 0 <= idx < total:
                row = patients.loc[idx]
                return row["NOM_norm"], row["PRENOM_norm"], row["DateNaissance"]
        except ValueError:
            pass
        print(f"  ⚠  Entrez un numéro entre 0 et {total - 1}")

def plot_al(df: pd.DataFrame, nom_norm: str, prenom_norm: str, dob: pd.Timestamp) -> None:
    pat = df[
        (df["NOM_norm"] == nom_norm) &
        (df["PRENOM_norm"] == prenom_norm) &
        (df["DateNaissance"] == dob)
    ].sort_values("DateMesure")
    if pat.empty:
        print(f"  Aucune donnée pour {nom_norm}")
        return

    nom_affiche = pat["NOM"].iloc[0]
    dob = pat["DateNaissance"].iloc[0]
    dob_str = dob.strftime("%d/%m/%Y") if pd.notna(dob) else "?"
    n = len(pat)

    # ── Figure ───────────────────────────────────────────────────────────────
    fig, ax = plt.subplots(figsize=(12, 5))
    fig.patch.set_facecolor("#0f172a")
    ax.set_facecolor("#111827")
    ax.grid(True, color="#1e293b", linewidth=0.8, zorder=0)
    ax.set_axisbelow(True)

    # Ligne de référence normale (22–24 mm = globe normal)
    ax.axhspan(22.0, 24.0, color="#10b981", alpha=0.05, zorder=0, label="Globe normal (22–24 mm)")
    ax.axhline(23.0, color="#10b981", linestyle=":", linewidth=1.0, alpha=0.4)

    # Courbes OD et OG
    od = pat.dropna(subset=["AL_OD"])
    og = pat.dropna(subset=["AL_OG"])

    ax.plot(od["DateMesure"], od["AL_OD"],
            marker="^", color="#a78bfa", linewidth=2.0, markersize=7,
            label="Longueur axiale OD (mm)")
    ax.plot(og["DateMesure"], og["AL_OG"],
            marker="v", color="#f472b6", linewidth=2.0, markersize=7,
            label="Longueur axiale OG (mm)")

    # Annotation tous les points
    for col, color, offset in [("AL_OD", "#a78bfa", 8), ("AL_OG", "#f472b6", -18)]:
        sub = pat.dropna(subset=[col])
        if sub.empty:
            continue
        for _, row in sub.iterrows():
            ax.annotate(
                f"{row[col]:.2f}",
                xy=(row["DateMesure"], row[col]),
                xytext=(0, offset),
                textcoords="offset points",
                fontsize=7, color=color,
                bbox=dict(boxstyle="round,pad=0.2", fc="#0f172a", ec=color, alpha=0.7),
                ha="center",
            )

    # ── Axe X ────────────────────────────────────────────────────────────────
    if n > 1:
        date_range = (pat["DateMesure"].max() - pat["DateMesure"].min()).days
        if date_range > 365 * 10:
            ax.xaxis.set_major_locator(mdates.YearLocator(5))
        elif date_range > 365 * 3:
            ax.xaxis.set_major_locator(mdates.YearLocator())
        elif date_range > 180:
            ax.xaxis.set_major_locator(mdates.MonthLocator(bymonth=[1, 4, 7, 10]))
        else:
            ax.xaxis.set_major_locator(mdates.MonthLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y" if date_range <= 365 * 3 else "%Y"))

    plt.setp(ax.xaxis.get_majorticklabels(), rotation=35, ha="right", fontsize=9)

    # ── Axe Y ────────────────────────────────────────────────────────────────
    all_al = pd.concat([od["AL_OD"], og["AL_OG"]]).dropna()
    ax.set_ylim(all_al.min() - 0.3, all_al.max() + 0.3)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:.2f} mm"))
    ax.tick_params(colors="#64748b", labelsize=9)
    for spine in ax.spines.values():
        spine.set_edgecolor("#1e293b")

    # ── Labels ───────────────────────────────────────────────────────────────
    ax.set_xlabel("Date de mesure", color="#94a3b8", fontsize=10, labelpad=8)
    ax.set_ylabel("Longueur axiale (mm)", color="#94a3b8", fontsize=10, labelpad=8)
    ax.set_title(
        f"Évolution de la longueur axiale — {nom_affiche}  (né le {dob_str})",
        color="#e2e8f0", fontsize=13, fontweight="bold", pad=14
    )

    # ── Légende ──────────────────────────────────────────────────────────────
    ax.legend(fontsize=8, framealpha=0.25, facecolor="#0f172a",
              edgecolor="#1e293b", labelcolor="#94a3b8", loc="upper left")

    # ── Stats ─────────────────────────────────────────────────────────────────
    if not od.empty and len(od) > 1:
        delta_od = od["AL_OD"].iloc[-1] - od["AL_OD"].iloc[0]
        duree = (od["DateMesure"].max() - od["DateMesure"].min()).days / 365.25
        stats_txt = (
            f"n={n} mesure(s)  |  durée {duree:.1f} ans\n"
            f"OD : {od['AL_OD'].min():.2f} → {od['AL_OD'].max():.2f} mm  (Δ {delta_od:+.2f} mm)"
        )
        ax.text(0.99, 0.97, stats_txt,
                transform=ax.transAxes, ha="right", va="top",
                fontsize=8, color="#64748b", fontfamily="monospace",
                bbox=dict(boxstyle="round,pad=0.4", fc="#0f172a", ec="#1e293b", alpha=0.7))

    plt.tight_layout()

    if SAUVEGARDER:
        outpath = Path(f"al_{nom_norm}.png")
        fig.savefig(outpath, dpi=150, bbox_inches="tight", facecolor=fig.get_facecolor())
        print(f"  Sauvegardé : {outpath}")
    else:
        plt.show()

    plt.close(fig)


def main():
    if not Path(FICHIER_BIOMETRIE).exists():
        raise FileNotFoundError(f"Fichier introuvable : {FICHIER_BIOMETRIE}")

    print(f"▶ Chargement de {FICHIER_BIOMETRIE}…")
    df = load_biometrie()
    print(f"  {len(df)} mesure(s), {df['NOM_norm'].nunique()} patient(s)")
    

    nom, prenom, dob = choisir_patient(df)
    print(f"\n▶ Tracé pour {nom}…")
    plot_al(df, nom, prenom, dob)
    print("✓ Terminé.")


if __name__ == "__main__":
    main()
