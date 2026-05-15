# Évolution Myopie - Suivi Longitudinal de la Longueur Axiale

Ensemble de scripts Python pour l'extraction, l'analyse et la visualisation de données ophtalmologiques, spécifiquement dédiés au suivi de la myopie et de l'allongement axial de l'œil.

## Structure du projet

```
evolution_myopie/
├── suivi_myopie.py         ← Single entry point: extraction, analysis and visualisation
├── requirements.txt
└── README.md
```

> **Note:** `config.py`, `extraire_biometrie.py`, `suivi_al.py` and `find_patient/patient.py` have been
> consolidated into `suivi_myopie.py`. All configuration (paths, patient IDs, flags) is now set via
> constants at the top of that file.

## Vue d'ensemble

Ce projet traite trois étapes clés du suivi ophtalmologique :

1. **Extraction** : récupération de données biométriques depuis des PDFs
2. **Jointure** : fusion des données biométriques avec un dossier ophtalmologique
3. **Visualisation** : graphes longitudinaux de l'évolution de la myopie et de la longueur axiale

---

## Scripts

### `suivi_myopie.py`
**Single consolidated script — all-in-one pipeline**

Connects directly to the Access database (`PUBLIC.MDB`) and the biometry PDF folder, then:
- Reads refractions from `tREFRACTION`, joined with `Consultation` and `Patients`
- Filters patients under 25 years old with at least 2 distinct measurement dates
- Extracts axial length (AL) from `MMT-Full_*.pdf` files via the `Documents` table
- Plots individual longitudinal SE + AL curves with COMET normative reference curves
- Auto-detects the active patient from the open Access form via COM (Windows only)
- Waits for NAS/network paths to become available before proceeding
- Saves charts to PNG if `SAUVEGARDER = True`

**Key configuration constants (top of file):**
| Constant | Description |
|---|---|
| `FICHIER_MDB` | Path to `PUBLIC.MDB` |
| `DOSSIER_PDF` | Folder containing the biometry PDFs |
| `PATIENT_IDS` | Force specific patient IDs (or `None` for interactive menu) |
| `TYPEREF` | Refraction type filter (`7` = subjective, `6` = auto, `None` = all) |
| `SAUVEGARDER` | `True` = save PNG, `False` = display |

**Dependencies:** `pandas`, `matplotlib`, `pyodbc`, `pdfplumber`, `pywin32` (optional, COM only)

---

## Flux de données

```
PUBLIC.MDB (Access)
  ├── tREFRACTION + Consultation + Patients
  │          ↓
  │    Spherical equivalent (SE) curves per patient
  │
  └── Documents → MMT-Full_*.pdf paths
             ↓
       pdfplumber → Axial length (AL) per eye
             ↓
    suivi_myopie.py
             ↓
    Individual chart (SE + AL + COMET reference)
```

---

## Installation

### Prérequis
- Python 3.10+

### Dépendances
```bash
pip install -r requirements.txt
```

### Setup
1. Cloner le dépôt
2. Installer les dépendances (voir ci-dessus)
3. Install the [Microsoft Access Database Engine 2016 (64-bit)](https://www.microsoft.com/en-us/download/details.aspx?id=54920) (Windows only)
4. Set `FICHIER_MDB` and `DOSSIER_PDF` at the top of `suivi_myopie.py`
5. Run the script

---

## Utilisation

```bash
python suivi_myopie.py
```

- **Interactive menu:** leave `PATIENT_IDS = None` → search by name or ID
- **Force a patient:** set `PATIENT_IDS = [1758507609]`
- **COM auto-detection (Windows):** if Access is open with a patient record, the chart is generated automatically without any prompt
- **Save PNG:** set `SAUVEGARDER = True` → files written next to the script

---

## Formats de données

The script reads directly from the Access database — no intermediate CSV/JSON files are needed.

### Tables Access utilisées
| Table | Colonnes clés |
|---|---|
| `Patients` | `Code patient`, `NOM`, `Prénom`, `Date de Naissance` |
| `Consultation` | `N° consultation`, `Code patient`, `Date` |
| `tREFRACTION` | `NumConsult`, `SphD`, `CylD`, `SphG`, `CylG`, `TypeRef` |
| `Documents` | `code patient`, `Photo externe` (PDF path) |

---

## Cas d'usage

- **Suivi myopique** : visualiser la progression de la myopie chez un patient
- **Analyse AL** : détecter l'allongement axial anormal (facteur clé de la myopie)
- **Génération de rapports** : exporter les graphes en PNG

---

## Dépannage

| Problème | Solution |
|----------|----------|
| Aucun PDF trouvé | Vérifier `DOSSIER_PDF` et le pattern `MMT-Full_*.pdf` |
| Graphe vide | Vérifier les dates et que `DateMesure` n'est pas null |
| Dates mal parsées | Adapter `_parse_dob()` au format de vos données |
| Cannot connect to MDB | Install Access Database Engine 2016 (64-bit) and check `FICHIER_MDB` path |
| NAS/network path unreachable | Script retries for 120 s — check network mount |
| COM auto-detection not working | Install `pywin32`; Access must be open with a patient record displayed |

---

## Licences & Attributions

Project créé pour le suivi ophtalmologique en recherche.

## Contact

Pour toute question ou amélioration, merci de ouvrir une issue.

---

**Dernière mise à jour :** Mai 2026
