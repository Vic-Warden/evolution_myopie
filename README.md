# Évolution Myopie - Suivi Longitudinal de la Longueur Axiale

Ensemble de scripts Python pour l'extraction, l'analyse et la visualisation de données ophtalmologiques, spécifiquement dédiés au suivi de la myopie et de l'allongement axial de l'œil.

## Structure du projet

```
evolution_myopie/
├── config.py               ← ⚙️  Configuration centralisée (chemins & paramètres)
├── extraire_biometrie.py   ← Étape 1 : extraction PDF → CSV/JSON
├── suivi_al.py             ← Étape 2 : graphe longueur axiale
├── suivi_myopie.py         ← Étape 3 : graphe réfraction longitudinale
├── requirements.txt
├── data/
│   ├── pdf/                ← 📄 Placez ici les PDFs biométrie (MMT-Full_*.pdf)
│   └── json/               ← 📄 Placez ici Patients.json, Consultation.json, tREFRACTION.json
└── output/
    ├── csv/                ← 📊 biometrie_extraite.csv / .json (généré automatiquement)
    └── graphs/             ← 🖼️  PNGs générés (si SAUVEGARDER=True dans config.py)
```

## Vue d'ensemble

Ce projet traite trois étapes clés du suivi ophtalmologique :

1. **Extraction** : récupération de données biométriques depuis des PDFs
2. **Jointure** : fusion des données biométriques avec un dossier ophtalmologique
3. **Visualisation** : graphes longitudinaux de l'évolution de la myopie et de la longueur axiale

---

## Scripts

### 1. `extraire_biometrie.py`
**Extraction de données biométriques depuis PDFs**

Parcourt `data/pdf/` et extrait :
- Nom du patient, date de naissance, date de mesure
- Longueur axiale composite OD et OG en mm

**Sortie :** `output/csv/biometrie_extraite.csv`

**Dépendances :** `pdfplumber`, `pandas`

---

### 2. `suivi_al.py`
**Visualisation de l'axial length longitudinale**

Trace le suivi temporel de la longueur axiale (AL) pour chaque patient.

**Source :** `output/csv/biometrie_extraite.json`

**Dépendances :** `pandas`, `matplotlib`
3. Graphe interactif montrant l'évolution d'AL en fonction de la date

**Dépendances :** `pandas`, `matplotlib`

---

### 3. `suivi_myopie.py`
**Suivi longitudinal de l'équivalent sphérique (réfraction)**

Visualise l'évolution de la réfraction à partir des JSONs dans `data/json/`.

**Sources :** `data/json/Patients.json`, `Consultation.json`, `tREFRACTION.json`

**Dépendances :** `pandas`, `matplotlib`

---

## Flux de données

```
PDFs biométrie
       ↓
 extraire_biometrie.py
       ↓
biometrie_extraite.csv/json
       ↓ (optionnel)
suivi_al.py → Graphe AL longitudinale
       
Base ophtalmologique (Patients/Consultation/tREFRACTION)
       ↓
 suivi_myopie.py
       ↓
Graphe réfraction longitudinale
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
3. Placer les PDFs dans `data/pdf/` et les JSONs dans `data/json/`
4. Ajuster les paramètres dans **`config.py`** si nécessaire
5. Exécuter les scripts dans l'ordre

---

## Utilisation

### Étape 1 : Extraire les données biométriques
```bash
python extraire_biometrie.py
```
Produit : `biometrie_extraite.csv` (ou `.json` si convertis)

### Étape 2 : Visualiser l'AL longitudinale
```bash
python suivi_al.py
```
Menu interactif → sélectionnez un patient → graphe

### Étape 3 : Visualiser la réfraction longitudinale
```bash
python suivi_myopie.py
```
Affiche les données réfractives par patient ou en cohorte

---

## Formats de données

### `biometrie_extraite.csv`
```
fichier,NOM,DateNaissance,DateMesure,AL_OD,AL_OG
MMT-Full_20260416_004404_1533.pdf,DUPONT Jean,01/01/1990,2026-04-16,23.45,23.52
```

### `Patients.json`
```json
[{"PatientID": 1, "Nom": "DUPONT", "Prenom": "Jean", "DateNaissance": "01/01/1990"}]
```

### `Consultation.json`
```json
[{"ConsultID": 1, "PatientID": 1, "DateConsult": "2026-04-16"}]
```

### `tREFRACTION.json`
```json
[{"RefID": 1, "ConsultID": 1, "Oeil": "D", "Sphere": "-3.50", "Cylindre": "-0.75", "TypeRef": 6}]
```

---

## Cas d'usage

- **Suivi myopique** : visualiser la progression de la myopie chez un patient
- **Analyse AL** : détecter l'allongement axial anormal (facteur clé de la myopie)
- **Recherche cohorte** : comparer l'évolution réfractive entre plusieurs patients
- **Génération de rapports** : exporter les graphes en PNG

---

## Dépannage

| Problème | Solution |
|----------|----------|
| Aucun PDF trouvé | Vérifier `DOSSIER_PDF` et le pattern `MMT-Full_*.pdf` |
| Erreur JSON encoding | Vérifier que les fichiers JSON sont en UTF-8 |
| Graphe vide | Vérifier les dates et que `DateMesure` n'est pas null |
| Dates mal parsées | Adapter `_parse_dob()` au format de vos données |

---

## Licences & Attributions

Project créé pour le suivi ophtalmologique en recherche.

## Contact

Pour toute question ou amélioration, merci de ouvrir une issue.

---

**Dernière mise à jour :** Mai 2026
