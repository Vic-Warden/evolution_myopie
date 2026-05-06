# Évolution Myopie - Suivi Longitudinal de la Longueur Axiale

Ensemble de scripts Python pour l'extraction, l'analyse et la visualisation de données ophtalmologiques, spécifiquement dédiés au suivi de la myopie et de l'allongement axial de l'œil.

## Vue d'ensemble

Ce projet traite trois étapes clés du suivi ophtalmologique :

1. **Extraction** : récupération de données biométriques depuis des PDFs
2. **Jointure** : fusion des données biométriques avec un dossier ophtalmologique
3. **Visualisation** : graphes longitudinaux de l'évolution de la myopie et de la longueur axiale

---

## Scripts

### 1. `extraire_biometrie.py`
**Extraction de données biométriques depuis PDFs**

Parcourt un dossier contenant des PDFs de biométrie (format "MMT-Full_*.pdf") et extrait :
- Nom du patient
- Date de naissance
- Date de mesure (détectée du nom de fichier ou du texte PDF)
- Longueur axiale composite OD (droit) et OG (gauche) en mm

**Configuration :**
- `DOSSIER_PDF` : chemin vers les PDFs (défaut : `c:\Stage\database\pdf`)
- `FICHIER_CSV` : fichier de sortie (défaut : `biometrie_extraite.csv`)

**Sortie :** CSV avec colonnes `[fichier, NOM, DateNaissance, DateMesure, AL_OD, AL_OG]`

**Dépendances :** `pdfplumber`, `pandas`

---

### 2. `suivi_al.py`
**Visualisation de l'axial length longitudinale**

Trace le suivi temporel de la longueur axiale (AL) pour chaque patient. Affiche :
- L'évolution de l'AL OD et OG sur le temps
- Une zone référence pour un globe normal (22–24 mm)
- La pente annuelle d'allongement (indicateur de progression myopique)

**Configuration :**
- `FICHIER_BIOMETRIE` : source des données (défaut : `biometrie_extraite.json`)
- `SAUVEGARDER` : True = PNG sur disque, False = fenêtre interactive

**Flux :**
1. Charge `biometrie_extraite.json`
2. Menu de sélection du patient
3. Graphe interactif montrant l'évolution d'AL en fonction de la date

**Dépendances :** `pandas`, `matplotlib`

---

### 3. `suivi_myopie.py`
**Suivi longitudinal de l'équivalent sphérique (réfraction)**

Visualise l'évolution de la réfraction (équivalent sphérique) à partir d'une base de données ophtalmologique structurée.

**Configuration :**
- `FICHIER_PATIENTS` : données patients (défaut : `Patients.json`)
- `FICHIER_CONSULTATIONS` : historique des consultations (défaut : `Consultation.json`)
- `FICHIER_REFRACTION` : mesures réfractives (défaut : `tREFRACTION.json`)
- `FICHIER_BIOMETRIE` : optionnel, pour enrichissement (défaut : `biometrie_extraite.json`)
- `TYPEREF` : filtrer par type de réfraction (6=Autoréfractomètre, 7=Subjectif, 16=Finale, None=tous)
- `MODE_COHORTE` : True = superpose toutes les courbes, False = graphe par patient
- `OEIL_COHORTE` : "D" ou "G" en mode cohorte
- `SAUVEGARDER` : True = PNG, False = fenêtre interactive

**Flux :**
1. Charge les 3 fichiers JSON (Patients, Consultations, tREFRACTION)
2. Jointure Patient → Consultation → Réfraction
3. Calcul de l'équivalent sphérique (SE = Sphère + Cylindre/2)
4. Visualisation par patient ou en cohorte

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
- Python 3.8+
- pip

### Dépendances
```bash
pip install pdfplumber pandas matplotlib
```

### Setup
1. Cloner le dépôt
2. Configurer les chemins dans chaque script (variables en haut)
3. Placer les données sources (PDFs ou JSONs) aux emplacements configurés
4. Exécuter les scripts dans l'ordre

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
