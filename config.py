"""
config.py — Configuration centralisée du projet evolution_myopie
================================================================
Modifiez ce fichier pour adapter les chemins et paramètres à votre environnement.
Tous les scripts importent leurs constantes depuis ici.
"""

from pathlib import Path

# Chemins racines

BASE_DIR = Path(__file__).parent

# Données en entrée
DATA_DIR      = BASE_DIR / "data"
DOSSIER_PDF   = DATA_DIR / "pdf"          # PDFs biométrie  (MMT-Full_*.pdf)
DOSSIER_JSON  = DATA_DIR / "json"         # JSONs ophtalmo   (Patients / Consultation / tREFRACTION)

# Sorties
OUTPUT_DIR    = BASE_DIR / "output"
OUTPUT_CSV    = OUTPUT_DIR / "csv"        # biometrie_extraite.csv / .json
OUTPUT_GRAPHS = OUTPUT_DIR / "graphs"     # PNGs générés

# Fichiers de données (extraire_biometrie.py → output, suivi_*.py → input)

FICHIER_CSV_BIOMETRIE  = OUTPUT_CSV  / "biometrie_extraite.csv"
FICHIER_JSON_BIOMETRIE = OUTPUT_CSV  / "biometrie_extraite.json"

FICHIER_PATIENTS      = DOSSIER_JSON / "Patients.json"
FICHIER_CONSULTATIONS = DOSSIER_JSON / "Consultation.json"
FICHIER_REFRACTION    = DOSSIER_JSON / "tREFRACTION.json"

# Paramètres — suivi_myopie.py

# IDs patients à tracer (None = menu interactif, liste = ex. ["1234", "5678"])
PATIENT_IDS  = None

# TypeRef à filtrer (6 = Autoréfractomètre, 7 = Subjectif, 16 = Finale, None = tous)
TYPEREF      = 6

# Mode cohorte : True = toutes les courbes superposées, False = individuel
MODE_COHORTE = False
OEIL_COHORTE = "D"         # "D" ou "G" (utilisé uniquement en mode cohorte)

# Paramètres — communs

# True = PNG enregistré dans output/graphs/, False = fenêtre interactive
SAUVEGARDER = False
