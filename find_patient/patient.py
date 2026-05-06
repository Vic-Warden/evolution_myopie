"""
find_patient_folder.py
======================
Module utilitaire : résout le dossier photos d'un patient
à partir de son code, en interrogeant PUBLIC.MDB.

Utilisation depuis un autre script :
    from patient import find_patient_folder
    folder = find_patient_folder("1234")
    if folder:
        # faire quelque chose avec folder (Path)

Dépendances : pip install pyodbc
"""

import logging
import sys
from pathlib import Path

# pyodbc est requis uniquement sur Windows
try:
    import pyodbc
    PYODBC_AVAILABLE = True
except ImportError:
    PYODBC_AVAILABLE = False

# Renseigner ces deux chemins avant utilisation
DEST_PHOTOS = Path(r"??")   # Racine du dossier photos patients (ex: r"C:\Stage\photos")
PUBLIC_MDB  = Path(r"??")   # Chemin vers la base Access PUBLIC.MDB (ex: r"C:\Stage\PUBLIC.MDB")

log = logging.getLogger(__name__)


def _db_connect(mdb_path: Path):
    # Ouvre une connexion ODBC vers un fichier Access .mdb/.accdb
    return pyodbc.connect(
        f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_path};"
    )


def find_patient_folder(patient_code: str) -> Path | None:
    """
    Interroge la table Documents de PUBLIC.MDB pour retrouver le dossier
    photos d'un patient à partir de son code.

    Le champ 'Photo externe' contient un chemin relatif (ex: "DUPONT\\2024").
    Ce chemin est reconstruit en chemin absolu via DEST_PHOTOS.

    Retourne un Path vers le dossier si trouvé et existant sur le disque,
    ou None si le patient est inconnu, le chemin absent, ou la base inaccessible.
    """
    if not PYODBC_AVAILABLE:
        log.error("pyodbc n'est pas installé — ce script nécessite Windows avec Access.")
        return None
    if not PUBLIC_MDB.exists():
        log.error(f"PUBLIC.MDB introuvable : {PUBLIC_MDB}")
        return None

    try:
        conn   = _db_connect(PUBLIC_MDB)
        cursor = conn.cursor()

        # On prend le premier enregistrement non-null pour ce patient
        # (un patient peut avoir plusieurs documents, on s'en sert juste pour déduire son dossier)
        cursor.execute(
            "SELECT TOP 1 [Photo externe] FROM Documents "
            "WHERE [code patient] = ? AND [Photo externe] IS NOT NULL",
            (int(patient_code),)
        )
        row = cursor.fetchone()
        conn.close()

        if not row or not row[0]:
            log.warning(f"Aucun document trouvé pour le patient {patient_code}.")
            return None

        # Le champ ressemble à "DUPONT\2024" — on découpe pour reconstruire le chemin complet
        parts = row[0].strip().strip("\\").split("\\")
        if len(parts) < 2:
            log.error(f"Format inattendu pour 'Photo externe' : {row[0]!r}")
            return None

        folder = DEST_PHOTOS / parts[0] / parts[1]

        # Vérifie que le dossier existe réellement sur le disque
        if not folder.is_dir():
            log.error(f"Dossier référencé en base mais absent sur le disque : {folder}")
            return None

        log.info(f"Dossier patient résolu : {folder}")
        return folder

    except Exception as e:
        log.error(f"Erreur lors de la recherche en base : {e}")
        return None


if __name__ == "__main__":
    # Mode test en ligne de commande : python patient.py <code_patient>
    logging.basicConfig(level=logging.INFO, format="%(levelname)-8s %(message)s")

    if len(sys.argv) != 2:
        print("Usage : python patient.py <code_patient>")
        sys.exit(1)

    dossier = find_patient_folder(sys.argv[1])
    if dossier:
        print(f"✓ Dossier trouvé : {dossier}")
    else:
        print("✗ Dossier introuvable.")
        sys.exit(1)