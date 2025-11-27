import os
import re
import glob
import csv
import logging


def extract_users_from_filename(path: str) -> int:
    """
    Extrait le nombre d'utilisateurs à partir du nom de fichier.
    Ex : ...results-1-users.csv -> 1
         ...results-12-user.csv -> 12
    """
    name = os.path.basename(path)
    m = re.search(r"results-(\d+)-user", name)
    if m:
        return int(m.group(1))
    return 999999


def find_scenario_files(results_folder: str):
    """
    On accepte :
      IDP API-results-1-user.csv
      IDP API-results-1-users.csv
    """
    pattern = os.path.join(results_folder, "IDP API-results-*user*.csv")
    logging.info("Recherche des fichiers avec le pattern : %s", pattern)
    files = glob.glob(pattern)

    if not files:
        raise FileNotFoundError(f"Aucun fichier trouvé avec le pattern : {pattern}")

    files = sorted(files, key=extract_users_from_filename)

    logging.info("Nombre de fichiers trouvés : %d", len(files))
    for f in files:
        logging.info(" - %s", f)

    return files


def read_jmeter_csv(path: str):
    logging.info("Lecture du fichier CSV : %s", path)
    rows = []
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for r in reader:
            rows.append(r)
    logging.info("  -> %d lignes lues (hors en-tête)", len(rows))
    return rows
