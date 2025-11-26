import os
import glob
import csv
import statistics
import re
import logging

from dotenv import load_dotenv
import xlsxwriter


# --------------------------------------------------------
# Configuration du logging
# --------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)


# --------------------------------------------------------
# Chargement des variables d'environnement
# --------------------------------------------------------
def load_env():
    """
    Charge .env et retourne (results_folder, output_file)
    """
    logging.info("Chargement du fichier .env...")
    load_dotenv()

    results_folder = os.getenv("RESULTS_FOLDER")
    output_file = os.getenv("OUTPUT_FILE", "recap_scenarios.xlsx")

    logging.info("RESULTS_FOLDER = %s", results_folder)
    logging.info("OUTPUT_FILE   = %s", output_file)

    if not results_folder:
        logging.error("La variable RESULTS_FOLDER n'est pas définie dans le fichier .env")
        raise ValueError("La variable RESULTS_FOLDER n'est pas définie dans le fichier .env")

    if not os.path.isdir(results_folder):
        logging.error("Le dossier RESULTS_FOLDER n'existe pas : %s", results_folder)
        raise ValueError(f"Le dossier RESULTS_FOLDER n'existe pas : {results_folder}")

    return results_folder, output_file


# --------------------------------------------------------
# Recherche des fichiers CSV JMeter
# --------------------------------------------------------
def find_scenario_files(results_folder: str):
    """
    Cherche les fichiers du type :
    IDP API-results-n-users.csv
    (n varie selon le scénario)
    """
    pattern = os.path.join(results_folder, "IDP API-results-*-users.csv")
    logging.info("Recherche des fichiers avec le pattern : %s", pattern)
    files = glob.glob(pattern)

    if not files:
        logging.warning("Aucun fichier trouvé avec le pattern : %s", pattern)
        raise FileNotFoundError(f"Aucun fichier trouvé avec le pattern : {pattern}")

    files = sorted(files)
    logging.info("Nombre de fichiers trouvés : %d", len(files))
    for f in files:
        logging.info(" - %s", f)

    return files


# --------------------------------------------------------
# Lecture d'un CSV JMeter
# --------------------------------------------------------
def read_jmeter_csv(path: str):
    """
    Lit un fichier CSV JMeter et renvoie une liste de dictionnaires.
    Hypothèse : le fichier a une ligne d'en-tête avec au moins :
    - label
    - elapsed
    - success
    """
    logging.info("Lecture du fichier CSV : %s", path)
    rows = []
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for r in reader:
            rows.append(r)

    logging.info("  -> %d lignes lues (hors en-tête)", len(rows))
    return rows


# --------------------------------------------------------
# Helpers pour convertir les champs
# --------------------------------------------------------
def to_float(value, default=None):
    try:
        return float(value)
    except Exception:
        return default


def to_bool_success(value):
    """
    Convertit la colonne 'success' en booléen.
    Gère true/false, 1/0, YES/NO...
    """
    if value is None:
        return False
    v = str(value).strip().lower()
    return v in ("true", "1", "yes", "y")


# --------------------------------------------------------
# Calcul du tableau récapitulatif pour un scénario
# --------------------------------------------------------
def compute_recap(rows):
    """
    rows : liste de dicts (une par ligne du CSV)
    Retourne une liste de dicts pour le recap.
    """
    logging.info("Calcul du tableau récapitulatif pour le scénario...")
    labels = {}

    for r in rows:
        label = r.get("label")
        elapsed_raw = r.get("elapsed")
        success_raw = r.get("success")

        if label is None or elapsed_raw is None:
            continue

        elapsed = to_float(elapsed_raw)
        if elapsed is None:
            continue

        success = to_bool_success(success_raw)

        if label not in labels:
            labels[label] = {
                "times": [],
                "errors": 0,
            }

        labels[label]["times"].append(elapsed)
        if not success:
            labels[label]["errors"] += 1

    recap = []

    total_samples_all = 0
    all_times = []
    total_errors_all = 0

    for label, data in sorted(labels.items(), key=lambda x: x[0]):
        times = data["times"]
        errors = data["errors"]
        samples = len(times)

        if samples == 0:
            continue

        avg = statistics.mean(times)
        mn = min(times)
        mx = max(times)
        err_pct = (errors / samples * 100.0) if samples else 0.0

        recap.append({
            "Label": label,
            "Samples": samples,
            "Average (ms)": round(avg, 2),
            "Min (ms)": mn,
            "Max (ms)": mx,
            "Error %": round(err_pct, 2),
        })

        total_samples_all += samples
        total_errors_all += errors
        all_times.extend(times)

    # Ligne TOTAL
    if all_times:
        total_avg = statistics.mean(all_times)
        total_min = min(all_times)
        total_max = max(all_times)
        total_err_pct = (total_errors_all / total_samples_all * 100.0) if total_samples_all else 0.0

        recap.append({
            "Label": "TOTAL",
            "Samples": total_samples_all,
            "Average (ms)": round(total_avg, 2),
            "Min (ms)": total_min,
            "Max (ms)": total_max,
            "Error %": round(total_err_pct, 2),
        })

    logging.info("  -> %d lignes dans le tableau récap (y compris TOTAL)", len(recap))
    return recap


# --------------------------------------------------------
# Ecriture du fichier Excel avec xlsxwriter
# --------------------------------------------------------
def sanitize_sheet_name(name: str) -> str:
    """
    Nettoie le nom de feuille Excel :
    - max 31 caractères
    - supprime les caractères interdits : : \ / ? * [ ]
    """
    name = re.sub(r'[:\\/?*\[\]]', "_", name)
    if len(name) > 31:
        name = name[:31]
    if not name:
        name = "Sheet"
    return name


def write_excel(output_file: str, scenarios_data: dict):
    """
    scenarios_data : dict { sheet_name: [ {col: val}, ... ] }
    Crée un fichier Excel avec une feuille par scénario.
    """
    logging.info("Création du fichier Excel : %s", output_file)
    workbook = xlsxwriter.Workbook(output_file)

    # Formats simples
    header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9D9D9", "border": 1})
    cell_fmt = workbook.add_format({"border": 1})
    num_fmt = workbook.add_format({"border": 1, "num_format": "0.00"})

    for sheet_name_raw, rows in scenarios_data.items():
        sheet_name = sanitize_sheet_name(sheet_name_raw)
        logging.info("  -> Création de la feuille : %s", sheet_name)

        worksheet = workbook.add_worksheet(sheet_name)

        if not rows:
            logging.warning("    (Aucune donnée pour ce scénario)")
            continue

        headers = ["Label", "Samples", "Average (ms)", "Min (ms)", "Max (ms)", "Error %"]

        # Ecrire les en-têtes
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_fmt)

        # Ecrire les données
        for row_idx, row_dict in enumerate(rows, start=1):
            for col_idx, header in enumerate(headers):
                value = row_dict.get(header, "")
                if isinstance(value, (int, float)):
                    worksheet.write(row_idx, col_idx, value, num_fmt)
                else:
                    worksheet.write(row_idx, col_idx, str(value), cell_fmt)

        # Ajuster la largeur des colonnes
        worksheet.set_column(0, 0, 40)  # Label
        worksheet.set_column(1, 5, 15)  # chiffres

    workbook.close()
    logging.info("Fichier Excel finalisé.")


# --------------------------------------------------------
# MAIN
# --------------------------------------------------------
def main():
    try:
        results_folder, output_file = load_env()
        files = find_scenario_files(results_folder)

        scenarios_data = {}

        for f in files:
            logging.info("--------------------------------------------------")
            logging.info("Traitement du fichier scénario : %s", f)

            csv_rows = read_jmeter_csv(f)
            recap_rows = compute_recap(csv_rows)

            base_name = os.path.splitext(os.path.basename(f))[0]
            scenarios_data[base_name] = recap_rows

        write_excel(output_file, scenarios_data)
        logging.info("Terminé ✅. Fichier Excel généré : %s", output_file)

    except Exception as e:
        logging.exception("❌ Erreur lors de l'exécution du script : %s", e)


if __name__ == "__main__":
    main()
