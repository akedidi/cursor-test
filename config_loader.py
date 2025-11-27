import os
import logging
from dotenv import load_dotenv


def setup_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )


def load_env():
    """
    Charge les variables d'environnement et normalise les chemins.
    """
    load_dotenv()
    setup_logging()

    results_folder = os.getenv("RESULTS_FOLDER")
    output_file = os.getenv("OUTPUT_FILE", "recap_scenarios.xlsx")
    doc_template = os.getenv("DOC_TEMPLATE")
    doc_output = os.getenv("DOC_OUTPUT")

    logging.info("RESULTS_FOLDER = %s", results_folder)
    logging.info("OUTPUT_FILE   = %s", output_file)
    logging.info("DOC_TEMPLATE  = %s", doc_template)
    logging.info("DOC_OUTPUT    = %s", doc_output)

    if not results_folder:
        raise ValueError("La variable RESULTS_FOLDER n'est pas définie dans le fichier .env")
    if not os.path.isdir(results_folder):
        raise ValueError(f"Le dossier RESULTS_FOLDER n'existe pas : {results_folder}")

    if os.path.isdir(output_file) or not os.path.splitext(output_file)[1]:
        output_file = os.path.join(output_file, "recap_scenarios.xlsx")
        logging.info("OUTPUT_FILE normalisé en : %s", output_file)

    return results_folder, output_file, doc_template, doc_output
