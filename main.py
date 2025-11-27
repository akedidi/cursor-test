import logging
from collections import defaultdict

from config_loader import load_env
from jmeter_io import find_scenario_files, read_jmeter_csv, extract_users_from_filename
from metrics import compute_recap
from excel_export import write_excel
from word_export import generate_word_report


def main():
    try:
        results_folder, output_file, doc_template, doc_output = load_env()
        files = find_scenario_files(results_folder)

        scenarios_data = {}
        scenarios_users = []
        rt_matrix = defaultdict(dict)   # label -> {users: avg}
        err_matrix = defaultdict(dict)  # label -> {users: error%}
        scenario_rows = {}              # users -> rows CSV
        scenario_recaps_by_users = {}   # users -> recap

        for f in files:
            logging.info("--------------------------------------------------")
            logging.info("Traitement du fichier scénario : %s", f)

            users = extract_users_from_filename(f)
            if users not in scenarios_users:
                scenarios_users.append(users)

            rows = read_jmeter_csv(f)
            recap = compute_recap(rows)

            base_name = os.path.splitext(os.path.basename(f))[0]
            scenarios_data[base_name] = recap
            scenario_rows[users] = rows
            scenario_recaps_by_users[users] = recap

            for r in recap:
                if r["Label"] == "TOTAL":
                    continue
                label = r["Label"]
                rt_matrix[label][users] = r["Average (ms)"]
                err_matrix[label][users] = r["Error %"]

        write_excel(output_file, scenarios_data, scenarios_users, rt_matrix, err_matrix)

        if doc_template and doc_output:
            generate_word_report(doc_template, doc_output,
                                 scenarios_users, scenario_recaps_by_users, scenario_rows)
        else:
            logging.info("DOC_TEMPLATE ou DOC_OUTPUT non défini, Word ignoré.")

        logging.info("Terminé ✅")

    except Exception as e:
        logging.exception("❌ Erreur lors de l'exécution du script : %s", e)


if __name__ == "__main__":
    main()
