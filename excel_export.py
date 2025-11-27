import re
import logging
import xlsxwriter
from metrics import LABEL_ORDER


def sanitize_sheet_name(name: str) -> str:
    name = re.sub(r'[:\\/?*\[\]]', "_", name)
    if len(name) > 31:
        name = name[:31]
    if not name:
        name = "Sheet"
    return name


def write_excel(output_file: str,
                scenarios_data: dict,
                scenarios_users: list,
                rt_matrix: dict,
                err_matrix: dict):
    logging.info("Création du fichier Excel : %s", output_file)
    workbook = xlsxwriter.Workbook(output_file)

    header_fmt = workbook.add_format({
        "bold": True,
        "bg_color": "#D9D9D9",
        "border": 1
    })
    cell_fmt = workbook.add_format({"border": 1})
    num_fmt = workbook.add_format({"border": 1, "num_format": "0.00"})
    int_fmt = workbook.add_format({"border": 1, "num_format": "0"})

    # colonnes type JMeter pour chaque scénario
    headers = [
        "Label",
        "# Samples",
        "Average",
        "Min",
        "Max",
        "Std. Dev.",
        "Error %",
        "Throughput",
        "Received KB/sec",
        "Sent KB/sec",
        "Avg. Bytes",
    ]

    col_key_map = {
        "Label": "Label",
        "# Samples": "Samples",
        "Average": "Average (ms)",
        "Min": "Min (ms)",
        "Max": "Max (ms)",
        "Std. Dev.": "Std Dev (ms)",
        "Error %": "Error %",
        "Throughput": "Throughput (/min)",
        "Received KB/sec": "Received KB/sec",
        "Sent KB/sec": "Sent KB/sec",
        "Avg. Bytes": "Avg Bytes",
    }

    # Feuilles par scénario
    for sheet_name_raw, rows in scenarios_data.items():
        sheet_name = sanitize_sheet_name(sheet_name_raw)
        logging.info("  -> Création de la feuille : %s", sheet_name)

        ws = workbook.add_worksheet(sheet_name)

        if not rows:
            logging.warning("    (Aucune donnée pour ce scénario)")
            continue

        for col, h in enumerate(headers):
            ws.write(0, col, h, header_fmt)

        for row_idx, row in enumerate(rows, start=1):
            for col_idx, h in enumerate(headers):
                key = col_key_map[h]
                val = row.get(key, "")
                if isinstance(val, (int, float)):
                    if h in ["Average", "Min", "Max"]:
                        ws.write(row_idx, col_idx, int(round(val)), int_fmt)
                    else:
                        ws.write(row_idx, col_idx, val, num_fmt)
                else:
                    ws.write(row_idx, col_idx, str(val), cell_fmt)

        ws.set_column(0, 0, 40)
        ws.set_column(1, len(headers) - 1, 16)

    # Onglet Data Time Response Time
    ws_rt = workbook.add_worksheet("Data Time Response Time")
    ws_rt.write(0, 0, "Scenario", header_fmt)
    ws_rt.write(0, 1, "API", header_fmt)
    ws_rt.write(0, 2, "Response Time (ms)", header_fmt)

    row_idx = 1
    for users in sorted(scenarios_users):
        start_row = row_idx
        for label in LABEL_ORDER:
            val = rt_matrix.get(label, {}).get(users, None)
            if val is None:
                continue
            ws_rt.write(row_idx, 1, label, cell_fmt)
            ws_rt.write(row_idx, 2, int(round(val)), int_fmt)
            row_idx += 1
        end_row = row_idx - 1
        if end_row >= start_row:
            ws_rt.merge_range(start_row, 0, end_row, 0, users, int_fmt)

    ws_rt.set_column(0, 0, 12)
    ws_rt.set_column(1, 1, 20)
    ws_rt.set_column(2, 2, 20)

    # Onglet Data Error Rate
    ws_err = workbook.add_worksheet("Data Error Rate")
    ws_err.write(0, 0, "Scenario", header_fmt)
    ws_err.write(0, 1, "API", header_fmt)
    ws_err.write(0, 2, "Error Rate (%)", header_fmt)

    row_idx = 1
    for users in sorted(scenarios_users):
        start_row = row_idx
        for label in LABEL_ORDER:
            val = err_matrix.get(label, {}).get(users, None)
            if val is None:
                continue
            ws_err.write(row_idx, 1, label, cell_fmt)

            if abs(val - round(val)) < 1e-9:
                s = str(int(round(val)))
            else:
                s = f"{val:.2f}"
            ws_err.write(row_idx, 2, s, cell_fmt)

            row_idx += 1
        end_row = row_idx - 1
        if end_row >= start_row:
            ws_err.merge_range(start_row, 0, end_row, 0, users, int_fmt)

    ws_err.set_column(0, 0, 12)
    ws_err.set_column(1, 1, 20)
    ws_err.set_column(2, 2, 20)

    workbook.close()
    logging.info("Fichier Excel finalisé.")
