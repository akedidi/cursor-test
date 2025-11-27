import logging
from zipfile import ZipFile
import xml.etree.ElementTree as ET

from metrics import compute_execution_range_string
# mêmes namespaces que dans ton script monolithique
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
ET.register_namespace("w", W_NS)


def xml_escape(text: str) -> str:
    if text is None:
        return ""
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def build_response_time_table_xml(recap):
    """
    Table JMeter-like :
      Label, # Samples, Average, Min, Max, Std. Dev., Error %, Throughput,
      Received KB/sec, Sent KB/sec, Avg. Bytes
    """
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

    header_row_xml = "<w:tr>"
    for h in headers:
        header_row_xml += f"""
        <w:tc>
          <w:tcPr/>
          <w:p>
            <w:r>
              <w:rPr>
                <w:b/>
                <w:sz w:val="16"/>
                <w:szCs w:val="16"/>
              </w:rPr>
              <w:t>{xml_escape(h)}</w:t>
            </w:r>
          </w:p>
        </w:tc>
        """
    header_row_xml += "</w:tr>"

    data_rows_xml = ""
    for r in recap:
        data_rows_xml += "<w:tr>"

        cells = [
            r["Label"],
            r["Samples"],
            int(r["Average (ms)"]),
            int(r["Min (ms)"]),
            int(r["Max (ms)"]),
            r["Std Dev (ms)"],
            f"{r['Error %']:.2f}%",
            r["Throughput (/min)"],
            r["Received KB/sec"],
            r["Sent KB/sec"],
            r["Avg Bytes"],
        ]

        for val in cells:
            text = xml_escape(val)
            data_rows_xml += f"""
            <w:tc>
              <w:tcPr/>
              <w:p>
                <w:r>
                  <w:rPr>
                    <w:sz w:val="16"/>
                    <w:szCs w:val="16"/>
                  </w:rPr>
                  <w:t>{text}</w:t>
                </w:r>
              </w:p>
            </w:tc>
            """

        data_rows_xml += "</w:tr>"

    table_xml = f"""
    <w:tbl xmlns:w="{W_NS}">
      <w:tblPr>
        <w:tblBorders>
          <w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>
          <w:left w:val="single" w:sz="8" w:space="0" w:color="000000"/>
          <w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>
          <w:right w:val="single" w:sz="8" w:space="0" w:color="000000"/>
          <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
          <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        </w:tblBorders>
      </w:tblPr>
      {header_row_xml}
      {data_rows_xml}
    </w:tbl>
    """
    return table_xml.strip()


def generate_word_report(template_path, output_path,
                         scenarios_users, scenario_recaps, scenario_rows):
    """
    Modifie le template Word (DOCX comme ZIP) :
      - remplit les dates d'exécution : {EXEC_DATE_1}, {EXEC_DATE_2}, ...
      - remplace le paragraphe contenant {RT_TABLE_n} par un <w:tbl> construit.
    """
    if not template_path:
        logging.warning("DOC_TEMPLATE non défini, génération Word ignorée.")
        return

    if not os.path.isfile(template_path):
        logging.error("DOC_TEMPLATE n'existe pas : %s", template_path)
        return

    logging.info("Ouverture du template Word (ZIP) : %s", template_path)

    with ZipFile(template_path, "r") as z:
        content = {name: z.read(name) for name in z.namelist()}

    if "word/document.xml" not in content:
        logging.error("word/document.xml introuvable dans le template.")
        return

    xml_bytes = content["word/document.xml"]
    root = ET.fromstring(xml_bytes)

    parent_map = {child: parent for parent in root.iter() for child in parent}

    # 1) Dates d'exécution
    exec_strings = []
    for users in sorted(scenarios_users):
        rows = scenario_rows.get(users, [])
        exec_strings.append(compute_execution_range_string(rows))

    for i, date_str in enumerate(exec_strings, start=1):
        placeholder = f"{{EXEC_DATE_{i}}}"
        found = False
        for t in root.findall(".//w:t", NS):
            if t.text == placeholder:
                t.text = date_str
                found = True
        if found:
            logging.info("Remplacement de %s par '%s'", placeholder, date_str)
        else:
            logging.info("Placeholder %s non trouvé dans le document.", placeholder)

    # 2) Tableaux Response time
    for idx, users in enumerate(sorted(scenarios_users), start=1):
        recap = scenario_recaps.get(users)
        if not recap:
            continue

        placeholder = f"{{RT_TABLE_{idx}}}"
        table_xml = build_response_time_table_xml(recap)
        table_el = ET.fromstring(table_xml)

        replaced = False
        for p in root.findall(".//w:p", NS):
            has_placeholder = False
            for t in p.findall(".//w:t", NS):
                if t.text == placeholder:
                    has_placeholder = True
                    break
            if not has_placeholder:
                continue

            parent = parent_map.get(p)
            if parent is None:
                continue

            idx_in_parent = list(parent).index(p)
            parent.remove(p)
            parent.insert(idx_in_parent, table_el)
            replaced = True
            logging.info("Tableau Response time inséré à la place de %s (users=%d)", placeholder, users)
            break

        if not replaced:
            logging.info("Placeholder %s non trouvé pour users=%d.", placeholder, users)

    new_xml_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    with ZipFile(output_path, "w") as z:
        for name, data in content.items():
            if name == "word/document.xml":
                z.writestr(name, new_xml_bytes)
            else:
                z.writestr(name, data)

    logging.info("Document Word généré : %s", output_path)
