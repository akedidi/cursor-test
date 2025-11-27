def build_response_time_table_xml(recap):
    # 🔥 Début du tableau Word
    table_xml = """
    <w:tbl>
      <w:tblPr>
        <w:tblBorders>
          <w:top    w:val="single" w:sz="8" w:space="0" w:color="000000"/>
          <w:left   w:val="single" w:sz="8" w:space="0" w:color="000000"/>
          <w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>
          <w:right  w:val="single" w:sz="8" w:space="0" w:color="000000"/>
          <w:insideH w:val="single" w:sz="6" w:space="0" w:color="000000"/>
          <w:insideV w:val="single" w:sz="6" w:space="0" w:color="000000"/>
        </w:tblBorders>
      </w:tblPr>
    """

    # ================================ 🔥 HEADER en gras ================================
    headers = ["Label","# Samples","Average","Min","Max","Std. Dev.","Error %","Throughput",
               "Received KB/sec","Sent KB/sec","Avg. Bytes"]

    table_xml += "<w:tr>"
    for h in headers:
        table_xml += f"""
        <w:tc><w:p><w:r><w:rPr><w:b/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
        <w:t>{xml_escape(h)}</w:t></w:r></w:p></w:tc>
        """
    table_xml += "</w:tr>"

    # ================================ 🔥 DATA ROWS ================================
    for row in recap:
        cells = [
            row["Label"],                              # BOLD COLUMN 1
            row["Samples"],
            int(row["Average (ms)"]),
            int(row["Min (ms)"]),
            int(row["Max (ms)"]),
            row["Std Dev (ms)"],
            f"{row['Error %']:.2f}%",
            row["Throughput (/min)"],
            row["Received KB/sec"],
            row["Sent KB/sec"],
            row["Avg Bytes"],
        ]

        table_xml += "<w:tr>"

        for i, val in enumerate(cells):
            text = xml_escape(str(val))  # échappement XML obligatoire

            # 🔥 Bold uniquement première colonne
            if i == 0:
                table_xml += f"""
                <w:tc><w:p><w:r><w:rPr><w:b/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>
                <w:t>{text}</w:t></w:r></w:p></w:tc>
                """
            else:
                table_xml += f"""
                <w:tc><w:p><w:r><w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>
                <w:t>{text}</w:t></w:r></w:p></w:tc>
                """

        table_xml += "</w:tr>"

    # ================================ 🔥 END TABLE ================================
    table_xml += "</w:tbl>"

    return table_xml
