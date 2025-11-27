for row in recap:
    cells = [
        row["Label"],
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
        text = xml_escape(str(val))  # ESCAPE OBLIGATOIRE

        if i == 0:   # 🔥 première colonne bold
            table_xml += f"""
            <w:tc><w:p><w:r><w:rPr><w:b/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>
            <w:t>{text}</w:t></w:r></w:p></w:tc>
            """
        else:        # colonnes normales
            table_xml += f"""
            <w:tc><w:p><w:r><w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>
            <w:t>{text}</w:t></w:r></w:p></w:tc>
            """

    table_xml += "</w:tr>"
