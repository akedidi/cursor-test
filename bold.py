for i, val in enumerate(cells):
    text = xml_escape(val)

    # 🔥 bold uniquement si première colonne (i == 0)
    if i == 0:
        data_rows_xml += f"""
        <w:tc>
          <w:tcPr/>
          <w:p>
            <w:r>
              <w:rPr>
                <w:b/>                      <!-- Gras -->
                <w:sz w:val="16"/>
                <w:szCs w:val="16"/>
              </w:rPr>
              <w:t>{text}</w:t>
            </w:r>
          </w:p>
        </w:tc>
        """
    else:
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
