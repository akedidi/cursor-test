def compute_sale_per_minute(rows):
    """
    Calcule le nombre de sales Coppel par minute.

    Hypothèse :
    - Sale = API Purchase
    - Exclure les labels finissant par _BNS
    - Durée = période complète du scénario basée sur timeStamp min/max
    """

    timestamps = []
    sale_count = 0

    for r in rows:
        label = str(r.get("label", "")).strip()

        if label.endswith("_BNS"):
            continue

        ts_raw = r.get("timeStamp")
        if ts_raw is not None:
            try:
                timestamps.append(int(ts_raw))
            except Exception:
                pass

        if label == "Purchase":
            sale_count += 1

    if not timestamps:
        return "0.00"

    duration_ms = max(timestamps) - min(timestamps)
    duration_min = duration_ms / 1000.0 / 60.0

    if duration_min <= 0:
        return "0.00"

    sale_per_min = sale_count / duration_min

    return f"{sale_per_min:.2f}"
    
    
    
    
    
    
    # 1.b) Sales per minute
for idx, users in enumerate(sorted(scenarios_users), start=1):

    rows = scenario_rows.get(users, [])
    if not rows:
        continue

    sale_min = compute_sale_per_minute(rows)

    placeholder = f"{{SALE_MIN_{idx}}}"

    found = False

    for t in root.findall(".//w:t", NS):

        if t.text == placeholder:
            t.text = sale_min
            found = True

    if found:
        logging.info(
            "Remplacement de %s par '%s'",
            placeholder,
            sale_min
        )
    else:
        logging.info(
            "Placeholder %s non trouvé dans le document.",
            placeholder
        )
    