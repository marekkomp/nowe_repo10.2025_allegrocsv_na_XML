# --- ZAMIANA read_excel_rows() ---
def read_excel_rows(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    # znajdź wiersz nagłówków skanując pierwsze 20 wierszy
    header_idx = None
    for i, r in enumerate(rows[:20]):
        cells = [str(c).strip().lower() if c is not None else "" for c in r]
        if "id oferty" in " | ".join(cells) or "tytuł oferty" in " | ".join(cells) or "tytuł_oferty" in " | ".join(cells):
            header_idx = i
            break
    if header_idx is None:
        # fallback do starego założenia (nagłówki w 4. wierszu)
        header_idx = 3 if len(rows) > 3 else 0

    headers = [str(h).strip() if h is not None else "" for h in rows[header_idx]]
    data = []
    for r in rows[header_idx + 1:]:
        row_dict = {headers[i]: (str(r[i]).strip() if i < len(r) and r[i] is not None else "") for i in range(len(headers))}
        if any(v for v in row_dict.values()):
            data.append(row_dict)

    print(f"[INFO] XLSM: nagłówki w wierszu {header_idx+1}, rekordów: {len(data)}")
    return data

# --- ZAMIANA read_csv_rows() ---
def read_csv_rows(path):
    # auto-detect separator
    delim = sniff_delimiter(path)
    # wczytaj cały plik, znajdź linię z nagłówkami
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        lines = f.readlines()

    header_line_idx = None
    for i, line in enumerate(lines[:30]):
        low = line.lower()
        if "id oferty" in low or "tytuł oferty" in low or "tytuł_oferty" in low:
            header_line_idx = i
            break
    if header_line_idx is None:
        header_line_idx = 3 if len(lines) > 3 else 0

    from io import StringIO
    buf = StringIO("".join(lines[header_line_idx:]))
    reader = csv.DictReader(buf, delimiter=delim)
    rows = [r for r in reader if any((str(v).strip() for v in r.values()))]
    print(f"[INFO] CSV: nagłówki w linii {header_line_idx+1}, rekordów: {len(rows)}")
    return rows
