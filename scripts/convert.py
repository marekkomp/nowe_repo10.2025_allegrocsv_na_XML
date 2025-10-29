# scripts/convert.py
import os
import xml.etree.ElementTree as ET
import openpyxl

INPUT_DIR = "input"
OUTPUT_DIR = "output"

REQUIRED = ["Tytuł oferty", "Cena PL"]

def _clean_headers(cells):
    hdr = [("" if v is None else str(v).strip()) for v in cells]
    while hdr and hdr[-1] == "":
        hdr.pop()
    return hdr

def _read_headers_stream(ws, max_col=500):
    row = next(ws.iter_rows(min_row=4, max_row=4, max_col=max_col, values_only=True))
    return _clean_headers(row)

def _read_headers_full(ws):
    max_col = ws.max_column if ws.max_column and ws.max_column > 0 else 500
    cells = [ws.cell(row=4, column=c).value for c in range(1, max_col + 1)]
    return _clean_headers(cells)

def _ensure_required(headers):
    return [h for h in REQUIRED if h not in headers]

def convert_file(in_path, out_path):
    # 1) podejście STREAM (szybkie)
    wb = openpyxl.load_workbook(in_path, read_only=True, data_only=True)
    ws = wb["Szablon"] if "Szablon" in wb.sheetnames else wb.worksheets[0]

    headers = _read_headers_stream(ws, max_col=500)
    print(f"[INFO] Arkusz: {ws.title}")
    print(f"[DEBUG] Nagłówków (stream): {len(headers)}")
    print(f"[DEBUG] Podgląd (stream): {headers[:30]}")
    missing = _ensure_required(headers)

    # 2) jeśli brakuje kolumn — tryb PEŁNY
    if missing:
        print(f"[WARN] Brak w stream: {missing} → przełączam na tryb pełny")
        wb.close()
        wb = openpyxl.load_workbook(in_path, data_only=True)  # bez read_only
        ws = wb["Szablon"] if "Szablon" in wb.sheetnames else wb.worksheets[0]
        headers = _read_headers_full(ws)
        print(f"[DEBUG] Nagłówków (full): {len(headers)}")
        print(f"[DEBUG] Podgląd (full): {headers[:30]}")
        missing = _ensure_required(headers)

    # 3) nadal brakuje → kończymy pustym XML
    if missing:
        print(f"[ERROR] Brak wymaganych nagłówków nawet w trybie pełnym: {missing}")
        ET.ElementTree(ET.Element("offers")).write(out_path, encoding="utf-8", xml_declaration=True)
        return

    idx_title = headers.index("Tytuł oferty")
    idx_price = headers.index("Cena PL")

    # 4) wiersze danych (min_row=5)
    rows = list(ws.iter_rows(min_row=5, values_only=True))

    # jeśli wszystkie wiersze są puste (często gdy są FORMUŁY i data_only=True zwraca None)
    if not any(any(c for c in r if c is not None and str(c).strip() != "") for r in rows):
        print("[WARN] Arkusz wygląda na formułowy (data_only puste). Odczyt z data_only=False.")
        wb.close()
        wb = openpyxl.load_workbook(in_path, data_only=False)  # pokaże tekst formuł/ciągi
        ws = wb["Szablon"] if "Szablon" in wb.sheetnames else wb.worksheets[0]
        rows = list(ws.iter_rows(min_row=5, values_only=True))

    root = ET.Element("offers")

    for row in rows:
        # zabezpieczenie przed krótszym wierszem
        title = ""
        price = ""
        if idx_title < len(row) and row[idx_title] is not None:
            title = str(row[idx_title]).strip()
        if idx_price < len(row) and row[idx_price] is not None:
            price = str(row[idx_price]).strip()

        if not title:
            continue

        offer = ET.SubElement(root, "offer")
        ET.SubElement(offer, "title").text = title
        ET.SubElement(offer, "price").text = price

    ET.indent(root, space="  ")
    ET.ElementTree(root).write(out_path, encoding="utf-8", xml_declaration=True)
    print(f"[OK] Zapisano: {out_path} | ofert: {len(root.findall('offer'))}")

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    any_processed = False
    for name in os.listdir(INPUT_DIR):
        if name.lower().endswith((".xlsm", ".xlsx", ".xls")):
            src = os.path.join(INPUT_DIR, name)
            dst = os.path.join(OUTPUT_DIR, os.path.splitext(name)[0] + ".xml")
            print(f"[RUN] {src} -> {dst}")
            convert_file(src, dst)
            any_processed = True
    if not any_processed:
        print("[INFO] Brak plików wejściowych w /input")

if __name__ == "__main__":
    main()
