# scripts/convert.py
import os
import xml.etree.ElementTree as ET
import openpyxl

INPUT_DIR = "input"
OUTPUT_DIR = "output"

REQUIRED = ["Tytuł oferty", "Cena PL"]

def _clean_headers(cells):
    # zamień None na "", przytnij spacje
    hdr = [("" if v is None else str(v).strip()) for v in cells]
    # utnij ogon pustych kolumn
    while hdr and hdr[-1] == "":
        hdr.pop()
    return hdr

def _read_headers_stream(ws, max_col=500):
    row = next(ws.iter_rows(min_row=4, max_row=4, max_col=max_col, values_only=True))
    return _clean_headers(row)

def _read_headers_full(ws):
    # pełny tryb (bez read_only): odczyt całego wiersza 4 po max_col
    max_col = ws.max_column if ws.max_column and ws.max_column > 0 else 500
    cells = [ws.cell(row=4, column=c).value for c in range(1, max_col + 1)]
    return _clean_headers(cells)

def _ensure_required(headers):
    missing = [h for h in REQUIRED if h not in headers]
    return missing

def convert_file(in_path, out_path):
    # 1) podejście STREAM (szybkie)
    wb = openpyxl.load_workbook(in_path, read_only=True, data_only=True)
    ws = wb["Szablon"] if "Szablon" in wb.sheetnames else wb.worksheets[0]

    headers = _read_headers_stream(ws, max_col=500)
    print(f"[INFO] Arkusz: {ws.title}")
    print(f"[DEBUG] Nagłówków (stream): {len(headers)}")
    print(f"[DEBUG] Podgląd (stream): {headers[:30]}")

    missing = _ensure_required(headers)

    # 2) jeśli brakuje kolumn — przełącz na tryb PEŁNY i czytaj wiersz 4 jeszcze raz
    if missing:
        print(f"[WARN] Brak w stream: {missing} → przełączam na tryb pełny")
        wb.close()
        wb = openpyxl.load_workbook(in_path, data_only=True)  # bez read_only
        ws = wb["Szablon"] if "Szablon" in wb.sheetnames else wb.worksheets[0]
        headers = _read_headers_full(ws)
        print(f"[DEBUG] Nagłówków (full): {len(headers)}")
        print(f"[DEBUG] Podgląd (full): {headers[:30]}")
        missing = _ensure_required(headers)

    # 3) jeśli nadal brakuje — zapisz pusty XML i jasno zgłoś błąd w logu
    if missing:
        print(f"[ERROR] Brak wymaganych nagłówków nawet w trybie pełnym: {missing}")
        ET.ElementTree(ET.Element("offers")).write(out_path, encoding="utf-8", xml_declaration=True)
        return

    idx_title = headers.index("Tytuł oferty")
    idx_price = headers.index("Cena PL")

    root = ET.Element("offers")

    # dane od wiersza 5 (oba tryby działają tak samo)
    for row in ws.iter_rows(min_row=5, values_only=True):
        # row może być krótszy — zabezpieczenie
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
