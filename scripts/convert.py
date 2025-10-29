# scripts/convert.py
import os
import xml.etree.ElementTree as ET
import openpyxl

INPUT_DIR = "input"
OUTPUT_DIR = "output"

NEEDED = ["Tytuł oferty", "Cena PL"]

def convert_file(in_path: str, out_path: str) -> None:
    # tryb read_only + iter_rows -> poprawny odczyt całego wiersza 4
    wb = openpyxl.load_workbook(in_path, read_only=True, data_only=True)

    ws = wb["Szablon"] if "Szablon" in wb.sheetnames else wb.worksheets[0]
    print(f"[INFO] Arkusz: {ws.title}")

    # nagłówki z wiersza 4
    header_row = next(ws.iter_rows(min_row=4, max_row=4, values_only=True))
    headers = [str(c).strip() if c is not None else "" for c in header_row]
    print(f"[DEBUG] Liczba kolumn: {len(headers)}")
    print(f"[DEBUG] Pierwsze 15 nagłówków: {headers[:15]}")

    # walidacja
    missing = [h for h in NEEDED if h not in headers]
    if missing:
        print(f"[ERROR] Brak wymaganych nagłówków: {missing}")
        # nadal zapisz pusty root, żeby workflow się nie wywalał
        ET.ElementTree(ET.Element("offers")).write(out_path, encoding="utf-8", xml_declaration=True)
        return

    idx_title = headers.index("Tytuł oferty")
    idx_price = headers.index("Cena PL")

    root = ET.Element("offers")

    # dane od wiersza 5
    for row in ws.iter_rows(min_row=5, values_only=True):
        # zabezpieczenie przed krótszym wierszem
        title = str(row[idx_title]).strip() if idx_title < len(row) and row[idx_title] is not None else ""
        price = str(row[idx_price]).strip() if idx_price < len(row) and row[idx_price] is not None else ""

        if not title:
            continue

        offer = ET.SubElement(root, "offer")
        ET.SubElement(offer, "title").text = title
        ET.SubElement(offer, "price").text = price

    ET.indent(root, space="  ")
    tree = ET.ElementTree(root)
    tree.write(out_path, encoding="utf-8", xml_declaration=True)
    print(f"[OK] Zapisano: {out_path} (ofert: {len(root.findall('offer'))})")

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    any_processed = False
    for name in os.listdir(INPUT_DIR):
        if not name.lower().endswith((".xlsm", ".xlsx", ".xls")):
            continue
        src = os.path.join(INPUT_DIR, name)
        dst = os.path.join(OUTPUT_DIR, os.path.splitext(name)[0] + ".xml")
        print(f"[RUN] {src} -> {dst}")
        convert_file(src, dst)
        any_processed = True
    if not any_processed:
        print("[INFO] Brak plików wejściowych w /input")

if __name__ == "__main__":
    main()
