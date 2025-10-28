import os
import openpyxl
import xml.etree.ElementTree as ET

INPUT_DIR = "input"
OUTPUT_DIR = "output"

def convert_file(in_path, out_path):
    wb = openpyxl.load_workbook(in_path, read_only=True, data_only=True)

    # otwieramy tylko arkusz "Szablon"
    if "Szablon" not in wb.sheetnames:
        raise ValueError("Brak arkusza 'Szablon' w pliku")

    ws = wb["Szablon"]
    print(f"[INFO] Przetwarzanie arkusza: {ws.title}")

    # wiersz 4 to nagłówki
    headers = [str(c.value).strip() if c.value else "" for c in ws[4]]
    print("[DEBUG] Nagłówki:", headers[:10])

    root = ET.Element("offers")

    # dane od wiersza 5
    for row in ws.iter_rows(min_row=5, values_only=True):
        data = {headers[i]: (row[i] if i < len(headers) and row[i] else "") for i in range(len(headers))}
        title = str(data.get("Tytuł oferty", "")).strip()
        price = str(data.get("Cena PL", "")).strip()

        if not title:
            continue

        offer = ET.SubElement(root, "offer")
        ET.SubElement(offer, "title").text = title
        ET.SubElement(offer, "price").text = price

    ET.indent(root, space="  ")
    ET.ElementTree(root).write(out_path, encoding="utf-8", xml_declaration=True)
    print(f"[OK] Zapisano: {out_path}")

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    for name in os.listdir(INPUT_DIR):
        if name.lower().endswith((".xlsm", ".xlsx", ".xls")):
            src = os.path.join(INPUT_DIR, name)
            dst = os.path.join(OUTPUT_DIR, os.path.splitext(name)[0] + ".xml")
            convert_file(src, dst)

if __name__ == "__main__":
    main()
