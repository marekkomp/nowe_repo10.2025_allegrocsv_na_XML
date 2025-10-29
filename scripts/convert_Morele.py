# scripts/convert_Morele.py
import os
import xml.etree.ElementTree as ET
from convert import convert_file, INPUT_DIR, OUTPUT_DIR

def convert_file_morele(in_path, out_path):
    # --- najpierw wygeneruj standardowy XML ---
    temp_path = os.path.join(OUTPUT_DIR, "_temp_base.xml")
    convert_file(in_path, temp_path)

    # --- teraz wczytaj i zmodyfikuj tylko to, co trzeba ---
    tree = ET.parse(temp_path)
    root = tree.getroot()

    for offer in root.findall("o"):
        # zmiana dostępności: aktywne tylko od 5 sztuk
        qty_attr = offer.get("stock", "0")
        avail = offer.get("avail", "99")

        try:
            qty = int(qty_attr) if qty_attr.isdigit() else int(float(qty_attr))
        except:
            qty = 0

        if avail == "1" and qty < 5:
            offer.set("avail", "99")
            offer.set("stock", "0")
            offer.set("basket", "0")

        # usuń <desc_json>
        for desc_json in offer.findall("desc_json"):
            offer.remove(desc_json)

    # --- zapis do pliku docelowego ---
    ET.indent(root, space="  ")
    tree.write(out_path, encoding="utf-8", xml_declaration=True)
    os.remove(temp_path)
    print(f"[Morele OK] Zapisano: {out_path}")

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    for name in os.listdir(INPUT_DIR):
        if name.lower().endswith((".xlsm", ".xlsx", ".xls")):
            src = os.path.join(INPUT_DIR, name)
            dst = os.path.join(OUTPUT_DIR, "morele.xml")
            print(f"[Morele] {src} -> {dst}")
            convert_file_morele(src, dst)
            break

if __name__ == "__main__":
    main()
