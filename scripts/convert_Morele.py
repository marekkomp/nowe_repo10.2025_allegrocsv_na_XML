import os
from convert_base import convert_file, INPUT_DIR, OUTPUT_DIR, _is_available, _looks_like_json, _desc_to_html
import openpyxl
import xml.etree.ElementTree as ET

def convert_file_morele(in_path, out_path):
    wb = openpyxl.load_workbook(in_path, data_only=True)
    ws = wb["Szablon"] if "Szablon" in wb.sheetnames else wb.worksheets[0]
    headers = [str(c.value).strip() if c.value else "" for c in ws[4]]
    rows = list(ws.iter_rows(min_row=5, values_only=True))

    i_id   = headers.index("ID oferty")
    i_tyt  = headers.index("Tytuł oferty")
    i_cena = headers.index("Cena PL")
    i_url  = headers.index("Link do oferty")
    i_stat = headers.index("Status oferty")
    i_qty  = headers.index("Liczba sztuk")
    i_cat  = headers.index("Kategoria główna")
    i_imgs = headers.index("Zdjęcia")
    i_desc = headers.index("Opis oferty")

    root = ET.Element("offers")

    for row in rows:
        id_offer = str(row[i_id]).strip()
        title = str(row[i_tyt]).strip()
        price = str(row[i_cena]).strip()
        url = str(row[i_url]).strip()
        status = str(row[i_stat]).strip().lower()
        qty = float(row[i_qty]) if row[i_qty] else 0
        cat = str(row[i_cat]).strip()
        imgs_raw = str(row[i_imgs]) if row[i_imgs] else ""
        desc_raw = str(row[i_desc]).strip() if row[i_desc] else ""

        if not id_offer or not title:
            continue

        # 🟡 Nowa logika dostępności:
        # - aktywna i qty >= 5 → dostępna
        # - reszta → niedostępna
        available = (status == "aktywna") and (qty >= 5)
        avail_val = "1" if available else "99"
        stock     = "1" if available else "0"
        basket    = "1" if available else "0"

        o = ET.SubElement(root, "o", {
            "id": id_offer,
            "url": url,
            "price": price,
            "avail": avail_val,
            "stock": stock,
            "basket": basket,
        })

        ET.SubElement(o, "cat").text = cat or ""
        ET.SubElement(o, "name").text = title

        # 🟣 Usuwamy desc_json całkowicie
        if desc_raw:
            desc_el = ET.SubElement(o, "desc")
            desc_el.text = _desc_to_html(desc_raw, strict=True)

    ET.indent(root, space="  ")
    ET.ElementTree(root).write(out_path, encoding="utf-8", xml_declaration=True)
    print(f"[Morele OK] Zapisano {out_path}")

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    for name in os.listdir(INPUT_DIR):
        if name.lower().endswith((".xlsx", ".xlsm", ".xls")):
            src = os.path.join(INPUT_DIR, name)
            dst = os.path.join(OUTPUT_DIR, "morele.xml")
            print(f"[Morele] {src} -> {dst}")
            convert_file_morele(src, dst)
            break

if __name__ == "__main__":
    main()
