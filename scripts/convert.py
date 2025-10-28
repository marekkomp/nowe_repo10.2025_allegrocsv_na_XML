import os
import re
import csv
import xml.etree.ElementTree as ET

INPUT_DIR = "input"
OUTPUT_DIR = "output"

META_EXCLUDE = {
    "ID oferty","Link do oferty","Akcja","Status oferty",
    "ID produktu (EAN/UPC/ISBN/ISSN/ID produktu Allegro)",
    "Kategoria główna","Podkategoria","Sygnatura/SKU Sprzedającego",
    "Liczba sztuk","Reguła Cenowa (PL)","Cena PL",
    "Reguła Cenowa (CZ)","Cena CZ","Reguła Cenowa (SK)","Cena SK",
    "Reguła Cenowa (HU)","Cena HU","Tytuł oferty","Zdjęcia",
    "Opis oferty","Cennik dostawy","Czas wysyłki","Kraj","Województwo",
    "Kod pocztowy","Miejscowość","Opcje faktury","Przedmiot oferty",
    "Stawki VAT","Podstawa wyłączenia z VAT","Warunki zwrotów",
    "Warunki reklamacji","Informacje o gwarancjach (opcjonalne)",
    "Termin wprowadzenia produktu na rynek UE",
    "Osoba odpowiedzialna za zgodność produktu",
    "Informacje o bezpieczeństwie"
}

ATTR_KEYWORDS = [
    "model","producent","seria","procesor","rdzeni","pamięć","ram","dysk",
    "ssd","hdd","pojemność","ekran","matryc","rozdziel","graf","karta",
    "złącza","port","hdmi","vesa","kolor","system","napęd","kamera",
    "mikrofon","materiał","zasil","taktowanie","interfejs","format",
    "standard","waga","bateria","akumulator"
]

def get(row, *names):
    for n in names:
        if n in row:
            val = str(row[n]).strip()
            if val:
                return val
    return ""

def is_attr_header(h):
    if not h or h in META_EXCLUDE:
        return False
    h_low = h.lower()
    return (
        h_low.endswith("_dict")
        or h_low.endswith("_text")
        or any(k in h_low for k in ATTR_KEYWORDS)
        or "[" in h and "]" in h
    )

def id_from_url(url):
    m = re.search(r"/(\d+)(?:[^0-9]?)*$", url.strip()) if url else None
    return m.group(1) if m else ""

def build_attrs(parent, row):
    attrs = ET.SubElement(parent, "attrs")
    for h, v in row.items():
        if v and is_attr_header(h):
            ET.SubElement(attrs, "a", {"name": h}).text = str(v).strip()

def build_offer(row):
    url = get(row, "Link do oferty")
    oid = get(row, "ID oferty") or id_from_url(url)
    price = get(row, "Cena PL") or "0"
    status = get(row, "Status oferty").lower()
    qty = get(row, "Liczba sztuk")
    qty_ok = qty.isdigit() and int(qty) > 0

    avail = "1" if (status.startswith("aktyw") and qty_ok) else "99"
    stock = "1" if qty_ok else "0"
    basket = "1" if avail == "1" else "0"

    o = ET.Element("o", {
        "id": oid,
        "url": url,
        "price": price,
        "avail": avail,
        "stock": stock,
        "basket": basket
    })
    ET.SubElement(o, "cat").text = get(row, "Kategoria główna")
    ET.SubElement(o, "name").text = get(row, "Tytuł oferty")

    photos = get(row, "Zdjęcia")
    imgs = ET.SubElement(o, "imgs")
    if "|" in photos:
        parts = [p.strip() for p in photos.split("|") if p.strip()]
        if parts:
            ET.SubElement(imgs, "main", {"url": parts[0]})
            for p in parts[1:]:
                ET.SubElement(imgs, "i", {"url": p})

    build_attrs(o, row)
    return o

def detect_header_row(rows):
    for i, r in enumerate(rows[:200]):
        line = " ".join(str(x).lower() for x in r if x)
        if "id oferty" in line and "tytuł" in line:
            return i
    return 3  # fallback

def read_excel_rows(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    all_data = []
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        header_idx = detect_header_row(rows)
        headers = [str(h).strip() if h else "" for h in rows[header_idx]]
        data = []
        for r in rows[header_idx + 1:]:
            row = {headers[i]: (r[i] if i < len(r) and r[i] else "") for i in range(len(headers))}
            if any(str(v).strip() for v in row.values()):
                data.append(row)
        print(f"[INFO] Arkusz '{ws.title}': nagłówki w wierszu {header_idx+1}, rekordów: {len(data)}")
        if len(data) > 0:
            all_data.extend(data)
    return all_data

def read_csv_rows(path):
    delim = ";"
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        lines = f.readlines()
    header_line = detect_header_row([l.split(delim) for l in lines[:200]])
    from io import StringIO
    reader = csv.DictReader(StringIO("".join(lines[header_line:])), delimiter=delim)
    return [r for r in reader if any(v.strip() for v in r.values())]

def convert_file(in_path, out_path):
    ext = os.path.splitext(in_path)[1].lower()
    rows = read_excel_rows(in_path) if ext in (".xls", ".xlsx", ".xlsm") else read_csv_rows(in_path)
    root = ET.Element("offers")
    for r in rows:
        root.append(build_offer(r))
    ET.indent(root, space="  ")
    ET.ElementTree(root).write(out_path, encoding="utf-8", xml_declaration=True)
    print(f"[OK] Zapisano {out_path} ({len(rows)} ofert)")

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    for name in os.listdir(INPUT_DIR):
        if not name.lower().endswith((".xls", ".xlsx", ".xlsm", ".csv")):
            continue
        src = os.path.join(INPUT_DIR, name)
        dst = os.path.join(OUTPUT_DIR, os.path.splitext(name)[0] + ".xml")
        print(f"[RUN] Konwersja: {src} -> {dst}")
        try:
            convert_file(src, dst)
        except Exception as e:
            print(f"[ERROR] {src}: {e}")

if __name__ == "__main__":
    main()
