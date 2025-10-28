# scripts/convert.py
import os
import re
import csv
import xml.etree.ElementTree as ET

INPUT_DIR = "input"
OUTPUT_DIR = "output"

# Kolumny nielądujące do <attrs> (logistyka/ceny/warunki itp.)
META_EXCLUDE = {
    "ID oferty", "ID oferty ", "Link do oferty", "Link", "URL", "url",
    "Akcja", "Status oferty",
    "ID produktu (EAN/UPC/ISBN/ISSN/ID produktu Allegro)",
    "Kategoria główna", "Podkategoria", "Sygnatura/SKU Sprzedającego",
    "Liczba sztuk",
    "Reguła Cenowa (PL)", "Cena PL", "Reguła Cenowa (CZ)", "Cena CZ",
    "Reguła Cenowa (SK)", "Cena SK", "Reguła Cenowa (HU)", "Cena HU",
    "Tytuł oferty", "Zdjęcia", "Opis oferty", "Opis oferty (HTML)", "opis_oferty",
    "Cennik dostawy", "Czas wysyłki", "Kraj", "Województwo",
    "Kod pocztowy", "Miejscowość", "Opcje faktury", "Przedmiot oferty",
    "Stawki VAT", "Podstawa wyłączenia z VAT", "Warunki zwrotów",
    "Warunki reklamacji", "Informacje o gwarancjach (opcjonalne)",
    "Termin wprowadzenia produktu na rynek UE",
    "Osoba odpowiedzialna za zgodność produktu",
    "Informacje o bezpieczeństwie",
}

# Heurystyka: które nagłówki traktować jako parametry/atrybuty
ATTR_KEYWORDS = [
    "model", "producent", "seria", "procesor", "rdzeni", "pamięć", "ram",
    "dysk", "ssd", "hdd", "pojemność", "ekran", "matryc", "przekątna",
    "rozdziel", "powłoka", "graf", "karta", "układ", "złącza", "port",
    "hdmi", "vesa", "kolor", "system", "napęd", "kamera", "mikrofon",
    "wifi", "bluetooth", "materiał", "zasil", "taktowanie", "interfejs",
    "format", "standard", "waga", "wymiary", "bateria", "akumulator"
]

# -------------------- Pomocnicze --------------------

def sniff_delimiter(path: str) -> str:
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        sample = "".join([next(f, "") for _ in range(10)])
    return csv.Sniffer().sniff(sample, delimiters=";,|\t,").delimiter

def get(row: dict, *names: str) -> str:
    for n in names:
        if n in row:
            val = row[n]
            if val is None:
                continue
            val = str(val).strip()
            if val != "":
                return val
    return ""

def id_from_url(url: str) -> str:
    if not url:
        return ""
    m = re.search(r"/(\d+)(?:[^0-9]?)*$", url.strip())
    return m.group(1) if m else ""

def is_attr_header(h: str) -> bool:
    if not h:
        return False
    if h in META_EXCLUDE:
        return False
    low = h.lower()
    return (
        low.endswith("_dict")
        or low.endswith("_text")
        or any(k in low for k in ATTR_KEYWORDS)
        or ("[" in h and "]" in h)  # nagłówki z jednostkami, np. [GB], [mm], itp.
    )

# -------------------- Budowa XML --------------------

def build_attrs(parent: ET.Element, row: dict) -> None:
    attrs = ET.SubElement(parent, "attrs")
    for h, v in row.items():
        if v is None:
            continue
        val = str(v).strip()
        if not val:
            continue
        if is_attr_header(h):
            ET.SubElement(attrs, "a", {"name": h}).text = val

def build_offer_node(row: dict) -> ET.Element:
    url   = get(row, "Link do oferty", "URL", "url", "Link")
    oid   = get(row, "ID oferty", "ID oferty ") or id_from_url(url) or get(row, "Sygnatura/SKU Sprzedającego")
    price = get(row, "Cena PL", "cena_pl", "Cena")
    status = get(row, "Status oferty", "status_oferty", "Status")
    qty    = get(row, "Liczba sztuk", "liczba_sztuk", "Ilość", "ilosc")

    active = status.lower().startswith("aktyw") if status else False
    qty_ok = qty.isdigit() and int(qty) > 0 if qty else False

    # Zasady użytkownika:
    # - avail: 1 gdy Aktywna i qty>0, inaczej 99
    # - stock: 1 gdy qty>0, inaczej 0
    # - basket: 1 gdy avail=1, inaczej 0
    avail  = "1" if (active and qty_ok) else "99"
    stock  = "1" if qty_ok else "0"
    basket = "1" if avail == "1" else "0"

    o = ET.Element("o", {
        "id": oid,
        "url": url,
        "price": price or "0",
        "avail": avail,
        "stock": stock,
        "basket": basket
    })

    cat  = get(row, "Kategoria główna", "kategoria_główna")
    name = get(row, "Tytuł oferty", "tytuł_oferty", "Tytuł", "tytul")
    ET.SubElement(o, "cat").text = cat
    ET.SubElement(o, "name").text = name

    photos = get(row, "Zdjęcia", "zdjęcia")
    photo_list = [u.strip() for u in photos.split("|")] if photos else []
    photo_list = [u for u in photo_list if u]
    imgs = ET.SubElement(o, "imgs")
    if photo_list:
        ET.SubElement(imgs, "main", {"url": photo_list[0]})
        for p in photo_list[1:]:
            ET.SubElement(imgs, "i", {"url": p})

    build_attrs(o, row)
    return o

# -------------------- Odczyt CSV/XLSX/XLSM --------------------

def detect_header_row(rows) -> int:
    """
    Znajdź indeks wiersza z nagłówkami skanując pierwsze ~30 wierszy.
    Szukamy fraz typu 'ID oferty' lub 'Tytuł oferty'.
    """
    max_scan = min(len(rows), 30)
    for i in range(max_scan):
        r = rows[i]
        cells = [str(c).strip().lower() if c is not None else "" for c in r]
        line = " | ".join(cells)
        if "id oferty" in line or "tytuł oferty" in line or "tytuł_oferty" in line:
            return i
    # fallback – Allegro klasycznie ma 3 wiersze techniczne
    return 3 if len(rows) > 3 else 0

def read_excel_rows(path: str) -> list:
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header_idx = detect_header_row(rows)
    headers = [str(h).strip() if h is not None else "" for h in rows[header_idx]]
    data = []
    for r in rows[header_idx + 1:]:
        row_dict = {headers[i]: (str(r[i]).strip() if i < len(headers) and r[i] is not None else "")
                    for i in range(len(headers))}
        if any(val for val in row_dict.values()):
            data.append(row_dict)
    print(f"[INFO] XLSX/XLSM: nagłówki w wierszu {header_idx+1}, rekordów: {len(data)}")
    return data

def read_csv_rows(path: str) -> list:
    delim = sniff_delimiter(path)
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        lines = f.readlines()

    # znajdź linię z nagłówkami
    header_line_idx = None
    max_scan = min(len(lines), 30)
    for i in range(max_scan):
        low = lines[i].lower()
        if ("id oferty" in low) or ("tytuł oferty" in low) or ("tytuł_oferty" in low):
            header_line_idx = i
            break
    if header_line_idx is None:
        header_line_idx = 3 if len(lines) > 3 else 0

    from io import StringIO
    buf = StringIO("".join(lines[header_line_idx:]))

    reader = csv.DictReader(buf, delimiter=delim)
    rows = []
    for r in reader:
        # Normalizacja: usuń None, przytnij spacje
        clean = { (k or "").strip(): (str(v).strip() if v is not None else "") for k, v in r.items() }
        if any(v for v in clean.values()):
            rows.append(clean)

    print(f"[INFO] CSV: nagłówki w linii {header_line_idx+1}, separator '{delim}', rekordów: {len(rows)}")
    return rows

# -------------------- Konwersja pliku --------------------

def convert_file(in_path: str, out_path: str) -> int:
    ext = os.path.splitext(in_path)[1].lower()
    if ext in (".xls", ".xlsx", ".xlsm"):
        rows = read_excel_rows(in_path)
    else:
        rows = read_csv_rows(in_path)

    root = ET.Element("offers")
    count = 0
    for row in rows:
        node = build_offer_node(row)
        root.append(node)
        count += 1

    tree = ET.ElementTree(root)
    ET.indent(tree, space="  ", level=0)
    tree.write(out_path, encoding="utf-8", xml_declaration=True)
    return count

# -------------------- Główne wejście --------------------

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    found = False
    total = 0
    for name in os.listdir(INPUT_DIR):
        ext = name.lower().rsplit(".", 1)[-1]
        if ext not in ("csv", "xls", "xlsx", "xlsm"):
            continue
        found = True
        src = os.path.join(INPUT_DIR, name)
        dst = os.path.join(OUTPUT_DIR, os.path.splitext(name)[0] + ".xml")
        print(f"[RUN] Konwersja: {src} -> {dst}")
        try:
            n = convert_file(src, dst)
            total += n
            print(f"[OK] Zapisano {dst} (ofert: {n})")
        except Exception as e:
            print(f"[ERROR] {src}: {e}")

    if not found:
        print("[INFO] Brak plików wejściowych w /input")
    else:
        print(f"[SUMMARY] Łącznie ofert: {total}")

if __name__ == "__main__":
    main()
