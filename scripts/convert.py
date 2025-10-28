import re, os, csv, xml.etree.ElementTree as ET

INPUT_DIR = "input"
OUTPUT_DIR = "output"

META_EXCLUDE = {
    # kolumny „nie-parametryczne” – nie trafią do <attrs>
    "ID oferty","Link do oferty","Akcja","Status oferty","ID produktu (EAN/UPC/ISBN/ISSN/ID produktu Allegro)",
    "Kategoria główna","Podkategoria","Sygnatura/SKU Sprzedającego","Liczba sztuk",
    "Reguła Cenowa (PL)","Cena PL","Reguła Cenowa (CZ)","Cena CZ","Reguła Cenowa (SK)","Cena SK",
    "Reguła Cenowa (HU)","Cena HU","Tytuł oferty","Zdjęcia","Opis oferty","Cennik dostawy","Czas wysyłki",
    "Kraj","Województwo","Kod pocztowy","Miejscowość","Opcje faktury","Przedmiot oferty","Stawki VAT",
    "Podstawa wyłączenia z VAT","Warunki zwrotów","Warunki reklamacji","Informacje o gwarancjach (opcjonalne)",
    "Termin wprowadzenia produktu na rynek UE","Osoba odpowiedzialna za zgodność produktu","Informacje o bezpieczeństwie",
    "ID oferty ","Link","URL","url","ean","Opis oferty (HTML)","opis_oferty"
}

ATTR_KEYWORDS = [
    "model","producent","seria","procesor","rdzeni","pamięć","ram","dysk","ssd","hdd","pojemność",
    "ekran","matryc","przekątna","rozdziel","powłoka","graf","karta","układ","złącza","port","hdmi","vesa",
    "kolor","system","napęd","kamera","mikrofon","wifi","bluetooth","materiał","zasil","taktowanie",
    "interfejs","format","standard","waga","wymiary","bateria","akumulator"
]

def sniff_delimiter(path):
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        sample = "".join([next(f, "") for _ in range(10)])
    return csv.Sniffer().sniff(sample, delimiters=";,|\t,").delimiter

def get(row, *names):
    for n in names:
        if n in row and str(row[n]).strip() != "":
            return str(row[n]).strip()
    return ""

def id_from_url(url):
    m = re.search(r"/(\d+)(?:[^0-9]?)*$", url.strip())
    return m.group(1) if m else ""

def is_attr_header(h):
    if h in META_EXCLUDE:
        return False
    low = h.lower()
    return (low.endswith("_dict") or low.endswith("_text")
            or any(k in low for k in ATTR_KEYWORDS)
            or "[" in h and "]" in h)  # nagłówki z jednostkami

def build_attrs(parent, row):
    attrs = ET.SubElement(parent, "attrs")
    for h, v in row.items():
        if v is None: 
            continue
        val = str(v).strip()
        if not val:
            continue
        if is_attr_header(h):
            ET.SubElement(attrs, "a", {"name": h}).text = val

def build_o(row):
    url  = get(row, "Link do oferty","URL","url","Link")
    oid  = get(row, "ID oferty","ID oferty ") or id_from_url(url) or get(row, "Sygnatura/SKU Sprzedającego")
    price = get(row, "Cena PL","cena_pl","Cena")
    status = get(row, "Status oferty","status_oferty","Status")
    qty = get(row, "Liczba sztuk","liczba_sztuk","Ilość","ilosc")

    active = status.lower().startswith("aktyw")
    qty_ok = qty.isdigit() and int(qty) > 0 if qty else False

    avail = "1" if (active and qty_ok) else "99"
    stock = "1" if qty_ok else "0"
    basket = "1" if avail == "1" else "0"

    o = ET.Element("o", {
        "id": oid,
        "url": url,
        "price": price or "0",
        "avail": avail,
        "stock": stock,
        "basket": basket
    })

    cat  = get(row, "Kategoria główna","kategoria_główna")
    name = get(row, "Tytuł oferty","tytuł_oferty","Tytuł","tytul")
    ET.SubElement(o, "cat").text = cat
    ET.SubElement(o, "name").text = name

    photos = get(row, "Zdjęcia","zdjęcia")
    pl = [u.strip() for u in photos.split("|") if u.strip()] if photos else []
    imgs = ET.SubElement(o, "imgs")
    if pl:
        ET.SubElement(imgs, "main", {"url": pl[0]})
        for p in pl[1:]:
            ET.SubElement(imgs, "i", {"url": p})

    build_attrs(o, row)
    return o

def convert_file(csv_path, out_path):
    delim = sniff_delimiter(csv_path)
    with open(csv_path, "r", encoding="utf-8", errors="ignore") as f:
        # pomiń 3 wiersze techniczne
        for _ in range(3):
            next(f, None)
        reader = csv.DictReader(f, delimiter=delim)
        root = ET.Element("offers")
        for row in reader:
            if not any((str(v).strip() if v is not None else "") for v in row.values()):
                continue
            root.append(build_o(row))
    tree = ET.ElementTree(root)
    ET.indent(tree, space="  ", level=0)
    tree.write(out_path, encoding="utf-8", xml_declaration=True)

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    done = False
    for name in os.listdir(INPUT_DIR):
        if name.lower().endswith(".csv"):
            done = True
            src = os.path.join(INPUT_DIR, name)
            dst = os.path.join(OUTPUT_DIR, os.path.splitext(name)[0] + ".xml")
            convert_file(src, dst)
            print(f"OK: {src} -> {dst}")
    if not done:
        print("Brak plików CSV w /input")

if __name__ == "__main__":
    main()

