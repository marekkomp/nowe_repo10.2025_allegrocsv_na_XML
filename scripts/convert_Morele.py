# scripts/convert_Morele.py
import os
import re
import html as _html
from lxml import etree as ET  # używamy lxml (obsługuje CDATA)
from convert import convert_file, INPUT_DIR, OUTPUT_DIR  # główny konwerter

# --------- USTAWIENIA ---------
BRAND_LINKS = {
    "dell":   "https://kompre.pl/pl/c/Laptopy-Dell/364",
    "lenovo": "https://kompre.pl/pl/c/Laptopy-Lenovo/366",
    "hp":     "https://kompre.pl/pl/c/Laptopy-HP/365",
    "apple":  "https://kompre.pl/pl/c/Laptopy-Apple/367",
    "fujitsu":"https://kompre.pl/pl/c/Laptopy-Fujitsu/368",
}
FOOTER_MARK = "<!---->"
LINKS_AS_PLAIN_TEXT = True

# --------- POMOCNICZE ---------
def _collect_attrs(o_el):
    out = {}
    attrs_el = o_el.find("attrs")
    if attrs_el is None:
        return out
    for a in attrs_el.findall("a"):
        name = (a.get("name") or "").strip()
        val = (a.text or "").strip()
        if name:
            out[name] = val
    return out

def _brand(attrs):
    return (attrs.get("Producent") or "").strip()

def _warranty(attrs):
    gw = (attrs.get("Informacje o gwarancjach") or attrs.get("Gwarancja") or "").strip()
    return gw

def _category(o_el):
    cat_el = o_el.find("cat")
    return (cat_el.text or "").strip() if cat_el is not None else ""

def _name(o_el):
    name_el = o_el.find("name")
    return (name_el.text or "").strip() if name_el is not None else ""

def _build_link_block(kategoria, producent):
    if not kategoria or not producent:
        return ""
    kat = kategoria.lower()
    brand = producent.lower()
    if "laptop" in kat:
        url = BRAND_LINKS.get(brand)
    elif "komputer" in kat:
        url = "https://kompre.pl/pl/c/Komputery-Stacjonarne/345"
    elif "monitor" in kat:
        url = "https://kompre.pl/monitory"
    else:
        return ""
    if not url:
        return ""
    return (
        f"<p>Posiadamy też inne modele {producent} – sprawdź: {url}. "
        f"Każdy egzemplarz jest testowany, czyszczony i przygotowany do pracy z aktualnym systemem. "
        f"Długa gwarancja door-to-door zapewnia wsparcie i bezpieczeństwo zakupu.</p>"
    )

def _build_footer_html(name, producent, gwarancja, kategoria):
    link_block = _build_link_block(kategoria, producent)
    gw = (gwarancja or "12 miesięcy").strip()
    gwarancja_txt = gw if "gwarancja" in gw.lower() else f"Gwarancja {gw}"
    return (
        f'{FOOTER_MARK}'
        f'<hr/><p><strong>{name}</strong> pochodzi z oferty <strong>Kompre.pl</strong> – '
        f'autoryzowanego sprzedawcy komputerów poleasingowych klasy biznes.</p> '
        f'{link_block}'
    )

def _inner_html(el: ET.Element) -> str:
    parts = []
    if el.text:
        parts.append(el.text)
    for c in el:
        parts.append(ET.tostring(c, encoding="unicode"))
    return "".join(parts)

def _set_desc_cdata(desc_el: ET.Element, html_string: str):
    desc_el.clear()
    desc_el.text = ET.CDATA(html_string)

def _already_has_footer(html: str) -> bool:
    return (FOOTER_MARK in html) or ("Kompre.pl" in html and "door-to-door" in html)

# --- NOWE: wyjątek dla laptopów Toshiba (bez link-block i stopki)
def _is_toshiba_laptop(producent: str, kategoria: str) -> bool:
    p = (producent or "").strip().lower()
    k = (kategoria or "").strip().lower()
    return p.startswith("toshiba") and ("laptop" in k)

def _append_footer_to_desc(o_el):
    desc_el = o_el.find("desc")
    if desc_el is None:
        return
    current_html = _inner_html(desc_el)
    if _already_has_footer(current_html):
        return

    attrs = _collect_attrs(o_el)
    name = _name(o_el)
    producent = _brand(attrs)
    kategoria = _category(o_el)

    # pomiń dodawanie wstawek dla laptopów Toshiba
    if _is_toshiba_laptop(producent, kategoria):
        return

    gwarancja = _warranty(attrs)
    footer_html = _build_footer_html(name, producent, gwarancja, kategoria)
    joiner = "\n" if current_html and not current_html.endswith("\n") else ""
    new_html = f"{current_html}{joiner}{footer_html}".strip()
    _set_desc_cdata(desc_el, new_html)

# --- Sanizacja i CDATA dla istniejącego HTML (opakowanie) ---
# Wycinamy <script>, <iframe>, <img> (Morele niech ma opis bez obrazków)
_SCRIPT_IFRAME_IMG_RE = re.compile(
    r"(?is)<script.*?</script>|<iframe.*?</iframe>|<img\b[^>]*>"
)

def _sanitize_basic(html_str: str) -> str:
    return _SCRIPT_IFRAME_IMG_RE.sub("", html_str or "")

def _has_html_tags(s: str) -> bool:
    return bool(re.search(r"<[a-zA-Z][^>]*>", s or ""))

# --- Edycje copy w opisie (reguły) ---
def _apply_copy_edits(s: str) -> str:
    rules = [
        (re.compile(r'(?i)Nawiąż kontakt z kim tylko chcesz'), 'Nawiąż znajomość z kim tylko chcesz'),
        (re.compile(r'(?i)Świetny stosunek jakości do ceny'), 'Świetna jakość'),
        (re.compile(r'(?i)\bw\s+gratisie\b'), ''),     # usuń "w Gratisie"
        (re.compile(r'(?i)\bgratis!?\b'), ''),        # usuń "Gratis" / "GRATIS!"
        (re.compile(r'(?i)Nie tylko cena,\s*'), ''),  # usuń "Nie tylko cena,"
        (re.compile(r'(?i)\bcenie\b'), 'ofercie'),    # zamień "cenie" na "ofercie"
        (re.compile(r'(?i)\bcena\b'), 'ofercie'),     # (zgodnie z wcześniejszym zapisem)
        (re.compile(r'(?i)Kup teraz'), ''),           # usuń "Kup teraz"
    ]
    out = s
    for rx, repl in rules:
        out = rx.sub(repl, out)
    out = re.sub(r'\s{2,}', ' ', out)
    return out.strip()

def _force_desc_cdata(o_el: ET.Element):
    """Opis w realnym HTML (CDATA), bez <img>, z poprawkami copy."""
    desc_el = o_el.find("desc")
    if desc_el is None:
        return
    raw = _inner_html(desc_el).strip()
    unescaped = _html.unescape(raw).strip()
    cleaned = _sanitize_basic(unescaped)
    cleaned = _apply_copy_edits(cleaned)
    if not _has_html_tags(cleaned) and cleaned:
        cleaned = f"<p>{cleaned}</p>"
    _set_desc_cdata(desc_el, cleaned)

# --- Formatowanie pojemności ---
def _format_capacity_unit(val: str) -> str:
    if not val:
        return ""
    m = re.search(r"(\d+(?:[.,]\d+)?)", val)
    if not m:
        return ""
    num = m.group(1).replace(",", ".")
    try:
        f = float(num)
    except:
        return ""
    i = int(round(f))
    return "1 TB" if i == 1 else f"{i} GB"

# --- Normalizacja cali w 'Przekątna ekranu' ---
def _normalize_inches(value: str) -> str:
    """Zwraca N[.N]\" (np. 14\", 12.5\"). Usuwa 'cali' itp., dokleja jeśli brak."""
    if not value:
        return value
    m = re.search(r"(\d+(?:[.,]\d+)?)", value)
    if not m:
        v = value.strip()
        return v if v.endswith('"') else (v + '"')
    num = m.group(1).replace(",", ".")
    if "." in num:
        num = num.rstrip("0").rstrip(".")
    return f'{num}"'

# --------- GŁÓWNA LOGIKA KONWERSJI ---------
def convert_file_morele(in_path, out_path):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    temp_path = os.path.join(OUTPUT_DIR, "_temp_base.xml")
    convert_file(in_path, temp_path)

    parser = ET.XMLParser(remove_blank_text=True)
    tree = ET.parse(temp_path, parser)
    root = tree.getroot()

    for o in root.findall("o"):
        # dostępność: aktywna tylko gdy stock >= 5
        try:
            stock_num = int(o.get("stock", "0"))
        except:
            try:
                stock_num = int(float(o.get("stock", "0")))
            except:
                stock_num = 0
        if o.get("avail") == "1" and stock_num < 5:
            o.set("avail", "99")
            o.set("stock", "0")
            o.set("basket", "0")

        # dopisz "poleasingowe" do kategorii
        cat_el = o.find("cat")
        if cat_el is not None and cat_el.text:
            cat_text = cat_el.text.strip()
            norm = cat_text.lower()
            if "poleasingowe" not in norm:
                if norm == "laptopy":
                    cat_el.text = "Laptopy poleasingowe"
                elif norm == "komputery":
                    cat_el.text = "Komputery poleasingowe"
                elif norm == "monitory komputerowe":
                    cat_el.text = "Monitory poleasingowe"

        # usuń desc_json (Morele korzysta z HTML)
        for dj in o.findall("desc_json"):
            parent = dj.getparent() if hasattr(dj, "getparent") else o
            parent.remove(dj)

        # --- ATRYBUTY: transformacje dla Morele ---
        attrs_el = o.find("attrs")
        if attrs_el is not None:
            # słownik atrybutów
            attrs = {}
            for a in attrs_el.findall("a"):
                name = (a.get("name") or "").strip()
                val = (a.text or "").strip()
                if name:
                    attrs[name] = val

            # 1) Stan: Używany -> Poleasingowy
            for a in attrs_el.findall("a"):
                if (a.get("name") or "").strip() == "Stan":
                    val = (a.text or "").strip()
                    if re.search(r"\bużywany\b", val, flags=re.IGNORECASE):
                        a.text = "Poleasingowy"

            # 2) Zmiany nazw atrybutów RAM / ekran / rozdzielczość
            for a in attrs_el.findall("a"):
                n = (a.get("name") or "").strip()
                if n == 'Wielkość pamięci RAM':
                    a.set("name", "Pamięć RAM (zainstalowana)")
                elif n == 'Przekątna ekranu ["]':
                    a.set("name", "Przekątna ekranu")
                elif n == 'Rozdzielczość (px)':
                    a.set("name", "Rozdzielczość")

            # 2a) Przekątna ekranu – wymuś format N[.N]"
            for a in attrs_el.findall("a"):
                if (a.get("name") or "").strip() == "Przekątna ekranu":
                    v = (a.text or "").strip()
                    if v:
                        a.text = _normalize_inches(v)

            # 3) Ekran dotykowy: tylko "Nie" lub "z ekranem dotykowym"
            for a in attrs_el.findall("a"):
                if (a.get("name") or "").strip() == "Ekran dotykowy":
                    v = (a.text or "").strip().lower()
                    if v == "tak":
                        a.text = "z ekranem dotykowym"
                    elif v == "nie":
                        a.text = "Nie"

            # 4) Dyski SSD/HDD z pojemnością
            typ = (attrs.get("Typ dysku twardego") or "").lower()
            cap_raw = attrs.get("Pojemność dysku [GB]") or ""
            cap_fmt = _format_capacity_unit(cap_raw)
            if cap_fmt:
                if "ssd" in typ and not any((x.get("name") or "") == "Dysk SSD" for x in attrs_el.findall("a")):
                    ET.SubElement(attrs_el, "a", {"name": "Dysk SSD"}).text = cap_fmt
                if "hdd" in typ and not any((x.get("name") or "") == "Dysk HDD" for x in attrs_el.findall("a")):
                    ET.SubElement(attrs_el, "a", {"name": "Dysk HDD"}).text = cap_fmt

            # 4a) Grafika zintegrowana -> dopisz pamięć karty jako "Współdzielona z RAM"
            rodzaj = (attrs.get("Rodzaj karty graficznej") or "").strip().lower()
            if "zintegrowana" in rodzaj:
                has_mem = any((x.get("name") or "").strip() == "Pamięć karty graficznej" for x in attrs_el.findall("a"))
                if not has_mem:
                    ET.SubElement(attrs_el, "a", {"name": "Pamięć karty graficznej"}).text = "Współdzielona z RAM"

            # 5) Informacje o gwarancjach -> Gwarancja (liczba)
            for a in attrs_el.findall("a"):
                if (a.get("name") or "").strip().lower() == "informacje o gwarancjach":
                    text = (a.text or "").strip()
                    m = re.search(r"(\d+)", text)
                    value = m.group(1) if m else ""
                    a.set("name", "Gwarancja")
                    a.text = value

        # --- OPIS: HTML w CDATA (bez IMG) + poprawki copy + stopka
        _force_desc_cdata(o)
        _append_footer_to_desc(o)

    tree.write(out_path, encoding="utf-8", xml_declaration=True, pretty_print=True)
    try:
        os.remove(temp_path)
    except FileNotFoundError:
        pass

    print(f"[Morele OK] Zapisano: {out_path}")

def main():
    for name in os.listdir(INPUT_DIR):
        if name.lower().endswith((".xlsm", ".xlsx", ".xls")):
            src = os.path.join(INPUT_DIR, name)
            dst = os.path.join(OUTPUT_DIR, "morele.xml")
            print(f"[Morele] {src} -> {dst}")
            convert_file_morele(src, dst)
            break

if __name__ == "__main__":
    main()
