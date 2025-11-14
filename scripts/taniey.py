# scripts/convert_taniey.py
import os
import re
import json
from lxml import etree as ET  # używamy lxml (obsługuje CDATA)
from convert import convert_file, INPUT_DIR, OUTPUT_DIR  # główny konwerter

# --------- USTAWIENIA ---------
FOOTER_MARK = "<!---->"      # znacznik, by nie dublować
LINKS_AS_PLAIN_TEXT = True   # linki w stopce jako zwykły tekst (bez <a>)

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

def _screen_inch(attrs):
    key = next((k for k in attrs.keys() if k.lower().startswith("przekątna ekranu")), None)
    if not key:
        return None
    txt = (attrs.get(key) or "").strip()
    m = re.search(r"(\d+(?:[.,]\d+)?)", txt)
    if not m:
        return None
    try:
        return float(m.group(1).replace(",", "."))
    except ValueError:
        return None

def _laptop_size_url(size_in):
    if size_in is None:
        return None
    s = size_in
    if s <= 12.5:
        return "https://kompre.pl/pl/c/Laptopy-12-cali/349"
    if 13.0 <= s <= 13.4:
        return "https://kompre.pl/pl/c/Laptopy-13-cali/394"
    if 14.0 <= s <= 14.15:
        return "https://kompre.pl/pl/c/Laptopy-14-cali/350"
    if 15.5 <= s <= 15.7:
        return "https://kompre.pl/pl/c/Laptopy-15-cali/351"
    if 16.9 <= s <= 17.35:
        return "https://kompre.pl/pl/c/Laptopy-17-cali/352"
    return None

def _build_link_block(kategoria, attrs):
    if not kategoria:
        return ""
    if "laptop" not in kategoria.lower():
        return ""

    size_in = _screen_inch(attrs)
    url = _laptop_size_url(size_in)
    if not url:
        return ""

    if size_in:
        size_txt = f"{size_in:.1f}".rstrip("0").rstrip(".")
    else:
        size_txt = ""

    return (
        f"<p>Sprawdź też inne modele laptopów z rozmiarem ekranu {size_txt}″: {url}. "
        f"Każdy komputer jest dokładnie sprawdzany, czyszczony i konfigurowany, aby zapewnić niezawodność w codziennym użytkowaniu. "
        f"Kupując sprzęt, zyskujesz jakość klasy biznes oraz pewność gwarancji door-to-door.</p>"
    )

def _build_footer_html(name, kategoria, attrs):
    link_block = _build_link_block(kategoria, attrs)
    return (
        f'{FOOTER_MARK}'
        f'<hr/><p><strong>{name}</strong> pochodzi z oferty <strong>Kompre.pl</strong> – '
        f'największego i autoryzowanego sprzedawcy biznesowych sprzętów outletowych, laptopów, komputerów PC i monitorów.</p> '
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

def _append_footer_to_desc(o_el):
    """Dopisuje stopkę do <desc> (HTML) i zapisuje w CDATA."""
    desc_el = o_el.find("desc")
    if desc_el is None:
        return

    current_html = _inner_html(desc_el)
    if _already_has_footer(current_html):
        return

    attrs = _collect_attrs(o_el)
    name = _name(o_el)
    kategoria = _category(o_el)

    footer_html = _build_footer_html(name, kategoria, attrs)
    joiner = "\n" if current_html and not current_html.endswith("\n") else ""
    new_html = f"{current_html}{joiner}{footer_html}".strip()
    _set_desc_cdata(desc_el, new_html)

def _append_footer_to_desc_json(o_el):
    """Dopisuje stopkę także do <desc_json> jako dodatkowy blok TEXT."""
    dj = o_el.find("desc_json")
    if dj is None or not (dj.text or "").strip():
        return

    raw = dj.text.strip()
    if FOOTER_MARK in raw:
        return  # już dopięte (po markerze)

    try:
        data = json.loads(raw)
    except Exception:
        return  # nieprawidłowy JSON — nie dotykamy

    # Przygotuj HTML stopki
    attrs = _collect_attrs(o_el)
    name = _name(o_el)
    kategoria = _category(o_el)
    footer_html = _build_footer_html(name, kategoria, attrs)

    def _append_text_block(sections_list):
        if not sections_list:
            sections_list.append({"items": [{"type": "TEXT", "content": footer_html}]})
            return
        # dopnij jako nowy blok, żeby nie mieszać z istniejącą treścią
        sections_list.append({"items": [{"type": "TEXT", "content": footer_html}]})

    if isinstance(data, dict):
        sections = data.get("sections")
        if isinstance(sections, list):
            _append_text_block(sections)
        else:
            data["sections"] = [{"items": [{"type": "TEXT", "content": footer_html}]}]
    elif isinstance(data, list):
        # rzadziej spotykane: JSON to lista sekcji
        _append_text_block(data)
    else:
        # nieobsługiwany kształt — nie modyfikujemy
        return

    # Zapisz z powrotem (bez ASCII-escape, z zachowaniem PL znaków)
    dj.text = json.dumps(data, ensure_ascii=False)

# --------- GŁÓWNA LOGIKA ---------
def convert_file_taniey(in_path, out_path):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    temp_path = os.path.join(OUTPUT_DIR, "_temp_base.xml")
    convert_file(in_path, temp_path)

    parser = ET.XMLParser(remove_blank_text=True)
    tree = ET.parse(temp_path, parser)
    root = tree.getroot()

    for o in root.findall("o"):
        # dostępność: aktywna tylko gdy stock >= 10
        try:
            stock_num = int(o.get("stock", "0"))
        except:
            try:
                stock_num = int(float(o.get("stock", "0")))
            except:
                stock_num = 0

        if o.get("avail") == "1" and stock_num < 10:
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

        # UWAGA: NIE USUWAMY już desc_json — zostaje w XML

        attrs_el = o.find("attrs")
        if attrs_el is not None:
            attrs = _collect_attrs(o)

            # Dodaj <a name="Marka"> jeśli brak, z wartością z Producent
            producent = attrs.get("Producent", "").strip()
            if producent and not any((a.get("name") or "").strip() == "Marka" for a in attrs_el.findall("a")):
                ET.SubElement(attrs_el, "a", {"name": "Marka"}).text = producent

            # Zamiana Przekątna ekranu ["] -> Przekątna ekranu (")
            for a in attrs_el.findall("a"):
                if (a.get("name") or "").strip() == 'Przekątna ekranu ["]':
                    a.set("name", 'Przekątna ekranu (")')

            # Gwarancja (jak wcześniej)
            for a in attrs_el.findall("a"):
                if (a.get("name") or "").strip().lower() == "informacje o gwarancjach":
                    text = (a.text or "").strip()
                    m = re.search(r"(\d+)", text)
                    value = m.group(1) if m else ""
                    a.set("name", "Gwarancja")
                    a.text = value

        # Dopnij stopkę do HTML
        _append_footer_to_desc(o)
        # Dopnij stopkę również do JSON-a
        _append_footer_to_desc_json(o)

    tree.write(out_path, encoding="utf-8", xml_declaration=True, pretty_print=True)
    try:
        os.remove(temp_path)
    except FileNotFoundError:
        pass

    print(f"[taniey OK] Zapisano: {out_path}")

def main():
    for name in os.listdir(INPUT_DIR):
        if name.lower().endswith((".xlsm", ".xlsx", ".xls")):
            src = os.path.join(INPUT_DIR, name)
            dst = os.path.join(OUTPUT_DIR, "taniey.xml")
            print(f"[taniey] {src} -> {dst}")
            convert_file_taniey(src, dst)
            break

if __name__ == "__main__":
    main()
