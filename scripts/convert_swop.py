# scripts/convert_swop.py
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

def _get_price(o_el):
    """Zwraca cenę jako float lub None."""
    raw = (o_el.get("price") or "").strip()
    if not raw:
        return None
    try:
        p = float(raw.replace(",", "."))
        return p
    except ValueError:
        return None

def _budget_url(price):
    """Zwraca URL kategorii budżetowej na podstawie ceny."""
    if price is None:
        return None
    if price <= 500:
        return "https://kompre.pl/pl/c/Laptopy-do-500-zl/390"
    if price <= 1000:
        return "https://kompre.pl/pl/c/Laptopy-do-1000-zl/389"
    if price <= 1500:
        return "https://kompre.pl/pl/c/Laptopy-do-1500-zl/391"
    if price <= 2000:
        return "https://kompre.pl/pl/c/Laptopy-do-2000-zl/392"
    if price <= 3000:
        return "https://kompre.pl/pl/c/Laptopy-do-3000-zl/399"
    if price <= 5000:
        return "https://kompre.pl/pl/c/Laptopy-do-5000-zl/500"
    # powyżej 5000 zł też linkujemy tę kategorię
    return "https://kompre.pl/pl/c/Laptopy-do-5000-zl/500"

def _build_link_block(kategoria, attrs, price):
    """
    Link po budżecie:
    'Sprawdź też inne niezawodne laptopy w Twoim budżecie: ...'
    """
    if not kategoria:
        return ""
    if "laptop" not in kategoria.lower():
        return ""

    url = _budget_url(price)
    if not url:
        return ""

    return (
        f"<p>Sprawdź też inne niezawodne laptopy w Twoim budżecie: {url}. "
        f"Każdy komputer jest dokładnie sprawdzany, czyszczony i konfigurowany, aby zapewnić niezawodność w codziennym użytkowaniu. "
        f"Kupując sprzęt, zyskujesz jakość klasy biznes oraz pewność gwarancji door-to-door.</p>"
    )

def _build_footer_html(name, kategoria, attrs, price):
    link_block = _build_link_block(kategoria, attrs, price)
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

def _replace_used_to_refurb(text: str) -> str:
    """używany/używane -> Odnowiony/Odnowione w treści HTML."""
    if not text:
        return text
    text = re.sub(r"(?i)\bużywany\b", "Odnowiony", text)
    text = re.sub(r"(?i)\bużywane\b", "Odnowione", text)
    return text

def _append_footer_to_desc(o_el):
    """Dopisuje stopkę do <desc> (HTML) i zapisuje w CDATA, z podmianą używany -> odnowiony."""
    desc_el = o_el.find("desc")
    if desc_el is None:
        return

    current_html = _inner_html(desc_el)
    current_html = _replace_used_to_refurb(current_html)

    if _already_has_footer(current_html):
        _set_desc_cdata(desc_el, current_html)
        return

    attrs = _collect_attrs(o_el)
    name = _name(o_el)
    kategoria = _category(o_el)
    price = _get_price(o_el)

    footer_html = _build_footer_html(name, kategoria, attrs, price)
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

    attrs = _collect_attrs(o_el)
    name = _name(o_el)
    kategoria = _category(o_el)
    price = _get_price(o_el)
    footer_html = _build_footer_html(name, kategoria, attrs, price)

    def _append_text_block(sections_list):
        if not sections_list:
            sections_list.append({"items": [{"type": "TEXT", "content": footer_html}]})
            return
        sections_list.append({"items": [{"type": "TEXT", "content": footer_html}]})

    if isinstance(data, dict):
        sections = data.get("sections")
        if isinstance(sections, list):
            _append_text_block(sections)
        else:
            data["sections"] = [{"items": [{"type": "TEXT", "content": footer_html}]}]
    elif isinstance(data, list):
        _append_text_block(data)
    else:
        return

    dj.text = json.dumps(data, ensure_ascii=False)

# --------- GŁÓWNA LOGIKA ---------
def convert_file_swop(in_path, out_path):
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

        # --- KATEGORIA: poleasingowe -> odnowione + dopisywanie odnowione ---
        cat_el = o.find("cat")
        if cat_el is not None and cat_el.text:
            text = cat_el.text.strip()
            # zamień każde 'poleasingowe' na 'odnowione'
            text = re.sub(r"(?i)poleasingowe", "odnowione", text)
            norm = text.lower()
            if "odnowione" not in norm:
                if norm == "laptopy":
                    text = "Laptopy odnowione"
                elif norm == "komputery":
                    text = "Komputery odnowione"
                elif norm == "monitory komputerowe":
                    text = "Monitory odnowione"
            cat_el.text = text

        # NIE usuwamy desc_json – zostaje w XML
        attrs_el = o.find("attrs")
        if attrs_el is not None:
            attrs = _collect_attrs(o)

            # Dodaj <a name="Marka"> jeśli brak, z wartością z Producent
            producent = attrs.get("Producent", "").strip()
            if producent and not any((a.get("name") or "").strip() == "Marka" for a in attrs_el.findall("a")):
                ET.SubElement(attrs_el, "a", {"name": "Marka"}).text = producent

            # Zamiana stanu: Używany/Używane -> Odnowiony/Odnowione
            for a in attrs_el.findall("a"):
                if (a.get("name") or "").strip().lower() == "stan":
                    val = (a.text or "").strip()
                    if re.search(r"(?i)\bużywany\b", val):
                        a.text = "Odnowiony"
                    elif re.search(r"(?i)\bużywane\b", val):
                        a.text = "Odnowione"

            # Gwarancja (jak wcześniej)
            for a in attrs_el.findall("a"):
                if (a.get("name") or "").strip().lower() == "informacje o gwarancjach":
                    text = (a.text or "").strip()
                    m = re.search(r"(\d+)", text)
                    value = m.group(1) if m else ""
                    a.set("name", "Gwarancja")
                    a.text = value

        # Dopnij stopkę do HTML (z podmianą używany -> odnowiony)
        _append_footer_to_desc(o)
        # Dopnij stopkę również do JSON-a
        _append_footer_to_desc_json(o)

    tree.write(out_path, encoding="utf-8", xml_declaration=True, pretty_print=True)
    try:
        os.remove(temp_path)
    except FileNotFoundError:
        pass

    print(f"[swop OK] Zapisano: {out_path}")

def main():
    for name in os.listdir(INPUT_DIR):
        if name.lower().endswith((".xlsm", ".xlsx", ".xls")):
            src = os.path.join(INPUT_DIR, name)
            dst = os.path.join(OUTPUT_DIR, "swop.xml")
            print(f"[swop] {src} -> {dst}")
            convert_file_swop(src, dst)
            break

if __name__ == "__main__":
    main()
