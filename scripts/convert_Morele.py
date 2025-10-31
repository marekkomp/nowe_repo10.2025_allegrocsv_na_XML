# scripts/convert_Morele.py
import os
import re
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
FOOTER_MARK = "<!---->"  # znacznik, by nie dublować (umieszczany jako komentarz w HTML/CDATA)
LINKS_AS_PLAIN_TEXT = True  # linki w stopce jako zwykły tekst (Morele ukrywa <a>)

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
    if "laptop" not in kategoria.lower():
        return ""
    url = BRAND_LINKS.get(producent.lower())
    if not url:
        return ""
    if LINKS_AS_PLAIN_TEXT:
        # goły URL (bez <a>)
        return f"<p>Posiadamy też inne modele – sprawdź: {url}</p>"
    else:
        return (f'<p>Posiadamy też inne modele – sprawdź: '
                f'<a href="{url}" rel="nofollow noopener" target="_blank">laptopy {producent.lower()}</a>.</p>')

def _build_footer_html(name, producent, gwarancja, kategoria):
    link_block = _build_link_block(kategoria, producent)
    gw = (gwarancja or "12 miesięcy").strip()
    gwarancja_txt = gw if "gwarancja" in gw.lower() else f"Gwarancja {gw}"
    return (
        f'{FOOTER_MARK}'
        f'<hr/><p><strong>{name}</strong> pochodzi z oferty <strong>Kompre.pl</strong> – '
        f'autoryzowanego sprzedawcy komputerów poleasingowych klasy biznes. '
        f'Każdy egzemplarz jest testowany, czyszczony i przygotowany do pracy z aktualnym systemem. '
        f'{gwarancja_txt} zapewnia wsparcie door-to-door i bezpieczeństwo zakupu.</p>'
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
    # zapisujemy zawartość jako CDATA, by Morele widziało prawdziwy HTML
    desc_el.text = ET.CDATA(html_string)

def _already_has_footer(html: str) -> bool:
    # marker lub charakterystyczne frazy
    return (FOOTER_MARK in html) or ("Kompre.pl" in html and "door-to-door" in html)

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
    gwarancja = _warranty(attrs)
    kategoria = _category(o_el)

    footer_html = _build_footer_html(name, producent, gwarancja, kategoria)
    joiner = "\n" if current_html and not current_html.endswith("\n") else ""
    new_html = f"{current_html}{joiner}{footer_html}".strip()
    _set_desc_cdata(desc_el, new_html)

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

        # dopisz "poleasingowe" do kategorii Laptopy/Komputery/Monitory komputerowe
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

        # usuń desc_json
        for dj in o.findall("desc_json"):
            parent = dj.getparent() if hasattr(dj, "getparent") else o
            parent.remove(dj)

        # --- Zmieniamy gwarancję na prosty format ---
        attrs_el = o.find("attrs")
        if attrs_el is not None:
            for a in attrs_el.findall("a"):
                if (a.get("name") or "").strip().lower() == "informacje o gwarancjach":
                    text = (a.text or "").strip()
                    m = re.search(r"(\d+)", text)
                    value = m.group(1) if m else ""
                    a.set("name", "Gwarancja")
                    a.text = value

        # dopnij footer i zapisz <desc> jako CDATA
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
