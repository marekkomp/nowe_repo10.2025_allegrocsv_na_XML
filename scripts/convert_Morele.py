# scripts/convert_Morele.py
import os
import xml.etree.ElementTree as ET
from convert import convert_file, INPUT_DIR, OUTPUT_DIR  # główny konwerter

# --------- USTAWIENIA ---------
BRAND_LINKS = {
    "dell":   "https://kompre.pl/pl/c/Laptopy-Dell/364",
    "lenovo": "https://kompre.pl/pl/c/Laptopy-Lenovo/366",
    "hp":     "https://kompre.pl/pl/c/Laptopy-HP/365",
    "apple":  "https://kompre.pl/pl/c/Laptopy-Apple/367",
    "fujitsu":"https://kompre.pl/pl/c/Laptopy-Fujitsu/368",
}
FOOTER_MARK = "<!---->"  # znacznik, by nie dublować

def _collect_attrs(o_el):
    """Zwraca dict z <attrs><a name="...">wartość</a>."""
    out = {}
    attrs_el = o_el.find("attrs")
    if attrs_el is None:
        return out
    for a in attrs_el.findall("a"):
        name = a.get("name") or ""
        val = (a.text or "").strip()
        if name:
            out[name.strip()] = val
    return out

def _brand(o_el, attrs):
    # priorytet: <a name="Producent">, fallback pusty
    return (attrs.get("Producent") or "").strip()

def _warranty(attrs):
    # priorytet: Informacje o gwarancjach > Gwarancja > puste
    gw = (attrs.get("Informacje o gwarancjach")
          or attrs.get("Gwarancja")
          or "").strip()
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
    # jeśli marketplace nie pozwala na linki, zmień na zwykły tekst
    return (f'<p>Posiadamy też inne modele - sprawdź '
        f'<a href="{url}" rel="nofollow noopener" target="_blank">polecane laptopy {producent.lower()}</a>.</p>')


def _build_footer_html(name, producent, gwarancja, kategoria):
    link_block = _build_link_block(kategoria, producent)
    gw = gwarancja or "12 miesięcy"
    # zwięzły, E-E-A-T + marker
    footer = (
        f'{FOOTER_MARK}'
        f'<hr><p><strong>{name}</strong> pochodzi z oferty <strong>Kompre.pl</strong> – '
        f'autoryzowanego sprzedawcy komputerów poleasingowych klasy biznes. '
        f'Każdy egzemplarz jest testowany, czyszczony i przygotowany do pracy z aktualnym systemem. '
        f'{gw if "gwarancja" in gw.lower() else "Gwarancja " + gw} zapewnia wsparcie door-to-door i bezpieczeństwo zakupu.</p>'

        f'{link_block}'
    )
    return footer

def _append_footer_to_desc(o_el):
    desc_el = o_el.find("desc")
    if desc_el is None:
        return  # brak opisu – nic nie robimy
    current = desc_el.text or ""
    if FOOTER_MARK in current:
        return  # już dodane

    attrs = _collect_attrs(o_el)
    producent = _brand(o_el, attrs)
    gwarancja = _warranty(attrs)
    kategoria = _category(o_el)
    name = _name(o_el)

    footer_html = _build_footer_html(name, producent, gwarancja, kategoria)

    # dopinamy na końcu aktualnego HTML-a w <desc>
    # (bez &nbsp;, tylko zwykłe spacje/newline)
    joiner = "\n" if current and not current.endswith("\n") else ""
    desc_el.text = f"{current}{joiner}{footer_html}"

def convert_file_morele(in_path, out_path):
    # 1) pełny XML tymczasowo
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    temp_path = os.path.join(OUTPUT_DIR, "_temp_base.xml")
    convert_file(in_path, temp_path)

    # 2) modyfikacje dla Morele
    tree = ET.parse(temp_path)
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

        # dodaj "poleasingowe" do kategorii Laptopy i Komputery
        cat_el = o.find("cat")
        if cat_el is not None and cat_el.text:
            cat_text = cat_el.text.strip()
            if cat_text.lower() in ("laptopy", "komputery"):
                cat_el.text = f"{cat_text} poleasingowe"

        # usuń desc_json
        for dj in o.findall("desc_json"):
            o.remove(dj)

        # >>> dopnij nasz footer na końcu opisu
        _append_footer_to_desc(o)

    # 3) zapis i sprzątanie
    try:
        ET.indent(root, space="  ")
    except AttributeError:
        # Python <3.9: brak indent – pomijamy
        pass

    tree.write(out_path, encoding="utf-8", xml_declaration=True)

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
