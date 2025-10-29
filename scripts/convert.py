#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
from pathlib import Path
from typing import List, Tuple, Dict, Any

import xml.etree.ElementTree as ET
from openpyxl import load_workbook


INPUT_DIR = Path("input")
OUTPUT_DIR = Path("output")
OUTPUT_FILE = OUTPUT_DIR / "1.xml"

HEADER_ROW = 4       # wiersz z nagłówkami
FIRST_DATA_ROW = 5   # dane od tego wiersza w dół

REQUIRED_COLS = [
    "ID oferty",
    "Tytuł oferty",
    "Cena PL",
    "Kategoria główna",
]

# Mapa kolumn Excela -> nazwy atrybutów <a name="..."> w XML
# (puste wartości są pomijane automatycznie)
ATTR_MAP: Dict[str, str] = {
    # Identyfikacja
    "Producent": "Producent",
    "Kod producenta": "kod_producenta",
    "Model": "model_1",
    "ID produktu (EAN/UPC/ISBN/ISSN/ID produktu Allegro)": "ean",
    "Sygnatura/SKU Sprzedającego": "sku",

    # CPU / RAM
    "Model procesora": "model_procesora",
    "Generacja procesora": "generacja_procesora",
    "Taktowanie bazowe procesora [GHz]": "taktowanie_bazowe_procesora",
    "Taktowanie maksymalne procesora [GHz]": "taktowanie_maksymalne_procesora",
    "Liczba rdzeni procesora": "liczba_rdzeni_procesora",
    "Liczba wątków procesora": "liczba_watkow_procesora",
    "Pamięć podręczna procesora [MB]": "pamiec_podreczna_procesora_mb",
    "Typ pamięci RAM": "typ_pamieci_ram",
    "Wielkość pamięci RAM": "wielkosc_pamieci_ram",
    "Taktowanie szyny pamięci RAM [MHz]": "taktowanie_pamieci_ram_mhz",
    "Maksymalna pojemność pamięci RAM [GB]": "max_pamiec_ram_gb",

    # Dysk / magazyn
    "Typ dysku twardego": "typ_dysku",
    "Pojemność dysku [GB]": "pojemnosc_dysku_gb",
    "Format dysku": "format_dysku",
    "Interfejs dysku": "interfejs_dysku",
    "Prędkość obrotowa dysku HDD": "predkosc_hdd_rpm",

    # Grafika
    "Producent karty graficznej": "producent_karty_graficznej",
    "Chipset karty graficznej": "chipset_karty_graficznej",
    "Pamięć karty graficznej": "pamiec_karty_graficznej",
    "Złącza karty graficznej": "zlacza_karty_graficznej",
    "Rodzaj karty graficznej": "rodzaj_karty_graficznej",

    # Ekran
    "Przekątna ekranu (cale) [\"]": "przekatna_ekranu",
    "Rozdzielczość natywna [px]": "rozdzielczosc_ekranu",
    "Typ matrycy": "typ_matrycy",
    "Powłoka matrycy": "powloka_matrycy",
    "Proporcje obrazu": "proporcje_obrazu",
    "Częstotliwość odświeżania [Hz]": "odswiezanie_hz",

    # Wejścia/wyjścia i łączność
    "Złącza": "zlacza_zewnetrzne",
    "Komunikacja": "komunikacja",
    "Standard HDMI": "standard_hdmi",

    # System / klawiatura / zasilanie
    "System operacyjny": "system_operacyjny",
    "Wersja systemu operacyjnego": "wersja_systemu_operacyjnego",
    "Typ klawiatury": "typ_klawiatury",
    "Układ klawiatury": "uklad_klawiatury",
    "Ładowarka w komplecie": "ladowarka_w_komplecie",

    # Stan / gwarancja
    "Stan": "stan_urządzenia",
    "Stan opakowania": "stan_opakowania",
    "Dołączone oprogramowanie": "dolaczone_oprogramowanie",
    "Informacje o gwarancjach (opcjonalne)": "gwarancja_info",
}


def _first_input_file() -> Path:
    # bierzemy pierwszy .xlsm/.xlsx z input/
    for ext in ("*.xlsm", "*.xlsx"):
        files = sorted(INPUT_DIR.glob(ext))
        if files:
            return files[0]
    print("[ERR] Brak pliku XLSX/XLSM w katalogu input/", file=sys.stderr)
    sys.exit(1)


def _as_str(v: Any) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return s


def _as_float(v: Any) -> float:
    try:
        return float(str(v).replace(",", "."))
    except Exception:
        return 0.0


def _as_int(v: Any) -> int:
    try:
        return int(float(str(v).replace(",", ".")))
    except Exception:
        return 0


def _read_headers(ws) -> List[str]:
    headers: List[str] = []
    max_col = ws.max_column
    for c in range(1, max_col + 1):
        headers.append(_as_str(ws.cell(row=HEADER_ROW, column=c).value))
    return headers


def _idx(headers: List[str], name: str) -> int:
    """Zwraca indeks kolumny po dokładnym nagłówku, -1 jeśli brak."""
    try:
        return headers.index(name)
    except ValueError:
        return -1


def main() -> None:
    INPUT_DIR.mkdir(exist_ok=True, parents=True)
    OUTPUT_DIR.mkdir(exist_ok=True, parents=True)

    src = _first_input_file()
    print(f"[RUN] {src.name} -> {OUTPUT_FILE}")

    # Wczytanie arkusza – bierzemy 'Szablon' jeśli jest, inaczej pierwszy
    wb = load_workbook(src, data_only=False, read_only=True)
    ws = wb["Szablon"] if "Szablon" in wb.sheetnames else wb.worksheets[0]

    headers = _read_headers(ws)
    print(f"[INFO] Arkusz: {ws.title}")
    print(f"[DEBUG] Liczba kolumn: {len(headers)}")
    print(f"[DEBUG] Pierwsze 20 nagłówków: {headers[:20]}")

    # Walidacja wymaganych kolumn
    missing = [h for h in REQUIRED_COLS if h not in headers]
    if missing:
        print(f"[ERROR] Brak wymaganych nagłówków: {missing}", file=sys.stderr)
        sys.exit(1)

    # Indeksy często używanych kolumn
    idx_id = _idx(headers, "ID oferty")
    idx_title = _idx(headers, "Tytuł oferty")
    idx_price = _idx(headers, "Cena PL")
    idx_cat = _idx(headers, "Kategoria główna")
    idx_link = _idx(headers, "Link do oferty")
    idx_imgs = _idx(headers, "Zdjęcia")
    idx_desc = _idx(headers, "Opis oferty")
    idx_stock_qty = _idx(headers, "Liczba sztuk")
    idx_status = _idx(headers, "Status oferty")

    root = ET.Element("offers")

    created = 0
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, len(headers) + 1)]
        # Pusta linia? – pomijamy
        if not any(row_vals):
            continue

        # ID oferty – obowiązkowe
        offer_id = _as_str(row_vals[idx_id]) if idx_id != -1 else ""
        if not offer_id:
            continue  # bez ID nie generujemy

        # Pola bazowe
        title = _as_str(row_vals[idx_title]) if idx_title != -1 else ""
        price = _as_float(row_vals[idx_price]) if idx_price != -1 else 0.0
        url = _as_str(row_vals[idx_link]) if idx_link != -1 else ""
        cat = _as_str(row_vals[idx_cat]) if idx_cat != -1 else ""
        desc_html = _as_str(row_vals[idx_desc]) if idx_desc != -1 else ""

        # Dostępność: Status oferty + liczba sztuk
        status_txt = _as_str(row_vals[idx_status]) if idx_status != -1 else ""
        qty = _as_int(row_vals[idx_stock_qty]) if idx_stock_qty != -1 else 0

        if status_txt.lower() == "aktywna" and qty > 0:
            avail = "1"
            basket = "1"
        else:
            avail = "99"
            basket = "0"

        # element <o ...>
        o = ET.SubElement(root, "o", {
            "id": offer_id,
            "url": url,
            "price": f"{price:.2f}",
            "avail": avail,
            "stock": "1",
            "basket": basket,
        })

        if cat:
            cat_el = ET.SubElement(o, "cat")
            cat_el.text = cat  # tylko Kategoria główna

        if title:
            name_el = ET.SubElement(o, "name")
            name_el.text = title

        # <desc> – HTML z "Opis oferty"
        if desc_html:
            desc_el = ET.SubElement(o, "desc")
            desc_el.text = desc_html  # jeżeli chcesz CDATA – daj znać

        # <imgs> – rozdzielenie po '|'
        if idx_imgs != -1:
            imgs_raw = _as_str(row_vals[idx_imgs])
            if imgs_raw:
                parts = [p.strip() for p in imgs_raw.split("|") if p.strip()]
                if parts:
                    imgs_el = ET.SubElement(o, "imgs")
                    # pierwszy jako <main>, reszta jako <i>
                    ET.SubElement(imgs_el, "main", {"url": parts[0]})
                    for p in parts[1:]:
                        ET.SubElement(imgs_el, "i", {"url": p})

        # <attrs> – tylko wypełnione parametry z mapy
        attrs_written = 0
        attrs_el = ET.SubElement(o, "attrs")
        for xl_col, attr_name in ATTR_MAP.items():
            i = _idx(headers, xl_col)
            if i == -1:
                continue
            val = _as_str(row_vals[i])
            if not val:
                continue
            ET.SubElement(attrs_el, "a", {"name": attr_name}).text = val
            attrs_written += 1

        created += 1

    # zapis XML
    tree = ET.ElementTree(root)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    tree.write(OUTPUT_FILE, encoding="utf-8", xml_declaration=True)

    print(f"[OK] Zapisano: {OUTPUT_FILE} | ofert: {created}")


if __name__ == "__main__":
    main()
