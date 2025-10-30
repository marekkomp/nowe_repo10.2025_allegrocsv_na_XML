# scripts/convert_Morele.py
import os
import xml.etree.ElementTree as ET
from convert import convert_file, INPUT_DIR, OUTPUT_DIR  # główny konwerter

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
    
    # 3) zapis i sprzątanie
    ET.indent(root, space="  ")
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
