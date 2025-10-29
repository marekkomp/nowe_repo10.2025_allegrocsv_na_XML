import os
from convert import convert_file, INPUT_DIR, OUTPUT_DIR

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    for name in os.listdir(INPUT_DIR):
        if name.lower().endswith((".xlsx", ".xlsm", ".xls")):
            src = os.path.join(INPUT_DIR, name)
            dst = os.path.join(OUTPUT_DIR, "klientB.xml")
            print(f"[KlientB] {src} -> {dst}")
            convert_file(src, dst)
            break

if __name__ == "__main__":
    main()

