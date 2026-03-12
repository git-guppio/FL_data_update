from PIL import Image
from pathlib import Path

SRC  = Path(r"C:\Users\a259046\OneDrive - Enel Spa\SCRIPT AHK e VBA\GITHUB\FL_data_update\dist\icon.png")
DEST = Path(r"C:\Users\a259046\OneDrive - Enel Spa\SCRIPT AHK e VBA\GITHUB\FL_data_update\dist\my_icon.ico")

if __name__ == "__main__":
    print(f"[DEBUG] Sorgente : {SRC}")
    print(f"[DEBUG] Esiste   : {SRC.exists()}")

    if not SRC.exists():
        print(f"[ERRORE] File sorgente non trovato: {SRC}")
        raise SystemExit(1)

    print(f"[DEBUG] Apertura immagine...")
    logo = Image.open(SRC)
    print(f"[DEBUG] Immagine aperta — size={logo.size}, mode={logo.mode}")

    sizes = [(16, 16), (32, 32), (48, 48), (256, 256)]
    print(f"[DEBUG] Conversione in ICO con sizes={sizes}")
    logo.save(DEST, format='ICO', sizes=sizes)

    print(f"[DEBUG] ICO salvato in: {DEST}")
    print(f"[DEBUG] File creato  : {DEST.exists()}, dimensione={DEST.stat().st_size} byte")