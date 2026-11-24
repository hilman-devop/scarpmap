"""Generate draft cafe demo websites from a local prospect spreadsheet.

Generated sites contain business names, contacts, and map links. Keep them
local and publish only with owner approval.

Pemakaian:
    python buat_demo.py hasil_kota.xlsx            # 5 demo teratas
    python buat_demo.py hasil_kota.xlsx --jumlah 10

Upload isi tiap folder ke Netlify Drop (https://app.netlify.com/drop)
untuk mendapatkan link demo gratis dalam hitungan detik.
"""

import re
import sys
from pathlib import Path

import openpyxl

# ---------------------------------------------------------------------------
# WAJIB DIEDIT SEBELUM DIPAKAI
# ---------------------------------------------------------------------------
NAMA_ANDA = "Nama Anda"
WA_ANDA = "Nomor WhatsApp Anda"  # ditampilkan di footer demo

KOL_NAMA, KOL_ALAMAT, KOL_WA, KOL_PETA = 2, 3, 5, 16

MENU_STANDAR = [
    ("Kopi Susu Gula Aren", "Rp 15.000"),
    ("Es Kopi Klepon", "Rp 18.000"),
    ("Cafe Latte", "Rp 20.000"),
    ("Matcha Latte", "Rp 22.000"),
    ("Kentang Goreng", "Rp 15.000"),
    ("Nasi Goreng Spesial", "Rp 25.000"),
]


def slugify(teks: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", teks.lower()).strip("-")[:40] or "cafe"


def ambil_wa(sel) -> str | None:
    mentah = None
    if sel.hyperlink and "wa.me/" in (sel.hyperlink.target or ""):
        mentah = sel.hyperlink.target.split("wa.me/")[1].split("?")[0]
    elif sel.value:
        mentah = str(sel.value)
    if not mentah:
        return None
    digit = re.sub(r"\D", "", mentah)
    if digit.startswith("0"):
        digit = "62" + digit[1:]
    return digit if digit.startswith("62") else None


def main() -> None:
    if len(sys.argv) < 2:
        raise SystemExit("Pemakaian: python buat_demo.py <hasil.xlsx> "
                         "[--jumlah N]")
    path = sys.argv[1]
    jumlah = 5
    if "--jumlah" in sys.argv:
        jumlah = int(sys.argv[sys.argv.index("--jumlah") + 1])

    kota = re.sub(r"^hasil_|\.xlsx$", "", Path(path).stem).title() or "Kota"
    template = Path("demo_template.html").read_text(encoding="utf-8")

    wb = openpyxl.load_workbook(path)
    ws = wb["Prospek Cafe"]

    dibuat = 0
    for r in range(4, ws.max_row + 1):
        if dibuat >= jumlah:
            break
        nama = ws.cell(r, KOL_NAMA).value
        if not nama:
            continue
        nama = str(nama)
        alamat = str(ws.cell(r, KOL_ALAMAT).value or kota)
        wa = ambil_wa(ws.cell(r, KOL_WA))
        peta_sel = ws.cell(r, KOL_PETA)
        maps_link = (peta_sel.hyperlink.target if peta_sel.hyperlink
                     else peta_sel.value) or \
            "https://www.google.com/maps/search/" + re.sub(r"\s+", "+", nama)

        wa_link = (f"https://wa.me/{wa}?text=Halo%20{nama.replace(' ', '%20')}"
                   "%2C%20saya%20mau%20pesan%20%2F%20reservasi%20tempat%20"
                   if wa else "#")

        menu_html = "\n".join(
            f'    <div class="menu-item"><span>{m}</span>'
            f'<span class="harga">{h}</span></div>'
            for m, h in MENU_STANDAR)

        seed = slugify(nama)
        halaman = (template
                   .replace("{{NAMA}}", nama)
                   .replace("{{KOTA}}", kota)
                   .replace("{{TAGLINE}}",
                            f"Tempat ngopi & bersantai favorit di {kota}")
                   .replace("{{FOTO_HERO}}",
                            f"https://picsum.photos/seed/{seed}-hero/1600/900")
                   .replace("{{FOTO1}}",
                            f"https://picsum.photos/seed/{seed}-1/600/400")
                   .replace("{{FOTO2}}",
                            f"https://picsum.photos/seed/{seed}-2/600/400")
                   .replace("{{FOTO3}}",
                            f"https://picsum.photos/seed/{seed}-3/600/400")
                   .replace("{{MENU_ITEM}}", menu_html)
                   .replace("{{ALAMAT}}", alamat)
                   .replace("{{MAPS_LINK}}", str(maps_link))
                   .replace("{{WA_LINK}}", wa_link)
                   .replace("{{NAMA_ANDA}}", NAMA_ANDA)
                   .replace("{{WA_ANDA}}", WA_ANDA))

        folder = Path("demo_sites") / seed
        folder.mkdir(parents=True, exist_ok=True)
        (folder / "index.html").write_text(halaman, encoding="utf-8")
        dibuat += 1
        print(f"[OK] {nama} -> demo_sites/{seed}/index.html")

    if NAMA_ANDA == "Nama Kamu":
        print("[WARN] NAMA_ANDA/WA_ANDA belum diedit di buat_demo.py!")
    print(f"\n[INFO] {dibuat} demo dibuat. Upload tiap folder ke "
          "https://app.netlify.com/drop untuk link demo gratis.")
    print("[INFO] Ganti foto placeholder dengan foto asli cafe "
          "(dari Google Maps/IG mereka) agar lebih meyakinkan!")


if __name__ == "__main__":
    main()
