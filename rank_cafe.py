"""Ranking cafe prospek dari yang paling menjanjikan sampai tidak.

Skor:
  +60  WA valid (bisa dihubungi langsung untuk outreach)
  +25  punya Instagram aktif (bisnis hidup & peduli marketing)
  +15  bonus kalau WA + IG keduanya ada (prospek paling layak dikejar)
  -40  nomor landline (kode area 03xx) di kolom WA -- bukan WA asli
  -999 franchise jaringan besar (Tomoro dll) -- keputusan website di
       kantor pusat, bukan pemilik lokal; buang ke dasar list
  -20  tanpa kontak sama sekali -- tidak bisa di-outreach

Pemakaian:
    python rank_cafe.py hasil_kota.xlsx
Output: urutan_prospek.md + ringkasan ke terminal.
"""

import re
import sys

import openpyxl

KOL_NAMA, KOL_TELP, KOL_WA, KOL_IG, KOL_CATATAN = 2, 4, 5, 10, 17

# Kafe jaringan/franchise nasional -- website diurus kantor pusat.
FRANCHISE = ("tomoro", "janji jiwa", "kopi kenangan", "fore coffee",
             "point coffee", "starbucks", "excelso", "kopi soe")


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


def skor_cafe(nama: str, wa: str | None, ig: str,
              catatan: str) -> tuple[int, str]:
    if any(f in nama.lower() for f in FRANCHISE):
        return -999, "Franchise -- keputusan website di kantor pusat"

    skor, alasan = 0, []
    ig_aktif = bool(ig) and "tidak ditemukan" not in catatan.lower()

    if wa:
        if re.match(r"^62?3\d{2}", wa):  # 0331 dst = landline, bukan WA
            skor -= 40
            alasan.append("nomor landline, bukan WA")
            wa = None
        else:
            skor += 60
            alasan.append("WA valid")
    if ig_aktif:
        skor += 25
        alasan.append("IG aktif")
    if wa and ig_aktif:
        skor += 15
        alasan.append("kontak ganda (WA+IG)")
    if not wa and not ig_aktif:
        skor -= 20
        alasan.append("tanpa kontak yang bisa dihubungi")

    return skor, ", ".join(alasan)


def main() -> None:
    path = sys.argv[1] if len(sys.argv) > 1 else "hasil_kota.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Prospek Cafe"]

    baris = []
    for r in range(4, ws.max_row + 1):
        nama = ws.cell(r, KOL_NAMA).value
        if not nama:
            continue
        nama = str(nama).strip()
        wa = ambil_wa(ws.cell(r, KOL_WA))
        ig = str(ws.cell(r, KOL_IG).value or "").strip()
        ig = "" if ig == "-" else ig
        catatan = str(ws.cell(r, KOL_CATATAN).value or "")
        skor, alasan = skor_cafe(nama, wa, ig, catatan)
        baris.append({"row": r, "nama": nama, "skor": skor,
                      "wa": bool(wa and not re.match(r"^62?3\d{2}", wa)),
                      "ig": bool(ig and "tidak ditemukan"
                                 not in catatan.lower()),
                      "alasan": alasan})

    baris.sort(key=lambda b: -b["skor"])

    garis = ["# Urutan Prospek Cafe (paling menjanjikan -> paling tidak)",
             "",
             "| # | Cafe | Skor | WA | IG | Alasan |",
             "|--:|------|-----:|:--:|:--:|--------|"]
    for i, b in enumerate(baris, 1):
        garis.append(f"| {i} | {b['nama']} | {b['skor']} | "
                     f"{'v' if b['wa'] else '-'} | {'v' if b['ig'] else '-'} | "
                     f"{b['alasan']} |")
    with open("urutan_prospek.md", "w", encoding="utf-8") as f:
        f.write("\n".join(garis) + "\n")

    print(f"[OK] {len(baris)} cafe diranking -> urutan_prospek.md\n")
    print("TOP 15:")
    for i, b in enumerate(baris[:15], 1):
        print(f"  {i:>2}. [{b['skor']:>4}] {b['nama']:<45} {b['alasan']}")


if __name__ == "__main__":
    main()
