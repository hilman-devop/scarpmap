"""Buat daftar blast WhatsApp (blast_wa.html) dari file Excel hasil ScrapMap.

Setiap baris punya tombol yang membuka WhatsApp dengan pesan penawaran
SUDAH TERISI otomatis (tinggal tekan Kirim). Status terkirim dicatat di
localStorage browser sehingga tidak ada yang terkirim dua kali.

Pemakaian:
    python buat_blast_wa.py hasil_jember.xlsx
"""

import html
import re
import sys
from urllib.parse import quote

import openpyxl

# ---------------------------------------------------------------------------
# WAJIB DIEDIT SEBELUM DIPAKAI
# ---------------------------------------------------------------------------
NAMA_ANDA = "Hilman Arifin Ilham"
KOTA_ANDA = "Jember"

TEMPLATE = """Halo Kak, selamat siang 🙏

Perkenalkan, saya {nama_anda} dari {kota_anda}. Saya menemukan
*{nama_cafe}* di Google Maps dan salut dengan cafenya!

Sekalian izin menawarkan: saya perhatikan {nama_cafe} belum punya
website resmi. Padahal banyak calon pelanggan yang mencari info menu,
jam buka, dan suasana cafe lewat Google sebelum datang.

Saya bisa bantu buatkan website sederhana khusus cafe, isinya:
• Daftar menu & foto-foto cafe
• Lokasi + tombol arah langsung ke Google Maps
• Tombol "Pesan via WhatsApp" (langsung ke nomor Kakak)
• Jam buka & info promo

Boleh saya kirimkan contoh websitenya dulu? Gratis, tanpa komitmen
apa pun — kalau cocok baru lanjut ngobrol. 🙏

Terima kasih, Kak. Ditunggu kabarnya ☕"""

# Kolom Excel hasil ScrapMap (1-based)
KOL_NAMA, KOL_TELP, KOL_WA = 2, 4, 5
# prefix nomor seluler Indonesia (setelah 62)
PREFIX_SELULER = ("812", "813", "821", "822", "823", "851", "852", "853",
                  "855", "856", "857", "858", "895", "896", "897", "898",
                  "899", "817", "818", "819", "859", "877", "878", "831",
                  "832", "833", "838", "881", "882", "883", "884", "885",
                  "886", "887", "888", "889")


def ambil_nomor_wa(ws, baris: int) -> str | None:
    """Ambil nomor WA (format 62xxx) dari sel; None jika bukan seluler."""
    sel = ws.cell(baris, KOL_WA)
    mentah = None
    if sel.hyperlink and "wa.me/" in sel.hyperlink.target:
        mentah = sel.hyperlink.target.split("wa.me/")[1].split("?")[0]
    elif sel.value:
        mentah = str(sel.value)
    if not mentah:
        return None
    digit = re.sub(r"\D", "", mentah)
    if digit.startswith("0"):
        digit = "62" + digit[1:]
    if not digit.startswith("62"):
        return None
    sisa = digit[2:]
    if not any(sisa.startswith(p) for p in PREFIX_SELULER):
        return None  # telepon rumah, tidak bisa di-WA
    return digit


def main() -> None:
    if len(sys.argv) < 2:
        raise SystemExit("Pemakaian: python buat_blast_wa.py <file_hasil.xlsx>")
    path = sys.argv[1]
    wb = openpyxl.load_workbook(path)
    ws = wb["Prospek Cafe"]

    entri = []
    for r in range(4, ws.max_row + 1):
        nama = ws.cell(r, KOL_NAMA).value
        if not nama:
            continue
        nomor = ambil_nomor_wa(ws, r)
        if nomor:
            entri.append((str(nama), nomor))

    if not entri:
        raise SystemExit("[WARN] Tidak ada nomor WA seluler di file ini.")

    if NAMA_ANDA == "Nama Kamu":
        print("[WARN] NAMA_ANDA/KOTA_ANDA belum diedit di buat_blast_wa.py!")

    baris_html = []
    for i, (nama, nomor) in enumerate(entri, start=1):
        pesan = TEMPLATE.format(nama_cafe=nama, nama_anda=NAMA_ANDA,
                                kota_anda=KOTA_ANDA)
        url = f"https://wa.me/{nomor}?text={quote(pesan)}"
        baris_html.append(
            f'<tr id="row{i}"><td>{i}</td><td>{html.escape(nama)}</td>'
            f"<td>+{nomor}</td>"
            f'<td><a class="btn" href="{url}" target="_blank" '
            f'onclick="tandai({i})">Kirim WA</a></td>'
            f'<td><input type="checkbox" id="chk{i}" '
            f'onchange="simpan({i})"></td></tr>'
        )

    halaman = HALAMAN.replace("__JUDUL__", path).replace(
        "__JUMLAH__", str(len(entri))).replace("__BARIS__", "\n".join(baris_html))
    with open("blast_wa.html", "w", encoding="utf-8") as f:
        f.write(halaman)
    print(f"[OK] {len(entri)} nomor WA -> blast_wa.html (buka di browser)")


HALAMAN = """<!DOCTYPE html>
<html lang="id"><head><meta charset="utf-8">
<title>Blast WA - __JUDUL__</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;max-width:900px;margin:24px auto;padding:0 16px}
table{border-collapse:collapse;width:100%}
th,td{border:1px solid #ddd;padding:8px;text-align:left}
th{background:#1F4E79;color:#fff}
.btn{background:#25D366;color:#fff;padding:6px 14px;border-radius:6px;
text-decoration:none;font-weight:bold}
tr.terkirim{background:#d4edda}
.info{background:#fff3cd;padding:10px;border-radius:6px;margin-bottom:12px}
</style></head><body>
<h1>Daftar Blast WhatsApp</h1>
<div class="info">Sumber: <b>__JUDUL__</b> &mdash; <b>__JUMLAH__</b> nomor.
Klik "Kirim WA" satu per satu (pesan sudah terisi otomatis, tinggal tekan
Kirim di WhatsApp). Centang = sudah terkirim (tersimpan di browser ini).
<b>Jangan kirim lebih dari 5-10 per jam agar nomor tidak di-banned.</b></div>
<table><tr><th>No</th><th>Nama Cafe</th><th>Nomor</th><th>Aksi</th><th>Terkirim</th></tr>
__BARIS__
</table>
<script>
function simpan(i){var c=document.getElementById('chk'+i);
localStorage.setItem('wa'+i,c.checked?'1':'');
document.getElementById('row'+i).className=c.checked?'terkirim':'';}
function tandai(i){var c=document.getElementById('chk'+i);
c.checked=true;simpan(i);}
document.querySelectorAll('input[type=checkbox]').forEach(function(c){
var i=c.id.replace('chk','');
if(localStorage.getItem('wa'+i)){c.checked=true;
document.getElementById('row'+i).className='terkirim';}});
</script></body></html>
"""


if __name__ == "__main__":
    main()
