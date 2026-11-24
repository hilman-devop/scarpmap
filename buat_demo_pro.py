"""Generate cafe demo websites from a local prospect spreadsheet and assets.

Downloaded business photos, logos, and generated demos stay local; use
owner-approved materials when publishing a demo.
"""

import re
import shutil
import sys
from pathlib import Path

import openpyxl

from rank_cafe import ambil_wa

KOL_NAMA, KOL_ALAMAT, KOL_WA, KOL_PETA, KOL_IG = 2, 3, 5, 16, 10

# ---------------------------------------------------------------------------
# WAJIB DIEDIT SEBELUM DIPAKAI
# ---------------------------------------------------------------------------
NAMA_ANDA = "Nama Anda"
WA_ANDA = "Nomor WhatsApp Anda"

TAHUN = "2026"

# ---------------------------------------------------------------------------
# MENU per jenis usaha: (kategori, [(nama, deskripsi, harga)])
# ---------------------------------------------------------------------------
MENU_KOPI = [
    ("Kopi", [
        ("Espresso", "Double shot house blend, bold & bersih", "Rp 12.000"),
        ("Kopi Susu Gula Aren", "Espresso, susu segar, gula aren asli", "Rp 15.000"),
        ("Cafe Latte", "Espresso lembut dengan susu steamed", "Rp 20.000"),
        ("Manual Brew V60", "Single origin pilihan, seduh perlahan", "Rp 22.000"),
    ]),
    ("Non-Kopi", [
        ("Matcha Latte", "Matcha premium dengan susu", "Rp 22.000"),
        ("Cokelat Klasik", "Cokelat pekat, bisa panas atau dingin", "Rp 18.000"),
        ("Lemon Tea", "Teh segar dengan perasan lemon asli", "Rp 12.000"),
    ]),
    ("Camilan", [
        ("Kentang Goreng", "Renyah di luar, lembut di dalam", "Rp 15.000"),
        ("Roti Bakar", "Pilihan cokelat, keju, atau kacang", "Rp 14.000"),
        ("Pisang Goreng Keju", "Disajikan hangat dengan taburan keju", "Rp 16.000"),
    ]),
]

MENU_EATERY = MENU_KOPI + [
    ("Makanan Berat", [
        ("Nasi Goreng Spesial", "Bumbu rempah, telur, dan ayam suwir", "Rp 25.000"),
        ("Ayam Geprek", "Ayam crispy dengan sambal bawang", "Rp 20.000"),
        ("Mie Goreng Jawa", "Dimasak di atas api besar, gurih smoky", "Rp 22.000"),
        ("Chicken Katsu Rice", "Katsu renyah dengan saus spesial", "Rp 28.000"),
    ]),
]

MENU_GELATO = [
    ("Gelato", [
        ("Gelato 1 Scoop", "Rasa harian: cokelat, vanilla, stroberi", "Rp 15.000"),
        ("Gelato 2 Scoop", "Kombinasi dua rasa favoritmu", "Rp 25.000"),
        ("Affogato", "Gelato vanilla disiram espresso panas", "Rp 25.000"),
    ]),
    ("Kopi & Minuman", [
        ("Kopi Susu Gula Aren", "Manis gurih, favorit pelanggan", "Rp 15.000"),
        ("Cafe Latte", "Espresso dengan susu steamed lembut", "Rp 20.000"),
        ("Butterfly Pea Lemonade", "Segar, berubah warna saat diaduk", "Rp 18.000"),
    ]),
    ("Dessert", [
        ("Croffle Es Krim", "Croissant-waffle hangat + gelato", "Rp 22.000"),
        ("Banana Split", "Klasik, tiga scoop dengan topping", "Rp 30.000"),
    ]),
]

# ---------------------------------------------------------------------------
# TEMA per cafe -- dikurasi manual dari foto asli & identitas visualnya.
# bentuk_logo: circle | square | sharp | arch | minimal
# ---------------------------------------------------------------------------
def font_link(*families: str) -> str:
    q = "&".join(families)
    return (f'<link rel="preconnect" href="https://fonts.googleapis.com">'
            f'<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>'
            f'<link href="https://fonts.googleapis.com/css2?{q}&display=swap" '
            f'rel="stylesheet">')


TEMA = {
    "kafe-contoh": {
        "fonts": ("Georgia, serif", "Arial, sans-serif"),
        "font_link": "",
        "bg": "#faf7f0", "surface": "#efe8db", "text": "#2e241c",
        "muted": "#71665b", "accent": "#795548", "accent_text": "#ffffff",
        "hero_veil": "rgba(20,14,10,.78)", "hero_text": "#fffaf2",
        "radius": "8px", "radius_lg": "16px",
        "tt_display": "none", "ls_display": "0", "fw_display": "600",
        "tt_brand": "none", "ls_brand": "0", "bentuk_logo": "circle",
        "tagline": "Ruang untuk menikmati kopi dan berbagi cerita.",
        "judul_tentang": "Tentang Kafe Contoh",
        "deskripsi": "Contoh tampilan. Ganti dengan informasi yang sudah diverifikasi oleh pemilik usaha.",
        "jam": "Jam operasional diisi pemilik usaha",
        "menu": [("Menu Contoh", [("Minuman Contoh", "Nama, deskripsi, dan harga diisi pemilik usaha", "Harga")])],
    },
}
def slugify(teks: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", teks.lower()).strip("-")[:40] or "cafe"


def inisial(nama: str) -> str:
    kata = re.findall(r"[A-Za-z0-9]+", nama)
    return "".join(k[0] for k in kata[:2]).upper()


def buat_logo_svg(nama: str, tema: dict) -> str:
    """Monogram SVG tipografis; font sistem yang mendekati font display."""
    ini = inisial(nama)
    acc, bg = tema["accent"], tema["accent_text"]
    bentuk = tema["bentuk_logo"]
    serif = "Georgia, 'Times New Roman', serif" if "serif" in \
        tema["fonts"][0].lower() and "sans" not in tema["fonts"][0].lower() \
        else "Arial, Helvetica, sans-serif"
    teks = (f'<text x="64" y="82" text-anchor="middle" font-family="{serif}" '
            f'font-size="52" font-weight="700" fill="{bg}" '
            f'letter-spacing="2">{ini}</text>')
    if bentuk == "circle":
        shape = f'<circle cx="64" cy="64" r="60" fill="{acc}"/>'
    elif bentuk == "arch":
        shape = (f'<path d="M8 128 V64 A56 56 0 0 1 120 64 V128 Z" '
                 f'fill="{acc}"/>')
    elif bentuk == "sharp":
        shape = f'<rect x="4" y="4" width="120" height="120" fill="{acc}"/>'
    elif bentuk == "minimal":
        return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 128 128">'
                f'<rect x="10" y="10" width="108" height="108" fill="none" '
                f'stroke="{acc}" stroke-width="5"/>{teks.replace(bg, acc)}'
                f'</svg>')
    else:  # square
        shape = (f'<rect x="6" y="6" width="116" height="116" rx="22" '
                 f'fill="{acc}"/>')
    return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 128 128">'
            f'{shape}{teks}</svg>')


# ---------------------------------------------------------------------------
# Util warna untuk mode gelap/terang otomatis
# ---------------------------------------------------------------------------
def _rgb(heks: str):
    heks = heks.lstrip("#")
    return tuple(int(heks[i:i + 2], 16) for i in (0, 2, 4))


def _heks(rgb) -> str:
    return "#{:02x}{:02x}{:02x}".format(
        *(max(0, min(255, round(c))) for c in rgb))


def _campur(c1: str, c2: str, p: float) -> str:
    """Campur c1 menuju c2 sebesar p (0..1)."""
    a, b = _rgb(c1), _rgb(c2)
    return _heks(tuple(a[i] + (b[i] - a[i]) * p for i in range(3)))


def _luminansi(heks: str) -> float:
    def lin(c):
        c /= 255
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = _rgb(heks)
    return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b)


def _kontras(c1: str, c2: str) -> float:
    l1, l2 = sorted((_luminansi(c1), _luminansi(c2)), reverse=True)
    return (l1 + .05) / (l2 + .05)


def _teks_untuk(aksen: str) -> str:
    """Pilih warna teks paling kontras di atas warna aksen."""
    gelap, terang = "#17140f", "#f7f4ec"
    return gelap if _kontras(aksen, gelap) > _kontras(aksen, terang) else terang


def mode_alternatif(t: dict):
    """Turunkan palet mode lawan dari tema dasar.

    Tema dasar terang -> mode alternatif 'gelap', dan sebaliknya.
    Bisa dioverride per cafe lewat dict TEMA[<slug>]['alt'].
    """
    if _luminansi(t["bg"]) > 0.3:  # dasar terang -> gelap
        bg = _campur(t["bg"], "#0d0d10", 0.90)
        alt = {
            "bg": bg,
            "surface": _campur(bg, "#ffffff", 0.055),
            "text": _campur("#f3f0e9", t["bg"], 0.06),
            "muted": _campur("#aba69c", t["bg"], 0.10),
            "accent": t["accent"],
            "accent_text": t["accent_text"],
        }
        if _kontras(alt["accent"], bg) < 3.0:
            # aksen terlalu gelap untuk latar gelap -> cerahkan
            p = 0.85 if _luminansi(alt["accent"]) < 0.06 else 0.38
            alt["accent"] = _campur(alt["accent"], "#ffffff", p)
            alt["accent_text"] = _teks_untuk(alt["accent"])
        nama = "gelap"
    else:  # dasar gelap -> terang
        bg = _campur("#f6f4ef", t["bg"], 0.05)
        alt = {
            "bg": bg,
            "surface": _campur(bg, "#000000", 0.05),
            "text": _campur("#211d18", t["bg"], 0.08),
            "muted": _campur("#6b665c", t["bg"], 0.08),
            "accent": t["accent"],
            "accent_text": t["accent_text"],
        }
        if _kontras(alt["accent"], bg) < 2.5:
            # aksen terlalu terang untuk latar terang -> gelapkan
            alt["accent"] = _campur(alt["accent"], "#000000", 0.30)
            alt["accent_text"] = _teks_untuk(alt["accent"])
        nama = "terang"
    alt.update(t.get("alt", {}))
    return nama, alt


def theme_css(t: dict) -> str:
    nama_alt, alt = mode_alternatif(t)
    dasar = f""":root {{
  --bg:{t['bg']}; --surface:{t['surface']}; --text:{t['text']};
  --muted:{t['muted']}; --accent:{t['accent']};
  --accent-text:{t['accent_text']};
  --hero-veil:{t['hero_veil']}; --hero-text:{t['hero_text']};
  --radius:{t['radius']}; --radius-lg:{t['radius_lg']};
  --font-display:{t['fonts'][0]}; --font-body:{t['fonts'][1]};
  --tt-display:{t['tt_display']}; --ls-display:{t['ls_display']};
  --fw-display:{t['fw_display']};
  --tt-brand:{t['tt_brand']}; --ls-brand:{t['ls_brand']};
}}"""
    alt_css = f"""html[data-mode="{nama_alt}"] {{
  --bg:{alt['bg']}; --surface:{alt['surface']}; --text:{alt['text']};
  --muted:{alt['muted']}; --accent:{alt['accent']};
  --accent-text:{alt['accent_text']};
}}"""
    return dasar + "\n" + alt_css


def menu_html(menu) -> str:
    out = []
    for kategori, items in menu:
        out.append(f'      <div class="menu-cat">{kategori}</div>')
        for nama, desc, harga in items:
            out.append(
                f'      <div class="menu-item"><div>'
                f'<div class="nama-menu">{nama}</div>'
                f'<div class="desc-menu">{desc}</div></div>'
                f'<span class="harga">{harga}</span></div>')
    return "\n".join(out)


def format_wa(wa: str) -> str:
    return f"+{wa[:2]} {wa[2:5]}-{wa[5:9]}-{wa[9:]}"


def main() -> None:
    path = sys.argv[1] if len(sys.argv) > 1 else "hasil_kota.xlsx"
    kota = re.sub(r"^hasil_|\.xlsx$", "", Path(path).stem).title() or "Kota"
    template = Path("demo_template_pro.html").read_text(encoding="utf-8")

    wb = openpyxl.load_workbook(path)
    ws = wb["Prospek Cafe"]
    data = {}
    for r in range(4, ws.max_row + 1):
        nama = ws.cell(r, KOL_NAMA).value
        if not nama:
            continue
        peta = ws.cell(r, KOL_PETA)
        ig = str(ws.cell(r, KOL_IG).value or "").strip()
        data[slugify(str(nama))] = {
            "nama": str(nama).strip(),
            "alamat": str(ws.cell(r, KOL_ALAMAT).value or kota),
            "wa": ambil_wa(ws.cell(r, KOL_WA)),
            "ig": "" if ig == "-" else ig,
            "maps": str(peta.hyperlink.target if peta.hyperlink
                        else peta.value or ""),
        }

    dibuat = 0
    for slug, d in data.items():
        tema = TEMA["kafe-contoh"]
        aset = Path("aset_cafe") / slug
        if not d or not aset.is_dir():
            print(f"[SKIP] {slug}: data/aset tidak ada")
            continue

        foto = sorted(aset.glob("img*.jpg"))
        foto = [f for i, f in enumerate(foto, 1)
                if i not in tema.get("skip_foto", [])]
        if not foto:
            print(f"[SKIP] {slug}: tanpa foto")
            continue

        out = Path("demo_sites") / slug
        (out / "img").mkdir(parents=True, exist_ok=True)
        nama_file = []
        for i, f in enumerate(foto, 1):
            shutil.copy2(f, out / "img" / f"foto{i}.jpg")
            nama_file.append(f"img/foto{i}.jpg")

        logo_asli = aset / "logo.jpg"
        if logo_asli.exists():
            shutil.copy2(logo_asli, out / "logo.jpg")
            logo_src = "logo.jpg"
        else:
            (out / "logo.svg").write_text(
                buat_logo_svg(d["nama"], tema), encoding="utf-8")
            logo_src = "logo.svg"

        wa = d["wa"]
        if wa:
            wa_link = (f"https://wa.me/{wa}?text=Halo%20"
                       f"{d['nama'].replace(' ', '%20')}%2C%20saya%20mau%20"
                       "pesan%20%2F%20reservasi%20tempat")
            wa_label = "Pesan via WhatsApp"
        elif d["ig"].startswith("http"):
            # tanpa WA -> tombol CTA mengarah ke DM Instagram
            wa_link = d["ig"]
            wa_label = "DM Instagram Kami"
        else:
            wa_link = "#"
            wa_label = "Pesan via WhatsApp"
        wa_disp = format_wa(wa) if wa else ("DM Instagram" if d["ig"] else "-")
        galeri = "\n".join(
            f'      <img src="{f}" alt="Suasana {d["nama"]}" loading="lazy">'
            for f in nama_file[:5])

        hal = (template
               .replace("{{FONT_LINK}}", tema["font_link"])
               .replace("{{THEME_CSS}}", theme_css(tema))
               .replace("{{MODE_ALT}}", mode_alternatif(tema)[0])
               .replace("{{NAMA}}", d["nama"])
               .replace("{{KOTA}}", kota)
               .replace("{{TAGLINE}}", tema["tagline"])
               .replace("{{JUDUL_TENTANG}}", tema["judul_tentang"])
               .replace("{{DESKRIPSI}}", tema["deskripsi"])
               .replace("{{FOTO_HERO}}", nama_file[0])
               .replace("{{FOTO_TENTANG}}",
                        nama_file[1] if len(nama_file) > 1 else nama_file[0])
               .replace("{{GALERI}}", galeri)
               .replace("{{MENU_ITEM}}", menu_html(tema["menu"]))
               .replace("{{ALAMAT_SINGKAT}}", d["alamat"].split(",")[0])
               .replace("{{ALAMAT}}", d["alamat"])
               .replace("{{JAM}}", tema["jam"])
               .replace("{{MAPS_LINK}}", d["maps"] or "#")
               .replace("{{WA_LINK}}", wa_link)
               .replace("{{WA_LABEL}}", wa_label)
               .replace("{{WA_DISPLAY}}", wa_disp)
               .replace("{{TAHUN}}", TAHUN))
        if logo_src == "logo.jpg":
            hal = hal.replace("logo.svg", "logo.jpg")

        (out / "index.html").write_text(hal, encoding="utf-8")
        dibuat += 1
        jenis_logo = "asli" if logo_asli.exists() else "monogram"
        print(f"[OK] {d['nama']} -> demo_sites/{slug}/ "
              f"({len(nama_file)} foto, logo {jenis_logo})")

    print(f"\n[INFO] {dibuat} demo dibuat.")
    if NAMA_ANDA == "Nama Kamu":
        print("[WARN] NAMA_ANDA/WA_ANDA belum diedit di buat_demo_pro.py!")


if __name__ == "__main__":
    main()
