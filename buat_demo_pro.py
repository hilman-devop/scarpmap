"""Generator demo website PRO per cafe -- anti-template, anti-slop.

Tiap cafe punya konfigurasi desain sendiri (tema warna dari identitas
visualnya, pasangan font Google Fonts yang cocok dengan vibe, bentuk logo
monogram, menu sesuai jenis usaha). Foto memakai foto ASLI dari Google
Maps yang sudah diunduh oleh ambil_aset_cafe.py ke aset_cafe/<slug>/.

Logo: jika aset_cafe/<slug>/logo.jpg ada (foto profil IG asli), dipakai.
Jika tidak, dibuat monogram SVG tipografis yang disetel per tema.

Pemakaian:
    python buat_demo_pro.py hasil_jember.xlsx
Output: demo_sites/<slug>/{index.html, logo.svg, img/...}
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
NAMA_ANDA = "Hilman Arifin Ilham"
WA_ANDA = "+62 823-3048-2713"

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
    "slovv-coffee": {
        # Signage hitam-putih modern, interior terang hangat -> dark elegan
        "fonts": ("'Fraunces', Georgia, serif", "'Inter', Arial, sans-serif"),
        "font_link": font_link("family=Fraunces:opsz,wght@9..144,500;9..144,600",
                               "family=Inter:wght@400;600;700"),
        "bg": "#111113", "surface": "#1a1b1e", "text": "#f2efe9",
        "muted": "#a9a49a", "accent": "#c9a25e", "accent_text": "#141312",
        "hero_veil": "rgba(10,10,11,.88)", "hero_text": "#f5f2ec",
        "radius": "6px", "radius_lg": "14px",
        "tt_display": "none", "ls_display": "-0.01em", "fw_display": "600",
        "tt_brand": "uppercase", "ls_brand": "0.14em",
        "bentuk_logo": "circle",
        "tagline": "Chill space di tengah Jember — tempat melambat, ngopi "
                   "tenang, dan kerja tanpa buru-buru.",
        "judul_tentang": "Ruang untuk Melambat",
        "deskripsi": "Slovv adalah chill space bergaya modern minimalis di "
                     "Jember. Interior terang, kursi nyaman, dan kopi yang "
                     "diseduh serius — cocok untuk kerja, baca buku, atau "
                     "sekadar kabur sebentar dari hiruk-pikuk.",
        "jam": "10.00–23.00", "menu": MENU_KOPI,
    },
    "grand-cafe-jember": {
        # Sofa merah, kayu, tegel klasik -> heritage hangat
        "fonts": ("'Playfair Display', Georgia, serif",
                  "'DM Sans', Arial, sans-serif"),
        "font_link": font_link("family=Playfair+Display:wght@500;600;700",
                               "family=DM+Sans:opsz,wght@9..40,400;9..40,600;9..40,700"),
        "bg": "#faf4e8", "surface": "#f1e7d3", "text": "#2e2018",
        "muted": "#6d5c4c", "accent": "#9e2b25", "accent_text": "#fbf5ea",
        "hero_veil": "rgba(30,16,10,.82)", "hero_text": "#faf3e7",
        "radius": "10px", "radius_lg": "18px",
        "tt_display": "none", "ls_display": "0", "fw_display": "600",
        "tt_brand": "none", "ls_brand": "0.02em",
        "bentuk_logo": "arch",
        "tagline": "Cafe & resto bernuansa klasik hangat — tempat keluarga "
                   "dan cerita berkumpul di Jember.",
        "judul_tentang": "Hangat Seperti Rumah",
        "deskripsi": "Grand Cafe Jember memadukan sentuhan klasik dan "
                     "kenyamanan rumahan: sofa empuk, kayu hangat, dan "
                     "lantai tegel yang penuh karakter. Tempat yang pas "
                     "untuk makan bersama keluarga atau ngobrol panjang.",
        "jam": "09.00–22.00", "menu": MENU_EATERY,
    },
    "randevoo-coffee-jember": {
        # Fasad beton putih, logo garis tipis -> minimal modern terang
        "fonts": ("'Space Grotesk', Arial, sans-serif",
                  "'Inter', Arial, sans-serif"),
        "font_link": font_link("family=Space+Grotesk:wght@400;500;700",
                               "family=Inter:wght@400;600"),
        "bg": "#f7f7f5", "surface": "#ececea", "text": "#171717",
        "muted": "#6f6f6c", "accent": "#171717", "accent_text": "#f7f7f5",
        "hero_veil": "rgba(15,15,15,.78)", "hero_text": "#f7f7f5",
        "radius": "2px", "radius_lg": "4px",
        "tt_display": "uppercase", "ls_display": "0.06em", "fw_display": "500",
        "tt_brand": "uppercase", "ls_brand": "0.22em",
        "bentuk_logo": "minimal",
        "tagline": "Coffee & space bergaya minimal — kopi jujur, ruang "
                   "lapang, dan suasana yang bikin betah berlama-lama.",
        "judul_tentang": "Sederhana, Sengaja",
        "deskripsi": "Randevoo percaya hal baik tidak perlu berlebihan. "
                     "Ruang beton terang, garis bersih, dan secangkir kopi "
                     "yang dibuat dengan presisi — semua dirancang agar kamu "
                     "fokus pada momennya.",
        "jam": "09.00–22.00", "menu": MENU_KOPI,
    },
    "nugas-jember-coffee-food-and-space": {
        # Taman outdoor + pendopo kayu -> natural hijau
        "fonts": ("'DM Serif Display', Georgia, serif",
                  "'Nunito Sans', Arial, sans-serif"),
        "font_link": font_link("family=DM+Serif+Display",
                               "family=Nunito+Sans:opsz,wght@6..12,400;6..12,700"),
        "bg": "#f4f1e6", "surface": "#e8e3d2", "text": "#24301f",
        "muted": "#5d6b54", "accent": "#4a6b3f", "accent_text": "#f6f4ea",
        "hero_veil": "rgba(18,26,14,.8)", "hero_text": "#f4f2e8",
        "radius": "12px", "radius_lg": "20px",
        "tt_display": "none", "ls_display": "0", "fw_display": "400",
        "tt_brand": "none", "ls_brand": "0.02em",
        "bentuk_logo": "circle",
        "tagline": "Nugas, makan, atau santai di taman — coffee, food & "
                   "space paling adem di Jember.",
        "judul_tentang": "Ademnya Nugas di Sini",
        "deskripsi": "Sesuai namanya, Nugas dibuat untuk kamu yang butuh "
                     "tempat produktif tanpa sumpek. Area taman outdoor, "
                     "pendopo kayu, dan kopi terjangkau — andalan mahasiswa "
                     "dan pekerja lepas Jember.",
        "jam": "09.00–23.00", "menu": MENU_EATERY,
    },
    "discuss-space-coffee": {
        # Fasad gelap, interior hangat, tempat nongkrong & diskusi
        "fonts": ("'Sora', Arial, sans-serif", "'Inter', Arial, sans-serif"),
        "font_link": font_link("family=Sora:wght@400;600;700",
                               "family=Inter:wght@400;600"),
        "bg": "#16130f", "surface": "#211c16", "text": "#f3ede2",
        "muted": "#b3a692", "accent": "#e0a458", "accent_text": "#1c150c",
        "hero_veil": "rgba(12,9,6,.85)", "hero_text": "#f5efe4",
        "radius": "999px", "radius_lg": "26px",
        "tt_display": "none", "ls_display": "-0.02em", "fw_display": "700",
        "tt_brand": "lowercase", "ls_brand": "0.04em",
        "bentuk_logo": "circle",
        "tagline": "Space & coffee untuk diskusi panjang, kerja kelompok, "
                   "dan obrolan yang tidak selesai-selesai.",
        "judul_tentang": "Datang untuk Kopinya, Tinggal untuk Obrolannya",
        "deskripsi": "Discuss adalah ruang bertemu di jantung Jember: "
                     "interior hangat, meja-meja nyaman untuk berdiskusi, "
                     "dan kopi yang menemani sampai obrolan usai.",
        "jam": "10.00–23.00", "menu": MENU_KOPI,
        "skip_foto": [1],  # img01 = screenshot UI ponsel
    },
    "kedai-kopi-wijaya-sumbersari-by-pak-yon-": {
        # Kopi gelas klasik, kayu gelap -> kedai kopi rustic
        "fonts": ("'Lora', Georgia, serif",
                  "'Source Sans 3', Arial, sans-serif"),
        "font_link": font_link("family=Lora:wght@500;600;700",
                               "family=Source+Sans+3:wght@400;600;700"),
        "bg": "#f6efe1", "surface": "#ece0c9", "text": "#33231a",
        "muted": "#7a6753", "accent": "#8c4a1f", "accent_text": "#faf3e6",
        "hero_veil": "rgba(32,20,10,.82)", "hero_text": "#f8f1e3",
        "radius": "8px", "radius_lg": "14px",
        "tt_display": "none", "ls_display": "0", "fw_display": "600",
        "tt_brand": "none", "ls_brand": "0.02em",
        "bentuk_logo": "circle",
        "tagline": "Kedai kopi legendaris Sumbersari — kopi tubruk dan "
                   "suasana ngopi yang sesungguhnya.",
        "judul_tentang": "Ngopi Seperti Dulu, Rasanya Tetap Juara",
        "deskripsi": "Kedai Kopi Wijaya menjaga tradisi ngopi yang jujur: "
                     "biji pilihan Pak Yon Coffee Supply, diseduh sederhana, "
                     "disajikan hangat. Tempat di mana kopi masih jadi alasan "
                     "utama untuk datang.",
        "jam": "08.00–22.00", "menu": MENU_KOPI,
    },
    "tigakata": {
        # Beton ekspos + terracotta -> industrial earthy
        "fonts": ("'Archivo', Arial, sans-serif",
                  "'Inter', Arial, sans-serif"),
        "font_link": font_link("family=Archivo:wght@500;700;800",
                               "family=Inter:wght@400;600"),
        "bg": "#efe9df", "surface": "#e2d9cb", "text": "#2a241d",
        "muted": "#76695a", "accent": "#b4552d", "accent_text": "#faf3e9",
        "hero_veil": "rgba(30,22,14,.8)", "hero_text": "#f3ecdf",
        "radius": "2px", "radius_lg": "6px",
        "tt_display": "uppercase", "ls_display": "0.04em", "fw_display": "800",
        "tt_brand": "uppercase", "ls_brand": "0.2em",
        "bentuk_logo": "sharp",
        "tagline": "Tiga kata, satu tempat: kopi, ruang, dan cerita — di "
                   "tengah tekstur beton dan tanah liat Jember.",
        "judul_tentang": "Mentah Itu Jujur",
        "deskripsi": "TIGAKATA merangkul material apa adanya: dinding beton "
                     "ekspos, lengkung bata, dan cahaya sore yang masuk "
                     "perlahan. Tempat ngopi untuk kamu yang suka ruang "
                     "berkarakter.",
        "jam": "10.00–23.00", "menu": MENU_KOPI,
    },
    "fox-coffee-roastery": {
        # Taman rimbun + pintu merah vintage -> eklektik hangat
        "fonts": ("'Fraunces', Georgia, serif", "'Karla', Arial, sans-serif"),
        "font_link": font_link("family=Fraunces:opsz,wght@9..144,500;9..144,700",
                               "family=Karla:wght@400;600;700"),
        "bg": "#f7f2e7", "surface": "#ede4d2", "text": "#2b2118",
        "muted": "#71624f", "accent": "#a03528", "accent_text": "#faf3e6",
        "hero_veil": "rgba(25,16,10,.8)", "hero_text": "#f7f1e4",
        "radius": "10px", "radius_lg": "16px",
        "tt_display": "none", "ls_display": "0", "fw_display": "700",
        "tt_brand": "none", "ls_brand": "0.03em",
        "bentuk_logo": "arch",
        "tagline": "Coffee & roastery di tengah taman rimbun — biji disangrai "
                   "sendiri, diseduh dengan bangga.",
        "judul_tentang": "Dari Roastery ke Cangkirmu",
        "deskripsi": "Fox Coffee & Roastery menyangrai bijinya sendiri, lalu "
                     "menyeduhnya di ruang hijau yang teduh dan penuh "
                     "karakter. Tempat untuk penikmat kopi yang ingin tahu "
                     "dari mana kopinya berasal.",
        "jam": "09.00–22.00", "menu": MENU_KOPI,
    },
    "xoxo-toko": {
        # Fasad kayu emas + wisteria ungu -> elegan playful (gelato)
        "fonts": ("'Playfair Display', Georgia, serif",
                  "'Jost', Arial, sans-serif"),
        "font_link": font_link("family=Playfair+Display:ital,wght@0,500;0,700;1,500",
                               "family=Jost:wght@400;500;600"),
        "bg": "#fbf7f0", "surface": "#f3ead9", "text": "#2e2436",
        "muted": "#7d7286", "accent": "#b98a2f", "accent_text": "#fffaf0",
        "hero_veil": "rgba(28,20,34,.8)", "hero_text": "#faf5ef",
        "radius": "18px", "radius_lg": "28px",
        "tt_display": "none", "ls_display": "0", "fw_display": "700",
        "tt_brand": "uppercase", "ls_brand": "0.18em",
        "bentuk_logo": "circle",
        "tagline": "Toko, cafe & gelato — manisnya jajan cantik di bawah "
                   "kanopi wisteria Jember.",
        "judul_tentang": "Tempat Jajan Paling Cantik",
        "deskripsi": "XOXO memadukan toko, cafe, dan gelato dalam satu bangunan "
                     "yang memikat: panel kayu keemasan, bunga wisteria, dan "
                     "gelato yang dibuat segar. Datang untuk fotonya, kembali "
                     "untuk rasanya.",
        "jam": "10.00–22.00", "menu": MENU_GELATO,
    },
    "kopi-kampus-jember": {
        # Mural jalanan "NEVER QUIT" -> bold street muda
        "fonts": ("'Bebas Neue', 'Arial Narrow', sans-serif",
                  "'Space Grotesk', Arial, sans-serif"),
        "font_link": font_link("family=Bebas+Neue",
                               "family=Space+Grotesk:wght@400;500;700"),
        "bg": "#f5f5f0", "surface": "#e8e8e0", "text": "#141414",
        "muted": "#5c5c56", "accent": "#d43d2a", "accent_text": "#fdfaf4",
        "hero_veil": "rgba(10,10,10,.8)", "hero_text": "#f6f6f0",
        "radius": "0px", "radius_lg": "0px",
        "tt_display": "uppercase", "ls_display": "0.03em", "fw_display": "400",
        "tt_brand": "uppercase", "ls_brand": "0.16em",
        "bentuk_logo": "sharp",
        "tagline": "Kopi mahasiswa, harga bersahabat — basecamp nugas dan "
                   "nongkrong paling ramah kantong di Jember.",
        "judul_tentang": "Never Quit, Ngopi Dulu",
        "deskripsi": "Kopi Kampus lahir dari semangat anak muda Jember: kopi "
                     "enak tanpa bikin dompet menangis, tempat luas untuk "
                     "nugas, rapat organisasi, atau sekadar istirahat di "
                     "sela kuliah.",
        "jam": "10.00–23.00", "menu": MENU_KOPI,
    },
    "perasa-coffee-eatery": {
        # Fasad hitam-putih modern, konter kuning backlight, meja kerja laptop
        # -> modern hangat rapi, aksen mustard
        "fonts": ("'Plus Jakarta Sans', Arial, sans-serif",
                  "'Inter', Arial, sans-serif"),
        "font_link": font_link("family=Plus+Jakarta+Sans:wght@500;700;800",
                               "family=Inter:wght@400;600"),
        "bg": "#faf6ee", "surface": "#f1e9d8", "text": "#241c10",
        "muted": "#7a6c54", "accent": "#d99a2b", "accent_text": "#231a0c",
        "hero_veil": "rgba(24,16,6,.85)", "hero_text": "#faf3e6",
        "radius": "14px", "radius_lg": "24px",
        "tt_display": "none", "ls_display": "-0.02em", "fw_display": "800",
        "tt_brand": "lowercase", "ls_brand": "0",
        "bentuk_logo": "square",
        "tagline": "Coffee & eatery modern di Jember — tempat kerja, nugas, "
                   "dan jeda sore yang selalu terasa pas.",
        "judul_tentang": "Rasa yang Bikin Kembali",
        "deskripsi": "Perasa memadukan kopi susu andalan, camilan hangat, "
                     "dan suasana kerja yang nyaman: colokan melimpah, WiFi "
                     "kencang, dan konter kuning ikonik yang jadi ciri "
                     "khasnya.",
        "jam": "10.00–22.00", "menu": MENU_EATERY,
    },
    "dekatsini-coffee-eatery": {
        # Beton putih + furnitur hitam + meja kayu -> industrial hangat gelap
        "fonts": ("'Bricolage Grotesque', Arial, sans-serif",
                  "'Inter', Arial, sans-serif"),
        "font_link": font_link("family=Bricolage+Grotesque:opsz,wght@12..96,500;12..96,700",
                               "family=Inter:wght@400;600"),
        "bg": "#141517", "surface": "#1d1e21", "text": "#eae7e0",
        "muted": "#9d998f", "accent": "#c87343", "accent_text": "#17110b",
        "hero_veil": "rgba(10,10,12,.86)", "hero_text": "#efede7",
        "radius": "4px", "radius_lg": "10px",
        "tt_display": "uppercase", "ls_display": "0.02em", "fw_display": "700",
        "tt_brand": "uppercase", "ls_brand": "0.18em",
        "bentuk_logo": "minimal",
        "tagline": "Sesuai namanya — kopi dan makanan enak yang selalu "
                   "'dekat sini' dari aktivitasmu di Jember.",
        "judul_tentang": "Selalu Dekat, Selalu Pas",
        "deskripsi": "Dekatsini menghadirkan suasana industrial yang hangat: "
                     "beton, kayu, dan lampu temaram. Tempat singgah cepat "
                     "atau duduk lama — kopi, rice bowl, dan camilan selalu "
                     "siap menemani.",
        "jam": "10.00–22.00", "menu": MENU_EATERY,
    },
    "navas": {
        # Bangunan kaca + taman rimbun, foto malam lampu hangat, rice bowl
        # -> elegan taman malam: hijau gelap + emas
        "fonts": ("'Marcellus', Georgia, serif",
                  "'Inter', Arial, sans-serif"),
        "font_link": font_link("family=Marcellus",
                               "family=Inter:wght@400;600"),
        "bg": "#101a15", "surface": "#182620", "text": "#ece7d8",
        "muted": "#9aa892", "accent": "#c8a45c", "accent_text": "#131a10",
        "hero_veil": "rgba(8,14,10,.86)", "hero_text": "#f2eddd",
        "radius": "10px", "radius_lg": "18px",
        "tt_display": "none", "ls_display": "0.02em", "fw_display": "400",
        "tt_brand": "uppercase", "ls_brand": "0.28em",
        "bentuk_logo": "circle",
        "tagline": "Coffee & eatery di tengah taman — meja marmer, tanaman "
                   "rimbun, dan malam Jember yang hangat.",
        "judul_tentang": "Makan Enak di Tengah Taman",
        "deskripsi": "Navas adalah oase di tengah kota: bangunan kaca yang "
                     "dikelilingi taman, lampu hangat saat malam, dan menu "
                     "rice bowl hingga camilan yang cocok untuk makan malam "
                     "santai maupun kumpul keluarga.",
        "jam": "10.00–22.00", "menu": MENU_EATERY,
    },
    "soen-coffee-eatery": {
        # Signage 'soen' bulat hijau menyala, bata ekspos, pasta/toast/rice
        # -> kasual segar hijau, ramah dan playful
        "fonts": ("'Poppins', Arial, sans-serif",
                  "'Inter', Arial, sans-serif"),
        "font_link": font_link("family=Poppins:wght@500;600;700",
                               "family=Inter:wght@400;600"),
        "bg": "#fbf9f3", "surface": "#f0ede1", "text": "#22301f",
        "muted": "#6f7a66", "accent": "#3e8f4f", "accent_text": "#f7faf2",
        "hero_veil": "rgba(16,26,16,.82)", "hero_text": "#f4f7ee",
        "radius": "999px", "radius_lg": "22px",
        "tt_display": "none", "ls_display": "-0.01em", "fw_display": "600",
        "tt_brand": "lowercase", "ls_brand": "0.01em",
        "bentuk_logo": "circle",
        "tagline": "Coffee & eatery yang segar dan ramah — dari pasta, "
                   "toast, sampai nasi goreng, semua dibuat dengan senang.",
        "judul_tentang": "Tempat Cerah untuk Hari Apa Pun",
        "deskripsi": "Soen menyambut dengan signage hijau ikonik dan interior "
                     "bata yang hangat. Menu beragam dari kopi, pasta, "
                     "sampai dessert — pas untuk brunch, kerja santai, atau "
                     "makan malam bareng teman.",
        "jam": "10.00–22.00", "menu": MENU_EATERY,
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
    path = sys.argv[1] if len(sys.argv) > 1 else "hasil_jember.xlsx"
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
    for slug, tema in TEMA.items():
        d = data.get(slug)
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
               .replace("{{TAHUN}}", TAHUN)
               .replace("{{NAMA_ANDA}}", NAMA_ANDA)
               .replace("{{WA_ANDA_LINK}}", re.sub(r"\D", "", WA_ANDA))
               .replace("{{WA_ANDA}}", WA_ANDA))
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
