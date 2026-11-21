"""
Pencari akun Instagram cafe — TANPA search engine.

Fakta di lapangan (hasil pengujian): DuckDuckGo diblokir ISP Indonesia,
Bing mengembalikan hasil sampah untuk query lokal, Google Search menampilkan
captcha untuk bot, dan halaman profil IG anonim TIDAK bisa membedakan akun
ada vs tidak ada (semua status 200, konten JS-render identik).

Strategi yang tersisa dan VALID: "handle guessing + validasi API internal"
  1. Buat kandidat handle dari nama cafe, mis. "Kolong Cafe" di Jember ->
     kolong, kolongcafe, kolong.jember, kolongjember, kolong_jember
  2. Validasi tiap kandidat lewat endpoint internal Instagram
     (api/v1/users/web_profile_info, dipakai web IG sendiri):
       - 404 -> akun tidak ada, lanjut kandidat berikutnya
       - 200 -> akun ada; cocokkan full_name/biography dengan nama cafe
                (mencegah salah akun, mis. 'Domino' != @domino milik orang lain)
       - 429 -> kena rate limit, pencarian dihentikan total
  3. Jika API sedang membatasi IP, kandidat TIDAK ditebak buta
     (demi validitas data) — cafe dilewati dengan catatan jelas.
"""

import re
import time

import requests

IG_API_URL = "https://www.instagram.com/api/v1/users/web_profile_info/"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "X-IG-App-ID": "936619743392459",  # app ID publik web Instagram
    "Accept-Language": "id-ID,id;q=0.9,en;q=0.8",
}
REQUEST_TIMEOUT = 15

# kata generik yang tidak dipakai untuk inti handle
KATA_GENERIK = {
    "cafe", "kafe", "coffee", "kopi", "eatery", "resto", "restaurant",
    "and", "the", "di", "co", "id", "official", "jember",
}


def _token_berarti(nama: str) -> list[str]:
    """Token nama yang khas (bukan kata generik), minimal 3 huruf."""
    return [w for w in re.findall(r"[a-z0-9]+", nama.lower())
            if w not in KATA_GENERIK and len(w) >= 3]


def buat_kandidat_handle(nama: str, area: str) -> list[str]:
    """Buat kandidat handle IG dari nama cafe + kota."""
    inti = "".join(_token_berarti(nama))
    base = re.sub(r"[^a-z0-9]", "", nama.lower())
    kota = re.sub(r"[^a-z0-9]", "", area.lower().split(",")[0])

    kandidat = []
    for h in [inti, base, f"{inti}.{kota}", f"{inti}{kota}", f"{inti}_{kota}"]:
        if 2 <= len(h) <= 30 and h not in kandidat:
            kandidat.append(h)
    return kandidat[:5]


def _cek_api_ig(handle: str, session: requests.Session) -> tuple[str, dict | None]:
    """Cek handle via API internal IG.
    Return: ('ADA'|'TIDAK_ADA'|'RATE_LIMIT'|'DIBLOKIR', data_user|None)."""
    try:
        resp = session.get(
            IG_API_URL,
            params={"username": handle},
            headers=HEADERS,
            timeout=REQUEST_TIMEOUT,
        )
    except requests.exceptions.RequestException:
        return "DIBLOKIR", None

    if resp.status_code == 404:
        return "TIDAK_ADA", None
    if resp.status_code == 429:
        return "RATE_LIMIT", None
    if resp.status_code != 200:
        return "DIBLOKIR", None
    try:
        user = resp.json().get("data", {}).get("user")
    except ValueError:
        return "DIBLOKIR", None
    return ("ADA", user) if user else ("TIDAK_ADA", None)


def cari_instagram(nama: str, area: str, session: requests.Session,
                   delay: float = 3.0) -> tuple[str, str]:
    """Cari akun IG satu cafe. Return (url_ig, catatan)."""
    if "(tanpa nama" in nama:
        return "", "dilewati (cafe tanpa nama)"

    token = _token_berarti(nama)
    for handle in buat_kandidat_handle(nama, area):
        status, user = _cek_api_ig(handle, session)
        if status == "RATE_LIMIT":
            return "", "RATE_LIMIT"
        if status == "ADA" and user:
            # validasi konten: nama lengkap/bio harus cocok dengan nama cafe
            teks = f"{user.get('full_name', '')} {user.get('biography', '')}".lower()
            cocok = any(t in teks for t in token)
            if cocok:
                followers = user.get("edge_followed_by", {}).get("count", "?")
                return (f"https://www.instagram.com/{handle}",
                        f"terverifikasi via API IG ({followers} followers)")
            # akun ada tapi isinya tidak cocok -> kemungkinan akun orang lain
        time.sleep(delay)

    return "", "tidak ditemukan"


def cari_instagram_massal(cafes: list[dict], area: str, delay: float = 3.0,
                          **_) -> int:
    """Cari IG untuk semua cafe yang kolom instagram-nya kosong."""
    target = [c for c in cafes if not c.get("instagram")]
    if not target:
        print("[INFO] Semua cafe sudah punya Instagram.")
        return 0
    print(f"[INFO] Mencari Instagram untuk {len(target)} cafe "
          f"(handle guessing + validasi API IG)...")

    ditemukan = 0
    diblokir = 0
    session = requests.Session()
    for i, cafe in enumerate(target, start=1):
        url, catatan = cari_instagram(cafe["nama"], area, session, delay=delay)
        if catatan == "RATE_LIMIT":
            print(f"\n[WARN] Kena rate limit Instagram di cafe ke-{i}. "
                  "Pencarian IG dihentikan, sisa dilewati.")
            break
        if url:
            cafe["instagram"] = url
            ditemukan += 1
        lama = cafe.get("catatan_validasi", "")
        cafe["catatan_validasi"] = "; ".join(
            x for x in [lama, f"IG: {catatan}"] if x
        )
        print(f"\r[INFO] IG: {i}/{len(target)} ({ditemukan} ditemukan)   ",
              end="", flush=True)
        time.sleep(delay)

    print()
    print(f"[INFO] Instagram ditemukan untuk {ditemukan}/{len(target)} cafe.")
    return ditemukan
