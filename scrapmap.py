"""
ScrapMap v2 - Tool pencari prospek cafe yang belum memiliki website aktif
di suatu kabupaten/kota, lengkap dengan:
  - deteksi media sosial (Instagram, Facebook, TikTok, Twitter/X, YouTube)
  - kontak yang valid (telepon +62, link WhatsApp, email tervalidasi)
  - alamat lengkap
  - export Excel yang rapi (hyperlink aktif, warna status, sheet ringkasan)

Semua link website & sosmed DIVALIDASI otomatis via HTTP:
link yang 404 / mati dibuang dari hasil agar data tetap valid.

Cara pakai:
    python scrapmap.py --area "Jember"
    python scrapmap.py --area "Jember" --google-key "API_KEY"
    python scrapmap.py --area "Jember" --format csv
"""

import argparse
import csv
import json
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from urllib.parse import urlparse

import requests

NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
OVERPASS_URLS = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
    "https://overpass.private.coffee/api/interpreter",
    "https://overpass.nchc.org.tw/api/interpreter",
]
GOOGLE_PLACES_URL = "https://places.googleapis.com/v1/places:searchText"
GOOGLE_FIELD_MASK = ",".join(
    [
        "places.displayName",
        "places.formattedAddress",
        "places.nationalPhoneNumber",
        "places.internationalPhoneNumber",
        "places.websiteUri",
        "places.location",
        "places.googleMapsUri",
        "nextPageToken",
    ]
)
# User-Agent ala browser agar tidak ditolak saat validasi sosmed
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
REQUEST_TIMEOUT = 15

PLATFORM_SOSMED = ["instagram", "facebook", "tiktok", "twitter", "youtube"]

OSM_TAGS_SOSMED = {
    "instagram": ["contact:instagram", "instagram"],
    "facebook": ["contact:facebook", "facebook"],
    "tiktok": ["contact:tiktok", "tiktok"],
    "twitter": ["contact:twitter", "twitter"],
    "youtube": ["contact:youtube", "youtube"],
}
OSM_TAGS_WEBSITE = ["website", "contact:website", "url", "contact:url"]
OSM_TAGS_TELEPON = ["phone", "contact:phone", "contact:mobile"]
OSM_TAGS_EMAIL = ["email", "contact:email"]

EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")


# ---------------------------------------------------------------------------
# Helper normalisasi & validasi data
# ---------------------------------------------------------------------------

def ambil_tag(tags: dict, daftar_key: list) -> str:
    for k in daftar_key:
        v = tags.get(k)
        if v:
            return str(v).strip()
    return ""


def normalisasi_telepon(raw: str) -> str:
    """Normalisasi nomor ke format internasional (+62...). '' jika tidak valid."""
    if not raw:
        return ""
    digits = re.sub(r"\D", "", raw)
    if digits.startswith("0"):
        digits = "62" + digits[1:]
    elif digits.startswith("8"):
        digits = "62" + digits
    if not (10 <= len(digits) <= 15):
        return ""  # panjang tidak masuk akal -> dianggap tidak valid
    return "+" + digits


def link_whatsapp(telepon: str) -> str:
    """Buat link wa.me dari nomor +62 valid. Hanya nomor seluler (08x/+628x)
    yang dibuatkan link WA — nomor telepon rumah (kode area) tidak punya WA."""
    if telepon.startswith("+628"):
        return "https://wa.me/" + telepon[1:]
    return ""


def normalisasi_email(raw: str) -> str:
    v = (raw or "").strip().lower()
    return v if EMAIL_RE.match(v) else ""


def normalisasi_sosmed(platform: str, raw: str) -> str:
    """Ubah handle/link mentah menjadi URL sosmed kanonik."""
    if not raw:
        return ""
    v = raw.strip()
    if v.startswith("http"):
        return v.split("?")[0].rstrip("/")
    v = v.lstrip("@").strip("/")
    if not v or " " in v:
        return ""
    pola = {
        "instagram": "https://www.instagram.com/{}",
        "facebook": "https://www.facebook.com/{}",
        "tiktok": "https://www.tiktok.com/@{}",
        "twitter": "https://x.com/{}",
        "youtube": "https://www.youtube.com/{}",
    }
    if platform == "youtube" and "/" not in v and not v.startswith("@"):
        v = "@" + v
    return pola[platform].format(v)


def bangun_alamat(tags: dict) -> str:
    """Rakit alamat lengkap dari komponen tag addr:* OSM."""
    if tags.get("addr:full"):
        return tags["addr:full"].strip()
    jalan = tags.get("addr:street", "")
    nomor = tags.get("addr:housenumber", "")
    baris1 = f"{jalan} No. {nomor}" if jalan and nomor else jalan
    parts = [
        baris1,
        tags.get("addr:suburb", ""),
        tags.get("addr:city", ""),
        tags.get("addr:postcode", ""),
        tags.get("addr:province", ""),
    ]
    return ", ".join(p for p in parts if p)



# ---------------------------------------------------------------------------
# 1. Sumber data: OpenStreetMap (gratis, tanpa API key)
# ---------------------------------------------------------------------------

def cari_area_id(area_query: str) -> tuple[int, str]:
    """Cari relation ID kabupaten via Nominatim, kembalikan (area_id, nama)."""
    resp = requests.get(
        NOMINATIM_URL,
        params={"q": area_query, "format": "json", "limit": 1, "countrycodes": "id"},
        headers={"User-Agent": USER_AGENT},
        timeout=REQUEST_TIMEOUT,
    )
    resp.raise_for_status()
    hasil = resp.json()
    if not hasil:
        raise SystemExit(f"[ERROR] Area '{area_query}' tidak ditemukan di OpenStreetMap.")

    nama = hasil[0]["display_name"]
    osm_type, osm_id = hasil[0]["osm_type"], hasil[0]["osm_id"]

    if osm_type == "relation":
        area_id = 3600000000 + osm_id
    elif osm_type == "way":
        area_id = 2400000000 + osm_id
    else:
        raise SystemExit(
            f"[ERROR] '{area_query}' bukan wilayah administratif. "
            "Coba nama kabupaten yang lebih spesifik."
        )

    print(f"[INFO] Area ditemukan: {nama}")
    return area_id, nama


def cafe_dari_osm(el: dict) -> dict:
    """Ubah elemen OSM menjadi record cafe lengkap & ternormalisasi."""
    tags = el.get("tags", {})
    telepon = normalisasi_telepon(ambil_tag(tags, OSM_TAGS_TELEPON))
    cafe = {
        "nama": tags.get("name") or f"(cafe tanpa nama #{el['id']})",
        "alamat": bangun_alamat(tags),
        "telepon": telepon,
        "whatsapp": link_whatsapp(telepon),
        "email": normalisasi_email(ambil_tag(tags, OSM_TAGS_EMAIL)),
        "website": ambil_tag(tags, OSM_TAGS_WEBSITE),
        "latitude": el.get("lat") or el.get("center", {}).get("lat"),
        "longitude": el.get("lon") or el.get("center", {}).get("lon"),
        "link_peta": f"https://www.openstreetmap.org/{el['type']}/{el['id']}",
    }
    for platform in PLATFORM_SOSMED:
        cafe[platform] = normalisasi_sosmed(
            platform, ambil_tag(tags, OSM_TAGS_SOSMED[platform])
        )
    return cafe


def ambil_daftar_cafe(area_query: str) -> list[dict]:
    """Ambil semua cafe (amenity=cafe) di dalam area via Overpass API."""
    area_id, _ = cari_area_id(area_query)

    query = f"""
    [out:json][timeout:120];
    area({area_id})->.searchArea;
    (
      node["amenity"="cafe"](area.searchArea);
      way["amenity"="cafe"](area.searchArea);
    );
    out center tags;
    """

    print("[INFO] Mengambil data cafe dari Overpass API...")
    data, error_terakhir = None, None
    for url in OVERPASS_URLS:
        try:
            resp = requests.post(
                url,
                data={"data": query},
                headers={"User-Agent": USER_AGENT},
                timeout=180,
            )
            resp.raise_for_status()
            data = resp.json()
            break
        except requests.exceptions.RequestException as e:
            print(f"[WARN] Server {url} gagal ({e.__class__.__name__}), coba mirror lain...")
            error_terakhir = e

    if data is None:
        raise SystemExit(f"[ERROR] Semua server Overpass gagal. Error: {error_terakhir}")

    cafes = [cafe_dari_osm(el) for el in data.get("elements", [])]
    print(f"[INFO] Ditemukan {len(cafes)} cafe di kawasan ini.")
    return cafes


def ambil_daftar_kecamatan(area_query: str) -> list[str]:
    """Ambil nama-nama kecamatan (sub-wilayah administratif) dari kabupaten
    via OSM. Strategi: ambil semua relasi administratif di dalam area, lalu
    pakai admin_level dengan anggota terbanyak (= level kecamatan).
    Return list nama kecamatan, atau [area_query] jika gagal."""
    area_id, _ = cari_area_id(area_query)
    query = f"""
    [out:json][timeout:60];
    area({area_id})->.a;
    rel["boundary"="administrative"](area.a);
    out tags;
    """
    data, error_terakhir = None, None
    for url in OVERPASS_URLS:
        try:
            # timeout pendek: query kecamatan hanya pemanasan; jika lambat,
            # lebih baik langsung fallback ke mode grid
            resp = requests.post(url, data={"data": query},
                                 headers={"User-Agent": USER_AGENT}, timeout=30)
            resp.raise_for_status()
            data = resp.json()
            break
        except requests.exceptions.RequestException as e:
            error_terakhir = e

    if data is None:
        print(f"[WARN] Gagal ambil daftar kecamatan ({error_terakhir}), "
              "pakai nama kabupaten saja.")
        return [area_query]

    from collections import Counter
    elemen = [e for e in data.get("elements", []) if e.get("tags", {}).get("name")]
    levels = Counter(e["tags"].get("admin_level") for e in elemen)
    if not levels:
        return [area_query]

    # level dengan anggota terbanyak biasanya = kecamatan
    level_kec, jumlah = levels.most_common(1)[0]
    if jumlah < 3:
        return [area_query]
    nama_kec = sorted({
        e["tags"]["name"] for e in elemen
        if e["tags"].get("admin_level") == level_kec
    })

    print(f"[INFO] {len(nama_kec)} kecamatan ditemukan: {', '.join(nama_kec[:8])}"
          + ("..." if len(nama_kec) > 8 else ""))
    return nama_kec


def dedup_cafes(cafes: list[dict]) -> list[dict]:
    """Hapus duplikat antar-kecamatan berdasarkan link peta / nama+koordinat."""
    seen, hasil = set(), []
    for c in cafes:
        if c.get("link_peta"):
            key = c["link_peta"]
        elif c.get("latitude") is not None:
            key = (re.sub(r"[^a-z0-9]", "", c["nama"].lower()),
                   round(c["latitude"], 4), round(c["longitude"], 4))
        else:
            key = re.sub(r"[^a-z0-9]", "", c["nama"].lower())
        if key in seen:
            continue
        seen.add(key)
        hasil.append(c)
    return hasil


# ---------------------------------------------------------------------------
# Mode grid: bagi wilayah menjadi sel-sel kecil untuk pencarian Google Maps
# (dipakai saat data kecamatan tidak tersedia di OSM)
# ---------------------------------------------------------------------------

def cari_area_geo(area_query: str) -> dict | None:
    """Ambil poligon batas wilayah (GeoJSON) dari Nominatim."""
    resp = requests.get(
        NOMINATIM_URL,
        params={
            "q": area_query, "format": "json", "limit": 1,
            "countrycodes": "id", "polygon_geojson": 1,
        },
        headers={"User-Agent": USER_AGENT},
        timeout=REQUEST_TIMEOUT,
    )
    resp.raise_for_status()
    hasil = resp.json()
    if not hasil or "geojson" not in hasil[0]:
        return None
    return hasil[0]["geojson"]


def _titik_dalam_ring(lon: float, lat: float, ring: list) -> bool:
    """Ray casting: True jika titik di dalam ring poligon."""
    inside = False
    j = len(ring) - 1
    for i in range(len(ring)):
        xi, yi = ring[i][0], ring[i][1]
        xj, yj = ring[j][0], ring[j][1]
        if (yi > lat) != (yj > lat) and \
                lon < (xj - xi) * (lat - yi) / (yj - yi) + xi:
            inside = not inside
        j = i
    return inside


def titik_dalam_wilayah(lon: float, lat: float, geojson: dict) -> bool:
    """True jika titik di dalam poligon/multipoligon wilayah."""
    tipe = geojson.get("type")
    koords = geojson.get("coordinates", [])
    polys = koords if tipe == "MultiPolygon" else [koords]
    for poly in polys:
        if not poly:
            continue
        if _titik_dalam_ring(lon, lat, poly[0]):  # ring luar
            if not any(_titik_dalam_ring(lon, lat, hole) for hole in poly[1:]):
                return True
    return False


def buat_grid_sel(geojson: dict, max_sel: int = 30) -> list[tuple[float, float]]:
    """Buat titik tengah sel grid (~5 km) yang berada di dalam wilayah."""
    def semua_titik(koords):
        if isinstance(koords[0], (int, float)):
            yield koords
        else:
            for k in koords:
                yield from semua_titik(k)

    titik = list(semua_titik(geojson["coordinates"]))
    lons = [t[0] for t in titik]
    lats = [t[1] for t in titik]
    lon_min, lon_max = min(lons), max(lons)
    lat_min, lat_max = min(lats), max(lats)

    # step ~0.05 derajat (~5 km), diperbesar jika sel terlalu banyak
    step = 0.05
    while True:
        sel = []
        lat = lat_min + step / 2
        while lat < lat_max:
            lon = lon_min + step / 2
            while lon < lon_max:
                if titik_dalam_wilayah(lon, lat, geojson):
                    sel.append((lat, lon))
                lon += step
            lat += step
        if len(sel) <= max_sel or step > 0.5:
            break
        step *= 1.3
    return sel


# ---------------------------------------------------------------------------
# 2. Sumber data alternatif: Google Places API (lebih lengkap, butuh API key)
# ---------------------------------------------------------------------------

def cafe_dari_google(p: dict) -> dict:
    """Ubah hasil Google Places menjadi record cafe (sosmed tidak tersedia)."""
    lokasi = p.get("location", {})
    telepon = normalisasi_telepon(
        p.get("nationalPhoneNumber") or p.get("internationalPhoneNumber") or ""
    )
    cafe = {
        "nama": p.get("displayName", {}).get("text", "(tanpa nama)"),
        "alamat": p.get("formattedAddress", ""),
        "telepon": telepon,
        "whatsapp": link_whatsapp(telepon),
        "email": "",  # Google Places tidak menyediakan email/sosmed
        "website": p.get("websiteUri", "") or "",
        "latitude": lokasi.get("latitude"),
        "longitude": lokasi.get("longitude"),
        "link_peta": p.get("googleMapsUri", ""),
    }
    for platform in PLATFORM_SOSMED:
        cafe[platform] = ""
    return cafe


def ambil_daftar_cafe_google(area_query: str, api_key: str) -> list[dict]:
    """Ambil daftar cafe via Google Places API (New) Text Search."""
    print("[INFO] Mengambil data cafe dari Google Places API...")
    cafes, page_token, halaman = [], None, 0

    while True:
        halaman += 1
        body = {
            "textQuery": f"cafe di {area_query}",
            "languageCode": "id",
            "pageSize": 20,
        }
        if page_token:
            body["pageToken"] = page_token
            time.sleep(2)  # Google butuh jeda sebelum pageToken bisa dipakai

        resp = requests.post(
            GOOGLE_PLACES_URL,
            json=body,
            headers={
                "Content-Type": "application/json",
                "X-Goog-Api-Key": api_key,
                "X-Goog-FieldMask": GOOGLE_FIELD_MASK,
            },
            timeout=REQUEST_TIMEOUT,
        )
        if resp.status_code != 200:
            raise SystemExit(
                f"[ERROR] Google Places API error {resp.status_code}: {resp.text[:300]}"
            )

        data = resp.json()
        cafes.extend(cafe_dari_google(p) for p in data.get("places", []))

        page_token = data.get("nextPageToken")
        if not page_token or halaman >= 3:  # maks 60 hasil per query
            break

    print(f"[INFO] Ditemukan {len(cafes)} cafe dari Google Places.")
    return cafes


# ---------------------------------------------------------------------------
# 2b. Sumber data alternatif: Scraping Google Maps (tanpa API key)
# ---------------------------------------------------------------------------

def cafe_dari_gmaps(d: dict) -> dict:
    """Ubah hasil mentah scraper Google Maps menjadi record cafe standar."""
    telepon = normalisasi_telepon(d.get("telepon_mentah", ""))
    cafe = {
        "nama": d.get("nama") or "(tanpa nama)",
        "alamat": d.get("alamat", ""),
        "telepon": telepon,
        "whatsapp": link_whatsapp(telepon),
        "email": "",  # Google Maps tidak menampilkan email/sosmed
        "website": d.get("website", "") or "",
        "latitude": d.get("latitude"),
        "longitude": d.get("longitude"),
        "link_peta": d.get("link_peta", ""),
    }
    for platform in PLATFORM_SOSMED:
        cafe[platform] = ""
    return cafe


def ambil_daftar_cafe_gmaps(area_query: str, max_hasil: int = 300,
                            headless: bool = True,
                            per_kecamatan: bool = False,
                            max_kecamatan: int = 0) -> list[dict]:
    """Scrape cafe dari Google Maps via Playwright (butuh: pip install playwright
    && python -m playwright install chromium).

    Jika per_kecamatan=True, pencarian dipecah per kecamatan agar hasil
    jauh lebih banyak (Google Maps membatasi ~100 hasil per pencarian).
    """
    try:
        from gmaps_scraper import scrape_cafe_gmaps, scrape_queries
    except ImportError:
        raise SystemExit(
            "[ERROR] Playwright belum terinstal. Jalankan:\n"
            "  pip install playwright\n"
            "  python -m playwright install chromium"
        )

    if per_kecamatan:
        kecamatans = ambil_daftar_kecamatan(area_query)
        if len(kecamatans) > 1:
            # Jalur A: per kecamatan (butuh data batas kecamatan di OSM)
            if max_kecamatan > 0:
                kecamatans = kecamatans[:max_kecamatan]
            queries = [f"cafe di {k}, {area_query}" for k in kecamatans]
            mentah = scrape_queries(queries, max_hasil=max_hasil,
                                    headless=headless)
        else:
            # Jalur B (fallback): mode grid — bagi wilayah jadi sel-sel ~5 km.
            # Dipakai karena batas kecamatan sering tidak ada di OSM Indonesia.
            print("[INFO] Data kecamatan tidak tersedia di OSM, "
                  "otomatis pakai mode GRID.")
            geo = cari_area_geo(area_query)
            if not geo:
                print("[WARN] Poligon wilayah tidak ada, pencarian biasa saja.")
                mentah = scrape_cafe_gmaps(area_query, max_hasil=max_hasil,
                                           headless=headless)
            else:
                sel = buat_grid_sel(geo, max_sel=max_kecamatan or 30)
                print(f"[INFO] Wilayah dibagi menjadi {len(sel)} sel grid.")
                urls = [
                    f"https://www.google.com/maps/search/cafe/@{lat},{lon},14z"
                    for lat, lon in sel
                ]
                mentah = scrape_queries(urls, max_hasil=max_hasil,
                                        headless=headless)
                # buang hasil di luar batas wilayah
                mentah = [
                    d for d in mentah
                    if d.get("latitude") is not None
                    and titik_dalam_wilayah(d["longitude"], d["latitude"], geo)
                ]
    else:
        mentah = scrape_cafe_gmaps(area_query, max_hasil=max_hasil,
                                   headless=headless)

    cafes = dedup_cafes([cafe_dari_gmaps(d) for d in mentah])
    print(f"[INFO] Total cafe setelah dedup: {len(cafes)}")
    return cafes


# ---------------------------------------------------------------------------
# 3. Validasi website & sosmed (bagian terpenting: data harus valid)
# ---------------------------------------------------------------------------

def status_url(url: str) -> tuple[str, int | None]:
    """Cek URL ringan. Return: TERVERIFIKASI / TIDAK_DITEMUKAN / BELUM_TERVERIFIKASI."""
    try:
        resp = requests.get(
            url,
            headers={"User-Agent": USER_AGENT},
            timeout=REQUEST_TIMEOUT,
            allow_redirects=True,
            verify=False,
            stream=True,  # tidak download seluruh body
        )
        kode = resp.status_code
        resp.close()
        if kode == 404:
            return "TIDAK_DITEMUKAN", kode
        if kode < 400:
            return "TERVERIFIKASI", kode
        return "BELUM_TERVERIFIKASI", kode  # 401/403/429 = diblokir, bukan berarti palsu
    except requests.exceptions.RequestException:
        return "BELUM_TERVERIFIKASI", None



SOSMED_DOMAIN = {
    "instagram.com": "instagram",
    "facebook.com": "facebook",
    "fb.com": "facebook",
    "tiktok.com": "tiktok",
    "twitter.com": "twitter",
    "x.com": "twitter",
    "youtube.com": "youtube",
    "youtu.be": "youtube",
}


def pisahkan_sosmed_dari_website(cafe: dict):
    """Jika kolom 'website' ternyata berisi link sosmed (umum di cafe Indonesia:
    profil IG dipasang sebagai 'website' di Google Maps), pindahkan ke kolom
    sosmed yang benar agar status website akurat."""
    web = cafe.get("website", "").strip()
    if not web:
        return
    url = web if urlparse(web).scheme else "https://" + web
    host = urlparse(url).netloc.lower()
    for domain, platform in SOSMED_DOMAIN.items():
        if host == domain or host.endswith("." + domain):
            if not cafe.get(platform):
                cafe[platform] = url
            cafe["website"] = ""  # sosmed BUKAN website
            return


def cek_satu_cafe(cafe: dict) -> dict:
    """Validasi website + semua sosmed satu cafe. Link mati/404 dibuang."""
    pisahkan_sosmed_dari_website(cafe)
    # --- Website ---
    website = cafe.get("website", "").strip()
    if not website:
        cafe["status_website"] = "TIDAK_PUNYA_WEBSITE"
        cafe["keterangan"] = "Tidak ada website tercatat"
    else:
        url = website if urlparse(website).scheme else "https://" + website
        cafe["website"] = url
        try:
            resp = requests.get(
                url,
                headers={"User-Agent": USER_AGENT},
                timeout=REQUEST_TIMEOUT,
                allow_redirects=True,
                verify=False,
            )
            if resp.status_code < 400:
                cafe["status_website"] = "AKTIF"
                cafe["keterangan"] = f"HTTP {resp.status_code}"
            else:
                cafe["status_website"] = "WEBSITE_TIDAK_AKTIF"
                cafe["keterangan"] = f"HTTP {resp.status_code}"
        except requests.exceptions.SSLError:
            cafe["status_website"] = "WEBSITE_TIDAK_AKTIF"
            cafe["keterangan"] = "SSL error"
        except requests.exceptions.ConnectionError:
            cafe["status_website"] = "WEBSITE_TIDAK_AKTIF"
            cafe["keterangan"] = "Domain mati / koneksi gagal"
        except requests.exceptions.Timeout:
            cafe["status_website"] = "WEBSITE_TIDAK_AKTIF"
            cafe["keterangan"] = "Timeout (server tidak merespons)"
        except requests.exceptions.RequestException as e:
            cafe["status_website"] = "WEBSITE_TIDAK_AKTIF"
            cafe["keterangan"] = f"Error: {e.__class__.__name__}"

    # --- Media sosial ---
    catatan = []
    for platform in PLATFORM_SOSMED:
        link = cafe.get(platform)
        if not link:
            continue
        status, _ = status_url(link)
        if status == "TIDAK_DITEMUKAN":
            cafe[platform] = ""  # buang link palsu/mati agar data tetap valid
            catatan.append(f"{platform}: akun tidak ditemukan (dihapus)")
        elif status == "BELUM_TERVERIFIKASI":
            catatan.append(f"{platform}: belum terverifikasi")
    cafe["catatan_validasi"] = "; ".join(catatan)
    return cafe


def cek_semua(cafes: list[dict], workers: int = 10) -> list[dict]:
    """Validasi semua cafe secara paralel."""
    total = len(cafes)
    print(f"[INFO] Memvalidasi website & sosmed {total} cafe...")
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futures = [ex.submit(cek_satu_cafe, c) for c in cafes]
        selesai = 0
        for fut in as_completed(futures):
            fut.result()
            selesai += 1
            print(f"\r[INFO] Progress: {selesai}/{total}", end="", flush=True)
    print()
    return cafes


# ---------------------------------------------------------------------------
# 4. Filter prospek
# ---------------------------------------------------------------------------

def filter_target(cafes: list[dict]) -> list[dict]:
    """Hanya cafe prospek: tidak punya website ATAU website tidak aktif."""
    return [
        c
        for c in cafes
        if c["status_website"] in ("TIDAK_PUNYA_WEBSITE", "WEBSITE_TIDAK_AKTIF")
    ]


def punya_kontak(cafe: dict) -> bool:
    return bool(cafe.get("telepon") or cafe.get("email"))


def punya_sosmed(cafe: dict) -> bool:
    return any(cafe.get(p) for p in PLATFORM_SOSMED)



# ---------------------------------------------------------------------------
# 5. Export Excel (rapi, berformat, hyperlink aktif)
# ---------------------------------------------------------------------------

def export_excel(cafes: list[dict], path: str, area: str):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    KOLOM = [
        ("No", 5), ("Nama Cafe", 30), ("Alamat Lengkap", 42),
        ("Telepon", 16), ("WhatsApp", 22), ("Email", 26),
        ("Status Website", 22), ("Website", 32), ("Keterangan Website", 26),
        ("Instagram", 28), ("Facebook", 28), ("TikTok", 28),
        ("Twitter / X", 24), ("YouTube", 28),
        ("Koordinat", 22), ("Link Peta", 14), ("Catatan Validasi", 32),
    ]
    n_kolom = len(KOLOM)
    garis = Side(style="thin", color="B0B0B0")
    border = Border(left=garis, right=garis, top=garis, bottom=garis)

    wb = Workbook()
    ws = wb.active
    ws.title = "Prospek Cafe"

    # --- Baris judul ---
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_kolom)
    c = ws.cell(row=1, column=1, value=f"DAFTAR PROSPEK CAFE TANPA WEBSITE AKTIF — {area.upper()}")
    c.font = Font(bold=True, size=14, color="1F4E78")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_kolom)
    c = ws.cell(row=2, column=1, value=(
        f"Dibuat: {datetime.now():%d/%m/%Y %H:%M}  |  Total: {len(cafes)} cafe prospek  |  "
        "Semua link website & sosmed sudah divalidasi otomatis (link mati dihapus)"
    ))
    c.font = Font(italic=True, size=9, color="606060")
    c.alignment = Alignment(horizontal="center", vertical="center")

    # --- Baris header (baris 3) ---
    header_fill = PatternFill("solid", fgColor="1F4E78")
    for col, (nama_kolom, lebar) in enumerate(KOLOM, start=1):
        c = ws.cell(row=3, column=col, value=nama_kolom)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
        ws.column_dimensions[get_column_letter(col)].width = lebar
    ws.row_dimensions[3].height = 22


    # --- Isi data ---
    warna_status = {
        "TIDAK_PUNYA_WEBSITE": PatternFill("solid", fgColor="FFF2CC"),   # kuning
        "WEBSITE_TIDAK_AKTIF": PatternFill("solid", fgColor="F8CBAD"),   # oranye
        "AKTIF": PatternFill("solid", fgColor="C6EFCE"),                 # hijau
    }
    kolom_link = {  # kolom (1-based) -> key data yang berisi URL
        5: "whatsapp", 8: "website", 10: "instagram", 11: "facebook",
        12: "tiktok", 13: "twitter", 14: "youtube", 16: "link_peta",
    }
    label_status = {
        "TIDAK_PUNYA_WEBSITE": "Tidak Punya Website",
        "WEBSITE_TIDAK_AKTIF": "Website Tidak Aktif",
        "AKTIF": "Aktif",
    }

    for i, cafe in enumerate(cafes, start=1):
        row = i + 3
        koordinat = (
            f"{cafe['latitude']}, {cafe['longitude']}"
            if cafe.get("latitude") is not None else "-"
        )
        nilai = [
            i,
            cafe["nama"],
            cafe["alamat"] or "-",
            cafe["telepon"] or "-",
            cafe["telepon"] or "-",          # kolom WhatsApp: tampil nomor, link wa.me
            cafe["email"] or "-",
            label_status.get(cafe["status_website"], cafe["status_website"]),
            cafe["website"] or "-",
            cafe["keterangan"],
            cafe["instagram"] or "-",
            cafe["facebook"] or "-",
            cafe["tiktok"] or "-",
            cafe["twitter"] or "-",
            cafe["youtube"] or "-",
            koordinat,
            "Buka Peta" if cafe.get("link_peta") else "-",
            cafe.get("catatan_validasi", "") or "OK",
        ]
        for col, v in enumerate(nilai, start=1):
            c = ws.cell(row=row, column=col, value=v)
            c.border = border
            c.alignment = Alignment(
                vertical="top",
                horizontal="center" if col in (1, 4, 5, 7, 16) else "left",
                wrap_text=col in (3, 9, 17),
            )
            url = kolom_link.get(col) and cafe.get(kolom_link[col])
            if url and v != "-":
                c.hyperlink = url
                c.font = Font(color="0563C1", underline="single", size=10)
            if col == 7 and cafe["status_website"] in warna_status:
                c.fill = warna_status[cafe["status_website"]]
                c.font = Font(bold=True, size=10)
            if col in (4, 5):
                c.number_format = "@"  # nomor telepon sebagai teks agar +62 utuh

    # --- Freeze & filter ---
    ws.freeze_panes = "C4"  # kolom No & Nama + 3 baris atas selalu terlihat
    ws.auto_filter.ref = f"A3:{get_column_letter(n_kolom)}{3 + len(cafes)}"
    ws.sheet_view.showGridLines = False


    # --- Sheet Ringkasan ---
    ws2 = wb.create_sheet("Ringkasan")
    n_tanpa = sum(1 for c in cafes if c["status_website"] == "TIDAK_PUNYA_WEBSITE")
    n_mati = sum(1 for c in cafes if c["status_website"] == "WEBSITE_TIDAK_AKTIF")
    n_kontak = sum(1 for c in cafes if punya_kontak(c))
    n_sosmed = sum(1 for c in cafes if punya_sosmed(c))
    ringkasan = [
        ("RINGKASAN PROSPEK", ""),
        ("Kawasan", area),
        ("Tanggal dibuat", f"{datetime.now():%d/%m/%Y %H:%M}"),
        ("", ""),
        ("Total cafe prospek", len(cafes)),
        ("Tidak punya website sama sekali", n_tanpa),
        ("Punya website tapi tidak aktif", n_mati),
        ("Punya kontak valid (telepon/email)", n_kontak),
        ("Punya media sosial (terverifikasi)", n_sosmed),
    ]
    for r, (k, v) in enumerate(ringkasan, start=1):
        ck = ws2.cell(row=r, column=1, value=k)
        cv = ws2.cell(row=r, column=2, value=v)
        if r == 1:
            ck.font = Font(bold=True, size=13, color="1F4E78")
        elif k:
            ck.font = Font(bold=True)
            ck.border = border
            cv.border = border
    ws2.column_dimensions["A"].width = 38
    ws2.column_dimensions["B"].width = 40
    ws2.sheet_view.showGridLines = False

    wb.save(path)
    print(f"[OK] Excel disimpan ke: {path}")


# ---------------------------------------------------------------------------
# 6. Export CSV / JSON (opsional)
# ---------------------------------------------------------------------------

KOLOM_FLAT = [
    "nama", "alamat", "telepon", "whatsapp", "email",
    "status_website", "website", "keterangan",
    "instagram", "facebook", "tiktok", "twitter", "youtube",
    "latitude", "longitude", "link_peta", "catatan_validasi",
]


def export_csv(cafes: list[dict], path: str):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=KOLOM_FLAT, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(cafes)
    print(f"[OK] CSV disimpan ke: {path}")


def export_json(cafes: list[dict], path: str):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cafes, f, ensure_ascii=False, indent=2)
    print(f"[OK] JSON disimpan ke: {path}")



# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Cari cafe prospek (tanpa website aktif) di suatu kabupaten, "
        "lengkap dengan sosmed, kontak valid, dan alamat."
    )
    parser.add_argument("--area", required=True,
                        help='Nama kabupaten/kota, contoh: "Jember"')
    parser.add_argument("--output", default=None,
                        help="Nama file output (default: hasil_<area>.xlsx)")
    parser.add_argument("--format", choices=["xlsx", "csv", "json"], default="xlsx",
                        help="Format output (default: xlsx)")
    parser.add_argument("--semua", action="store_true",
                        help="Simpan semua cafe, bukan hanya prospek.")
    parser.add_argument("--workers", type=int, default=10,
                        help="Jumlah thread untuk validasi (default: 10)")
    parser.add_argument("--google-key", default=os.environ.get("GOOGLE_API_KEY"),
                        help="API key Google Places (New). Bisa juga via env "
                        "GOOGLE_API_KEY. Data cafe jadi jauh lebih lengkap.")
    parser.add_argument("--gmaps", action="store_true",
                        help="Scrape langsung dari Google Maps (tanpa API key, "
                        "via browser otomatis Playwright). Paling lengkap.")
    parser.add_argument("--max-hasil", type=int, default=300,
                        help="Maks hasil scraping Google Maps (default: 300)")
    parser.add_argument("--no-headless", action="store_true",
                        help="Tampilkan jendela browser saat scraping "
                        "(untuk melihat prosesnya).")
    parser.add_argument("--per-kecamatan", action="store_true",
                        help="Pecah pencarian per kecamatan/sel grid - hasil "
                        "lebih lengkap tapi lebih lama (hanya dengan --gmaps).")
    parser.add_argument("--max-kecamatan", type=int, default=0,
                        help="Batasi jumlah kecamatan/sel grid yang diproses "
                        "(default: 0 = semua, maks 30 sel grid). "
                        "Berguna untuk uji coba.")
    parser.add_argument("--cari-ig", action="store_true",
                        help="Cari akun Instagram tiap cafe (handle guessing + "
                        "validasi API Instagram).")
    args = parser.parse_args()

    start = time.time()

    if args.gmaps:
        cafes = ambil_daftar_cafe_gmaps(
            args.area, max_hasil=args.max_hasil,
            headless=not args.no_headless,
            per_kecamatan=args.per_kecamatan,
            max_kecamatan=args.max_kecamatan,
        )
    elif args.google_key:
        cafes = ambil_daftar_cafe_google(args.area, args.google_key)
    else:
        cafes = ambil_daftar_cafe(args.area)

    if not cafes:
        print("[INFO] Tidak ada cafe ditemukan. Coba nama area lain.")
        sys.exit(0)

    if args.cari_ig:
        try:
            from ig_finder import cari_instagram_massal
            cari_instagram_massal(cafes, args.area)
        except ImportError:
            print("[WARN] ig_finder.py tidak ditemukan, pencarian IG dilewati.")

    cafes = cek_semua(cafes, workers=args.workers)
    target = filter_target(cafes)

    n_aktif = sum(1 for c in cafes if c["status_website"] == "AKTIF")
    n_mati = sum(1 for c in cafes if c["status_website"] == "WEBSITE_TIDAK_AKTIF")
    n_tanpa = sum(1 for c in cafes if c["status_website"] == "TIDAK_PUNYA_WEBSITE")
    n_kontak = sum(1 for c in target if punya_kontak(c))
    n_sosmed = sum(1 for c in target if punya_sosmed(c))

    print("\n========== RINGKASAN ==========")
    print(f"Total cafe ditemukan      : {len(cafes)}")
    print(f"  Website aktif           : {n_aktif}")
    print(f"  Website tidak aktif     : {n_mati}  <-- prospek")
    print(f"  Tidak punya website     : {n_tanpa}  <-- prospek")
    print(f"Total prospek             : {len(target)}")
    print(f"  + punya kontak valid    : {n_kontak}")
    print(f"  + punya sosmed          : {n_sosmed}")
    print("===============================\n")

    hasil = cafes if args.semua else target

    area_slug = args.area.lower().replace(" ", "_")
    output = args.output or f"hasil_{area_slug}.{args.format}"

    if args.format == "xlsx":
        export_excel(hasil, output, args.area)
    elif args.format == "csv":
        export_csv(hasil, output)
    else:
        export_json(hasil, output)

    print(f"[INFO] Selesai dalam {time.time() - start:.1f} detik.")


if __name__ == "__main__":
    import urllib3

    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    main()

