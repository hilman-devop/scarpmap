"""
Scraper Google Maps untuk ScrapMap (tanpa API key).

Menggunakan Playwright (browser Chromium otomatis) untuk:
  1. Mencari query (mis. "kafe di kota Anda") di Google Maps
  2. Scroll panel hasil sampai semua listing termuat
  3. Membuka tiap listing untuk mengambil nama, alamat, telepon, website,
     dan koordinat

Mendukung banyak query dalam SATU sesi browser (untuk mode per-kecamatan).

PERINGATAN: Scraping Google Maps melanggar Terms of Service Google.
Gunakan sewajarnya agar IP tidak kena rate-limit sementara.
Selector DOM Google Maps bisa berubah sewaktu-waktu.
"""

import re
import time
from urllib.parse import quote_plus

from playwright.sync_api import sync_playwright

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)


def _teks(page, selector: str) -> str:
    """Ambil inner_text elemen pertama yang cocok; '' jika tidak ada."""
    try:
        el = page.query_selector(selector)
        if el:
            return el.inner_text().strip()
    except Exception:
        pass
    return ""


def _attr(page, selector: str, attr: str) -> str:
    try:
        el = page.query_selector(selector)
        if el:
            return (el.get_attribute(attr) or "").strip()
    except Exception:
        pass
    return ""


def _handle_consent(page):
    """Klik tombol setuju jika dialihkan ke halaman consent Google."""
    if "consent.google" not in page.url:
        return
    try:
        for b in page.query_selector_all("button"):
            t = (b.inner_text() or "").lower()
            if any(k in t for k in ("accept all", "terima semua", "setuju", "i agree")):
                b.click()
                page.wait_for_load_state("domcontentloaded")
                time.sleep(2)
                return
    except Exception:
        pass


def _kumpulkan_link_hasil(page, max_hasil: int) -> list[str]:
    """Scroll panel hasil sampai mentok, kembalikan daftar URL tiap tempat."""
    try:
        page.wait_for_selector('div[role="feed"]', timeout=20000)
    except Exception:
        return []  # query ini tidak menghasilkan apa-apa
    feed = page.query_selector('div[role="feed"]')

    jumlah_sebelumnya, tidak_berubah = -1, 0
    while True:
        feed.evaluate("el => el.scrollTo(0, el.scrollHeight)")
        time.sleep(1.2)
        n = len(page.query_selector_all("a.hfpxzc"))
        print(f"\r[INFO] Memuat daftar... {n} tempat", end="", flush=True)
        if n == jumlah_sebelumnya:
            tidak_berubah += 1
        else:
            tidak_berubah = 0
        jumlah_sebelumnya = n
        if tidak_berubah >= 3 or n >= max_hasil:
            break
    print()

    hrefs = page.eval_on_selector_all("a.hfpxzc", "els => els.map(e => e.href)")
    seen, hasil = set(), []
    for h in hrefs:
        if h not in seen:
            seen.add(h)
            hasil.append(h)
    return hasil[:max_hasil]


def _parse_detail(page, url: str) -> dict:
    """Buka halaman detail tempat dan ekstrak informasinya."""
    page.goto(url, timeout=45000)
    page.wait_for_selector("h1", timeout=20000)

    nama = _teks(page, "h1.DUwDvf") or _teks(page, "h1")
    alamat = _teks(page, 'button[data-item-id="address"] .Io6YTe')
    if not alamat:
        alamat = _attr(page, 'button[data-item-id="address"]', "aria-label")
        alamat = re.sub(r"^Alamat:\s*|^Address:\s*", "", alamat)
    telepon = _teks(page, 'button[data-item-id^="phone"] .Io6YTe')
    website = _attr(page, 'a[data-item-id="authority"]', "href")

    lat = lon = None
    # format 1: URL tampilan peta .../@-8.17,113.71,17z
    m = re.search(r"@(-?\d+\.\d+),(-?\d+\.\d+)", page.url)
    if not m:
        # format 2: URL detail tempat ...!3d-8.17!4d113.71...
        m = re.search(r"!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)", page.url)
    if m:
        lat, lon = float(m.group(1)), float(m.group(2))

    return {
        "nama": nama,
        "alamat": alamat,
        "telepon_mentah": telepon,
        "website": website,
        "latitude": lat,
        "longitude": lon,
        "link_peta": url.split("?")[0],
    }


def scrape_queries(queries: list[str], max_hasil: int = 300,
                   headless: bool = True, delay: float = 0.7) -> list[dict]:
    """Scrape beberapa query Google Maps dalam SATU sesi browser.

    Return: list[dict] mentah dengan key: nama, alamat, telepon_mentah,
    website, latitude, longitude, link_peta.
    """
    hasil = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        ctx = browser.new_context(
            user_agent=USER_AGENT,
            locale="id-ID",
            viewport={"width": 1366, "height": 900},
        )
        page = ctx.new_page()

        for qi, query in enumerate(queries, start=1):
            if query.startswith("http"):
                url = query  # URL langsung (mis. sel grid berkoordinat)
                label = f"sel {qi}"
            else:
                url = f"https://www.google.com/maps/search/{quote_plus(query)}?hl=id"
                label = f"'{query}'"
            print(f"[INFO] [{qi}/{len(queries)}] Pencarian: {label}")
            try:
                page.goto(url, timeout=60000)
                _handle_consent(page)
                links = _kumpulkan_link_hasil(page, max_hasil)
            except Exception as e:
                print(f"[WARN] Query {label} gagal: {e.__class__.__name__}")
                continue

            print(f"[INFO] {len(links)} tempat ditemukan. Mengambil detail...")
            for i, link in enumerate(links, start=1):
                try:
                    detail = _parse_detail(page, link)
                    if detail["nama"]:
                        hasil.append(detail)
                    print(
                        f"\r[INFO] Detail: {i}/{len(links)} - {detail['nama'][:40]:<40}",
                        end="", flush=True,
                    )
                except Exception as e:
                    print(f"\n[WARN] Gagal baca detail ke-{i}: {e.__class__.__name__}")
                time.sleep(delay)  # sopan: jangan hajar server Google
            print()

        browser.close()

    print(f"[INFO] Total mentah terkumpul: {len(hasil)} tempat.")
    return hasil


def scrape_cafe_gmaps(area: str, max_hasil: int = 300, headless: bool = True,
                      delay: float = 0.7) -> list[dict]:
    """Scrape daftar cafe di Google Maps untuk satu area."""
    return scrape_queries([f"cafe di {area}"], max_hasil=max_hasil,
                          headless=headless, delay=delay)

