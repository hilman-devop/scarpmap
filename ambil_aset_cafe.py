"""Ambil aset visual asli tiap cafe prospek: foto Google Maps + logo IG.

Untuk N cafe teratas (sesuai urutan rank_cafe.py), script ini:
  1. Membuka halaman Google Maps cafe (Playwright) dan mengumpulkan URL
     foto lh*.googleusercontent.com di panel detail, lalu mendownload
     sampai MAX_FOTO foto (resolusi dinaikkan ke w1200).
  2. Mencoba mengambil foto profil Instagram (logo) lewat endpoint
     web_profile_info tanpa login. Jika gagal -> generator demo akan
     membuat monogram tipografis sebagai pengganti.

Output: aset_cafe/<slug>/img01.jpg..., logo.jpg (opsional), meta.json

PERINGATAN: scraping Google Maps melanggar ToS Google; pakai sewajarnya.
"""

import json
import re
import sys
import time
from pathlib import Path

import openpyxl
import requests
from playwright.sync_api import sync_playwright

from rank_cafe import (ambil_wa, skor_cafe,
                       KOL_NAMA, KOL_WA, KOL_IG, KOL_CATATAN)

KOL_PETA = 16
MAX_FOTO = 6
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")


def slugify(teks: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", teks.lower()).strip("-")[:40] or "cafe"


def username_ig(url: str) -> str | None:
    m = re.search(r"instagram\.com/([A-Za-z0-9_.]+)", url or "")
    if m and m.group(1) not in ("p", "reel", "reels", "stories"):
        return m.group(1)
    return None


def ambil_logo_ig(user: str) -> bytes | None:
    try:
        r = requests.get(
            "https://www.instagram.com/api/v1/users/web_profile_info/",
            params={"username": user},
            headers={"X-IG-App-ID": "936619743392459", "User-Agent": UA},
            timeout=15)
        if r.status_code != 200:
            return None
        pic = (r.json().get("data", {}).get("user", {})
               .get("profile_pic_url_hd"))
        if not pic:
            return None
        img = requests.get(pic, headers={"User-Agent": UA}, timeout=20)
        return img.content if img.status_code == 200 else None
    except Exception:
        return None


def url_foto_besar(src: str) -> str:
    # ...=w203-h114-k-no -> naikkan ke w1200; tanpa suffix -> tambahkan
    if src.startswith("//"):
        src = "https:" + src
    if re.search(r"=w\d+", src):
        return re.sub(r"=w\d+[^/]*$", "=w1200-h900-k-no", src)
    return src + "=w1200-h900-k-no"


def _url_foto_valid(src: str) -> bool:
    """Foto tempat = path /p/, /gps-cs, /grass-cs; bukan avatar /a/ w36."""
    if "googleusercontent" not in src:
        return False
    if "/a/" in src or "/a-/" in src:  # avatar reviewer
        return False
    return any(p in src for p in ("/p/", "/gps-cs", "/grass-cs"))


def _kumpulkan_url_foto(page) -> list[str]:
    srcs = page.eval_on_selector_all("img", "els => els.map(e => e.src)")
    bgs = page.eval_on_selector_all(
        "[style*='background-image']",
        "els => els.map(e => e.style.backgroundImage)")
    for b in bgs:
        m = re.search(r'url\("?([^")]+)"?\)', b)
        if m:
            srcs.append(m.group(1))
    out, seen = [], set()
    for s in srcs:
        if not _url_foto_valid(s):
            continue
        dasar = s.split("=")[0]
        if dasar in seen:
            continue
        seen.add(dasar)
        out.append(url_foto_besar(s))
    return out


def ambil_foto_maps(page, url: str) -> list[str]:
    page.goto(url, timeout=45000)
    page.wait_for_selector("h1", timeout=20000)
    time.sleep(2.5)
    for _ in range(3):
        page.mouse.wheel(0, 800)
        time.sleep(0.6)
    out = _kumpulkan_url_foto(page)

    # Kalau foto sedikit, buka tab "Foto" untuk galeri penuh
    if len(out) < 4:
        try:
            for btn in page.query_selector_all("button[role='tab'], button"):
                label = (btn.get_attribute("aria-label") or
                         btn.inner_text() or "").lower()
                if "foto" in label or "photo" in label:
                    btn.click()
                    time.sleep(2.5)
                    for _ in range(4):
                        page.mouse.wheel(0, 1200)
                        time.sleep(0.5)
                    out = _kumpulkan_url_foto(page)
                    break
        except Exception:
            pass
    return out[:MAX_FOTO]


def download(url: str, path: Path) -> bool:
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=30)
        if r.status_code == 200 and len(r.content) > 5000:
            path.write_bytes(r.content)
            return True
    except Exception:
        pass
    return False


def main() -> None:
    path = sys.argv[1] if len(sys.argv) > 1 else "hasil_jember.xlsx"
    jumlah = int(sys.argv[2]) if len(sys.argv) > 2 else 10
    filter_slug = sys.argv[3] if len(sys.argv) > 3 else None

    wb = openpyxl.load_workbook(path)
    ws = wb["Prospek Cafe"]

    kandidat = []
    for r in range(4, ws.max_row + 1):
        nama = ws.cell(r, KOL_NAMA).value
        if not nama:
            continue
        nama = str(nama).strip()
        wa = ambil_wa(ws.cell(r, KOL_WA))
        ig = str(ws.cell(r, KOL_IG).value or "").strip()
        ig = "" if ig == "-" else ig
        catatan = str(ws.cell(r, KOL_CATATAN).value or "")
        skor, _ = skor_cafe(nama, wa, ig, catatan)
        # filter_slug = ambil cafe spesifik, abaikan batas skor
        if not filter_slug and skor < 60:
            continue
        peta = ws.cell(r, KOL_PETA)
        link = (peta.hyperlink.target if peta.hyperlink else peta.value) or ""
        kandidat.append({"nama": nama, "skor": skor, "ig": ig,
                         "maps": str(link)})
    kandidat.sort(key=lambda k: -k["skor"])
    if filter_slug:
        # dukung beberapa slug dipisah koma, mis. "perasa,navas"
        inc = [s.strip() for s in filter_slug.split(",") if s.strip()]
        kandidat = [k for k in kandidat
                    if any(s in slugify(k["nama"]) for s in inc)]
    else:
        kandidat = kandidat[:jumlah]

    print(f"[INFO] Mengambil aset untuk {len(kandidat)} cafe teratas...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(user_agent=UA, locale="id-ID",
                                  viewport={"width": 1366, "height": 900})
        page = ctx.new_page()
        for i, k in enumerate(kandidat, 1):
            slug = slugify(k["nama"])
            folder = Path("aset_cafe") / slug
            folder.mkdir(parents=True, exist_ok=True)
            print(f"[{i}/{len(kandidat)}] {k['nama']}")

            foto = []
            if k["maps"].startswith("http"):
                try:
                    foto = ambil_foto_maps(page, k["maps"])
                except Exception as e:
                    print(f"  [WARN] Maps gagal: {e.__class__.__name__}")
            n_ok = 0
            for j, u in enumerate(foto, 1):
                if download(u, folder / f"img{j:02d}.jpg"):
                    n_ok += 1
            print(f"  foto: {n_ok}/{len(foto)} terunduh")

            logo_ok = False
            user = username_ig(k["ig"])
            if user:
                data = ambil_logo_ig(user)
                if data:
                    (folder / "logo.jpg").write_bytes(data)
                    logo_ok = True
            print(f"  logo IG ({user or '-'}): {'OK' if logo_ok else 'gagal'}")

            (folder / "meta.json").write_text(json.dumps({
                "nama": k["nama"], "skor": k["skor"], "ig": k["ig"],
                "maps": k["maps"], "foto": n_ok, "logo": logo_ok,
            }, ensure_ascii=False, indent=2), encoding="utf-8")
            time.sleep(1.0)
        browser.close()
    print("[OK] Selesai. Aset tersimpan di aset_cafe/")


if __name__ == "__main__":
    main()
