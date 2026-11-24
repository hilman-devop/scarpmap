# ScarpMap — Local Business Lead Generation

![Python](https://img.shields.io/badge/Python-3776AB?style=flat&logo=python&logoColor=white) ![Playwright](https://img.shields.io/badge/Playwright-2EAD33?style=flat&logo=playwright&logoColor=white) ![OpenPyXL](https://img.shields.io/badge/Excel-217346?style=flat&logo=microsoft-excel&logoColor=white)

> Tool Python untuk menemukan bisnis lokal yang belum memiliki website dan mengekspor prospek ke Excel.

## Tentang proyek

ScarpMap membantu riset prospek untuk agensi web: mencari bisnis berdasarkan area, memeriksa status website, menormalisasi kontak, dan menghasilkan data yang lebih siap ditindaklanjuti.

## Fitur utama

- Pencarian berbasis grid per kecamatan
- Validasi website hidup atau mati
- Prediksi dan validasi handle Instagram
- Normalisasi nomor WhatsApp ke format internasional
- Ekspor prospek ke Excel
- Modul terpisah untuk scraping, ranking, aset, dan pembuatan demo

Prospect exports, downloaded business photos and generated demos are local-only and excluded from Git. Use verified, owner-approved material when creating a demo.

## Teknologi

- Python
- Playwright
- Google Places
- OpenStreetMap
- openpyxl

## Menjalankan secara lokal

```bash
pip install -r requirements.txt
python scrapmap.py --area "Kota Anda"
```

## Keamanan dan etika penggunaan

Gunakan API, scraper, dan data kontak dengan mematuhi terms of service, rate limit, privasi, dan hukum yang berlaku. Jangan commit credential API atau data pribadi hasil riset ke repository.

## Status

Proyek ini merupakan bagian dari portofolio **Hilman Arifin Ilham** dan dapat dikembangkan sesuai kebutuhan penggunaan.

## Lisensi

Belum ada lisensi open-source yang dinyatakan di repository ini. Hubungi pemilik repository sebelum menggunakan ulang atau mendistribusikan kode.

---

Dibuat oleh [Hilman Arifin Ilham](https://github.com/hilman-devop).
