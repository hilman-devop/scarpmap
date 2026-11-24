# Template Chat WhatsApp — Penawaran Jasa Website untuk Cafe

Target: daftar prospek dari ScrapMap (cafe tanpa website, punya nomor WA valid).

## Strategi Persuasi (baca dulu sebelum kirim!)

Ada 2 pintu yang harus dibuka: **NEED** ("cafe saya butuh website?") lalu
**TRUST** ("orang ini bisa dipercaya?"). Template di bawah dirancang untuk
membuka keduanya.

### Membangun NEED (yakin butuh website)

1. **Jual kerugian, bukan keuntungan** — "calon pelanggan yang googling malah
   mampir ke cafe sebelah karena info cafe Kakak tidak lengkap".
2. **Skenario konkret lokal** — wilayah dengan banyak mahasiswa; setiap tahun banyak
   mahasiswa baru mencari cafe lewat Google. Yang menang = yang punya
   menu + foto + jam buka yang bisa dicari.
3. **ROI bahasa warung** — "2 pelanggan tambahan/hari × Rp25rb = Rp1,5jt/bulan.
   Website balik modal dalam hitungan minggu."
4. **Antisipasi mitos "sudah ada IG"** — IG itu numpang (bisa kena suspend,
   jangkauan diatur algoritma); website = aset milik sendiri yang muncul di
   pencarian Google. Bio IG muat 1 link → taruh website di sana.
5. **Fitur yang terasa harian** — menu QR di meja (hemat cetak ulang tiap
   ganti harga), tombol reservasi WA. Bukan "company profile" yang abstrak.

### Membangun TRUST (percaya pada kamu)

1. **Jurus terkuat: DEMO DULU BARU NAWARIN.** Buat contoh website pakai
   nama & foto cafe mereka (ambil dari Google Maps/IG), host gratis di
   Netlify/Vercel, kirim link-nya. 1 template + ganti konten = ±15 menit
   per cafe. Resiprositas (sungkan karena sudah dikasih) menaikkan balasan
   5-10x dibanding tawaran dingin.
2. **Bukti sosial sekecil apa pun** — 1 klien pertama (walau gratis/diskon)
   = portfolio + screenshot testimoni. "Cafe X sudah pakai" > klaim apa pun.
3. **Nol risiko untuk mereka** — "bayar setelah jadi & Kakak suka" / DP 30% /
   garansi revisi. Berani karena biaya marginal pembuatan hampir nol.
4. **Orang lokal yang bisa ditemui** — foto profil asli, tawarkan kopdar di
   cafe mereka ("sekalian saya foto menunya gratis"). Kamu pun jadi
   pelanggan mereka hari itu.
5. **Jangan oversell** — akui keterbatasan dengan jujur; justru menaikkan
   kepercayaan.

## Template 0 — Demo-First (PALING KUAT, butuh siapkan demo dulu)

```
Halo Kak, selamat siang 🙏

Saya {nama_anda} dari {kota_anda}. Kemarin saya lihat *{nama_cafe}*
di Google Maps — tempatnya bagus, tapi sayang belum punya website,
jadi calon pelanggan yang googling menunya malah mampir ke cafe lain.

Saya iseng buatkan contoh website untuk {nama_cafe}, ini linknya:
👉 {link_demo}

Isinya menu, foto, tombol pesan via WhatsApp, dan arah Google Maps.
Silakan dilihat dulu Kak — gratis, linknya aktif seminggu.

Kalau Kakak suka dan mau pakai, kita ngobrol santai. Kalau tidak,
tidak apa-apa sama sekali 😊
```

## Template 1 — Pembuka Sopan (tanpa demo)

```
Halo Kak, selamat pagi/siang 🙏

Perkenalkan, saya {nama_anda} dari {kota_anda}. Saya menemukan
*{nama_cafe}* di Google Maps dan salut dengan cafenya!

Sekalian izin menawarkan: saya perhatikan {nama_cafe} belum punya
website resmi. Sayang sekali — tiap hari ada calon pelanggan yang
googling menu & jam buka kafe di kota Anda, lalu mampir ke cafe lain
yang infonya lebih lengkap.

Saya bisa bantu buatkan website sederhana khusus cafe, isinya:
• Daftar menu & foto-foto cafe
• Lokasi + tombol arah langsung ke Google Maps
• Tombol "Pesan via WhatsApp" (langsung ke nomor Kakak)
• Menu QR untuk ditaruh di meja

Hitung-hitungannya sederhana Kak: kalau website hanya membawa
2 pelanggan tambahan per hari, biayanya balik modal dalam hitungan
minggu — sisanya keuntungan terus.

Boleh saya kirimkan contoh websitenya dulu? Gratis, tanpa komitmen
apa pun. Kalau nanti jadi pakai, pembayarannya pun setelah website
selesai dan Kakak cocok. 🙏

Terima kasih, Kak. Ditunggu kabarnya ☕
```

## Template 2 — Versi Singkat (jika sudah sempat chat)

```
Halo Kak 🙏 Saya {nama_anda}, jasa pembuatan website khusus cafe.

Saya lihat *{nama_cafe}* belum punya website — saya bisa buatkan
yang simpel & terjangkau (menu, foto, tombol WA, arah Maps).

Boleh saya kirim contohnya dulu? Gratis kok Kak, tanpa komitmen ☕
```

## Template 3 — Follow-Up (kirim 2-3 hari setelah template 1, jika belum dibalas)

```
Halo lagi Kak 🙏 Kemarin saya sempat menawarkan pembuatan website
untuk *{nama_cafe}*.

Barangkali Kakak sibuk — tidak apa-apa. Kalau nanti tertarik,
contoh websitenya masih bisa saya kirim kapan saja, gratis.

Terima kasih, Kak. Sukses terus untuk {nama_cafe}! ☕
```

## Placeholder

| Placeholder | Isi |
|---|---|
| `{nama_cafe}` | Nama cafe dari kolom Excel (diisi otomatis oleh `buat_blast_wa.py`) |
| `{nama_anda}` | Nama kamu |
| `{kota_anda}` | Kota/domisili kamu |

## Tips Penting (baca sebelum kirim!)

1. **Jangan blast massal sekaligus.** Kirim manual 5-10 chat per jam dengan
   jeda acak. WhatsApp bisa mem-banned nomor yang mengirim puluhan pesan
   identik ke nomor asing dalam waktu singkat.
2. **Isi `nama_anda` & `kota_anda` dulu** di `buat_blast_wa.py` sebelum
   generate — pesan yang terasa personal dibalas 3-5x lebih sering.
3. **Jam terbaik kirim:** 10.00-12.00 atau 15.00-17.00 (cafe biasanya
   tidak sibuk). Hindari jam buka tutup & weekend malam.
4. **Sapa sesuai jam:** ganti "pagi/siang" sesuai waktu kirim.
5. **Siapkan contoh/portfolio** sebelum blast — pertanyaan pertama pasti
   "bisa lihat contohnya?". Satu demo website cafe dummy sudah cukup.
6. **Yang tidak dibalas ≠ rugi.** Jangan kirim follow-up lebih dari 1x.
7. Nomor berawal `+62812/813/821/822/823/851/852/853/857/858/895/896/897/898/899`
   = seluler, aman untuk WA. Nomor lain (telepon rumah) tidak bisa di-WA —
   generator otomatis melewatinya.

## Cara Pakai Generator

```powershell
python buat_blast_wa.py hasil_kota.xlsx
# menghasilkan blast_wa.html -> buka di browser -> klik nomor satu per satu
# -> WhatsApp terbuka dengan pesan SUDAH TERISI -> tinggal tekan Kirim
```

## Menjawab Keberatan Umum (simpan, pasti ditanya!)

| Keberatan | Jawaban |
|---|---|
| "Sudah ada Instagram" | "IG itu numpang Kak — akun bisa kena suspend dan jangkauannya diatur algoritma. Website milik sendiri, muncul di pencarian Google, dan bisa ditaruh di bio IG. Keduanya saling melengkapi, bukan ganti-gantian." |
| "Mahal ya?" | "Simpelnya begini Kak: kalau websitenya hanya bawa 2 pelanggan tambahan per hari, biayanya balik modal dalam hitungan minggu. Dan Kakak bayar setelah website jadi & cocok, jadi nol risiko." |
| "Saya tidak ngerti teknologi" | "Justru itu Kak — Kakak tidak perlu menyentuh apa pun. Semua saya urus, Kakak tinggal terima link dan sesekali kirim foto menu baru kalau mau update." |
| "Nanti saja deh" | "Siap Kak, santai saja 😊 Sebagai gambaran: tiap bulan ada mahasiswa/pendatang baru di area setempat yang mencari cafe lewat Google — makin cepat onlinenya, makin cepat mereka menemukan cafe Kakak. Tapi tidak ada tekanan, kabari saja kapan pun siap." |
| "Bisa lihat contohnya?" | Kirim link demo/portfolio. (Ini alasan demo harus disiapkan SEBELUM blast.) |
| "Kamu dari mana dapat nomor saya?" | "Dari Google Maps Kak — nomornya memang tercantum publik di profil cafe Kakak. Saya hubungi karena cafenya menarik dan belum punya website." |


