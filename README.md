# Isi Formulir Word (web server)

Upload file **.docx** → isi titik-titik (`.....`) dan centang kotak (`☐`) di browser → **unduh Word** yang sudah terisi.
Titik-titik & kotak centang dideteksi otomatis, jadi bisa dipakai untuk formulir Word lain yang pola isiannya sama.

## Cara menjalankan

Butuh Python 3.9+.

```bash
pip install -r requirements.txt
python app.py
```

Lalu buka **http://127.0.0.1:5000** di browser.

Opsional:
- Ganti port:  `PORT=8080 python app.py`  (Windows PowerShell: `$env:PORT=8080; python app.py`)
- Bisa diakses dari HP/komputer lain di jaringan yang sama: `HOST=0.0.0.0 python app.py`, lalu buka `http://IP-KOMPUTER:5000`

## Opsi di bar atas (aktif secara default)
- **Pas di titik-titik** – tulisan diletakkan di awal garis titik-titik dan sisa titiknya tetap ada sampai ujung garis asli.
  Jika isian terlalu panjang, ukuran font isian itu mengecil otomatis (minimal ±7,5 pt); jika masih tidak muat, teks dibiarkan turun ke baris berikutnya.
- **Muat 1 halaman** – seluruh formulir dipadatkan agar tepat 1 halaman. Urutan pemadatan: rapatkan spasi → perkecil margin → perkecil font seperlunya (sekecil mungkin yang masih dibutuhkan, bukan selalu maksimal).

## Mendukung form model "Label : nilai" (mis. Form Konversi Nilai MBKM)
- Isian `Label : nilai` yang sudah ada isinya muncul sebagai kotak yang bisa **diedit**
- `Label :` kosong (rata dengan tab) menjadi isian kosong, mis. "Total Durasi Kegiatan"
- Tanggal seperti `6 Juli 2026` dan `Cirebon, ......` jadi pemilih tanggal
- Sel tabel yang masih kosong menjadi isian per sel
- Kotak centang ActiveX diganti ☐/☑ biasa di file hasil
- Kolom tanda tangan bertitik-titik diberi label otomatis (Nama Pemohon, Nama Dosen Pembimbing)

## Fitur
- Tampilan mirip dokumen asli; isian berwarna kuning, hijau jika sudah terisi
- Kotak centang pilihan tunggal otomatis (Program Studi, Ya/Tidak, Diterima/Belum Diterima); daftar berkas bisa dicentang banyak
- Tanggal ("Cirebon, ..... 20.....") diisi lewat kalender → hasil `4 Oktober 2026`
- Nama & NIM otomatis tersalin ke bagian lain (tanda tangan, "Nama :", "NIM :"), tetap bisa diubah manual
- Tombol Enter pindah ke isian berikutnya
- Format, font, dan tabel di Word tetap utuh; hanya titik-titik dan ☐ yang diganti (☐ → ☑)

## Catatan
- Hanya `.docx` (bukan `.doc`; simpan ulang dulu dari Word).
- File upload disimpan sementara di folder temp sistem dan otomatis dihapus setelah 12 jam.
- Server ini untuk pemakaian lokal/internal. Jika dipasang di internet, jalankan di balik gunicorn/nginx + HTTPS.

## Deploy ke Vercel
Aplikasi sudah stateless (tidak menyimpan file di server), jadi langsung cocok untuk Vercel:
```bash
npm i -g vercel
vercel        # jalankan di folder ini; Vercel otomatis mengenali Flask (app.py + requirements.txt)
vercel --prod
```
Batas upload di Vercel 4,5 MB, jadi file formulir dibatasi 3 MB.
