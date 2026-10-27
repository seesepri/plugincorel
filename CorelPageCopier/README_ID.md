# Corel Page Copier — UI Modern

Versi ini menyederhanakan tampilan aplikasi dengan gaya Windows Forms yang lebih modern: header ringkas, dua kartu fungsi, tombol datar, hierarki visual lebih jelas, dan area status yang mudah dibaca.

## Fitur yang dipertahankan
- Koneksi ke CorelDRAW melalui COM.
- Daftar halaman tujuan, pilih semua/kosongkan, dan muat ulang.
- Copy objek terpilih ke halaman tujuan dengan validasi serta laporan hasil.
- Auto Select berdasarkan jenis objek dan cakupan halaman.
- Status progres saat memindai semua halaman.
- Navigasi hasil berhenti pada hasil terakhir dan menampilkan pemberitahuan.

## Fitur watermark (baru)
Buka lewat tombol **"Watermark..."** di kanan atas jendela utama. Halaman yang sedang
dicentang di daftar otomatis menjadi cakupan ("Halaman yang dicentang (N)").

- Watermark **teks** dan/atau **logo** (PNG/JPG) di layer khusus `WATERMARK`.
- **Logo bawaan SETA** tertanam di aplikasi dalam **dua varian** — cukup centang
  "Pakai logo bawaan SETA" lalu pilih **Warna logo**: *Hitam* (default, untuk desain
  berlatar terang) atau *Putih* (untuk desain berlatar gelap). Tidak perlu cari file.
- Opasitas default **50%**.
- Posisi jangkar 3×3 + margin, atau mode **tile** (pola berulang menutupi halaman).
- Opasitas, rotasi, warna teks, font, lebar logo (mm), nomor halaman otomatis.
- Cakupan: semua halaman atau halaman aktif saja.
- **Hapus watermark** kapan saja (idempoten: tekan Terapkan berkali-kali tetap 1 watermark).
- **Batch**: terapkan ke semua file `.cdr` dalam folder (+subfolder), default disimpan
  sebagai `nama_wm.cdr` sehingga file asli tidak tertimpa.
- Satu Ctrl+Z membatalkan seluruh operasi watermark pada satu dokumen.

File: `WatermarkLayout.vb` (model + matematika), `WatermarkEngine.vb` (COM),
`WatermarkDialog.vb` (UI). Detail desain & cara pasang: lihat `../ANALISA-WATERMARK.md`.

## Cara menjalankan
1. Ekstrak ZIP ke folder baru.
2. Buka `CorelPageCopier.vbproj` di Visual Studio.
3. Pilih konfigurasi x64/Release lalu Build Solution.
4. Buka CorelDRAW 2022 dan dokumen CDR yang akan digunakan.
5. Jalankan aplikasi hasil build.

## Audit reliabilitas (2026-10-27)

Perbaikan bug yang berisiko menghasilkan **laporan sukses palsu** atau output tidak valid.
Detail lengkap: `../docs/AUDIT-RELIABILITY.md`.

Perubahan perilaku yang perlu diketahui:

1. **Verifikasi copy dipisah**: setiap halaman tujuan berstatus **Verified** (geometri + struktur
   cocok dalam toleransi 0,5 mm), **Partial** (paste terjadi tetapi tidak dapat diverifikasi
   penuh), atau **Gagal**. Halaman *Partial* tidak pernah dihitung sebagai sukses.
2. **Konfirmasi sebelum mengubah dokumen**: dialog menampilkan nama dokumen, jumlah halaman,
   halaman sumber/tujuan, dan satuan dokumen.
3. **Satuan dokumen**: `Document.Unit` dibaca dulu; bila tidak terbaca atau tidak bisa diubah ke
   mm, operasi dibatalkan (tidak ada asumsi milimeter). Satuan asli selalu dipulihkan, dan
   kegagalan pemulihan dilaporkan.
4. **Undo**: bila grup undo gagal dibuka, aplikasi memberi tahu dan tidak lagi menjanjikan
   "satu Ctrl+Z".
5. **Batch**: file berakhiran `_wm` (output batch sebelumnya) **tidak** diproses ulang; status
   per file memisahkan OK / SEBAGIAN / GAGAL, dan status OK hanya bila watermark **dan**
   penyimpanan berhasil.
6. **Koneksi CorelDRAW**: instance yang sedang dipakai diutamakan; instance baru hanya dibuka
   **setelah Anda menyetujui** dialog konfirmasi.

## Pengujian
Panduan uji langkah-demi-langkah (data uji, angka ekspektasi, troubleshooting): `../PANDUAN-UJI.md`.

Test otomatis tanpa CorelDRAW (butuh .NET SDK 8):

```bat
python tools\static-checks.py
dotnet run --project tests\CorelPageCopier.RegressionTests
```

Kode dan struktur arsip diperiksa secara statis, ditambah 103 regression test otomatis yang tidak
membutuhkan CorelDRAW. **Build Visual Studio/MSBuild di Windows dan integrasi langsung dengan
CorelDRAW 2022 belum diverifikasi** (status NOT RUN) — lihat `../docs/AUDIT-RELIABILITY.md` §7.
Uji copy pada dokumen salinan terlebih dahulu sebelum digunakan pada dokumen penting.
