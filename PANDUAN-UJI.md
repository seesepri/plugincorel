# Panduan Uji — Corel Page Copier + Fitur Watermark

Dokumen ini dipakai untuk **menguji build nyata di Windows + CorelDRAW**. Setiap skenario berisi langkah, hasil yang diharapkan (dengan angka konkret bila memungkinkan), dan tanda gagal.

> Aturan emas: **semua pengujian dilakukan pada salinan dokumen**, bukan file produksi. Batch default menyimpan `nama_wm.cdr` (non-destruktif), jangan mengaktifkan "Timpa file asli" selama pengujian.

---

## 0. Prasyarat

| Kebutuhan | Keterangan |
|---|---|
| OS | Windows 10/11 x64 |
| CorelDRAW | 2022 (ProgID `CorelDRAW.Application.24`). Versi lain masih dicoba lewat fallback `CorelDRAW.Application` — catat hasilnya bila bukan 2022 |
| .NET Framework | 4.8 (sudah bawaan Windows 10/11) |
| PowerShell | untuk membuka `%APPDATA%` & melihat log — opsional |

Waktu pengujian: **±20 menit** untuk smoke test (§2), **±60–90 menit** untuk seluruh dokumen ini.

---

## 1. Menyiapkan build

### Opsi A — ambil dari CI (tanpa Visual Studio)

1. Buka tab **Actions** repo → pilih run terbaru yang hijau.
2. Unduh artifact **`CorelPageCopier-Windows-x64`**.
3. Ekstrak ZIP → jalankan `CorelPageCopier.exe`.

### Opsi B — build sendiri

1. Buka `CorelPageCopier/CorelPageCopier.vbproj` di Visual Studio.
2. Pilih konfigurasi **Release | x64** → *Build Solution*.
3. Hasil: `CorelPageCopier/bin/x64/Release/CorelPageCopier.exe`.

### Cek cepat sebelum uji manual

```bat
python tools\static-checks.py
dotnet run --project tests\CorelPageCopier.RegressionTests
```
1. `static-checks.py` harus berakhir `SEMUA PEMERIKSAAN STATIS LULUS` (**200 pemeriksaan**).
2. Regression test (butuh .NET SDK 8) harus berakhir `Ringkasan: 103 lulus, 0 gagal`.

Kalau ada `FAIL`, jangan lanjut uji manual — laporkan dulu.

---

## 2. Data uji (buat sekali, pakai berulang)

Buat folder `D:\uji-watermark\`. Buat dokumen berikut **di CorelDRAW**, lalu simpan dengan nama persis seperti di tabel.

| # | Nama file | Isi dokumen | Dipakai untuk |
|---|---|---|---|
| 1 | `INV-2026-0912 klien budi.cdr` | 12 halaman A4 (210×297 mm), tiap halaman beda warna latar | token `{order}`, `{page}`, batch |
| 2 | `desain final.cdr` | 3 halaman A4, ada teks + bitmap + kurva | `{order}` fallback (tanpa angka), Auto Select |
| 3 | `dokumen inch.cdr` | 2 halaman **Letter**, satuan dokumen **inch** | normalisasi satuan (§3.2) |
| 4 | `campur ukuran.cdr` | 4 halaman: A4, A5, 100×150 mm, 210×297 mm | watermark ukuran halaman campur |
| 5 | `grup rumit.cdr` | 1 halaman, berisi grup + powerclip + layer tambahan | verifikasi geometri copy, Auto Select grup |
| 6 | `logo bertumpuk.cdr` | 1 halaman dengan 1 logo & 1 teks di posisi berbeda | uji copy ke halaman lain |

**Cara mengatur satuan dokumen ke inch (dokumen #3):** klik ganda penggaris → *Ruler* → *Units* → **Inches**, atau *Tools ▸ Options ▸ Document ▸ Rulers*.

**Tips:** untuk dokumen #1, beri teks kecil "HALAMAN 1", "HALAMAN 2", … di tiap halaman supaya mudah memastikan watermark per halaman tidak tertukar.

---

## 3. Uji per fitur

### 3.1 Koneksi & daftar halaman

| Langkah | Hasil yang diharapkan |
|---|---|
| Buka CorelDRAW + `INV-2026-0912 klien budi.cdr`, lalu jalankan aplikasi | Jendela terbuka, daftar halaman berisi 12 baris, halaman aktif ditandai `(AKTIF)` |
| Klik **Muat ulang** | Daftar tetap 12 baris, status "Dokumen terdeteksi." |
| Tutup dokumen CorelDRAW (biarkan aplikasi terbuka) → klik Muat ulang | Muncul pesan error yang jelas, **aplikasi tidak crash** |
| Jalankan aplikasi saat CorelDRAW belum terbuka | Pesan menyebut CorelDRAW 2022 harus terbuka + dokumen aktif |

### 3.2 Salin objek (temuan #1 & #7)

| Langkah | Hasil yang diharapkan |
|---|---|
| Di halaman 1, pilih 1 objek → centang halaman 2 & 3 → **Salin ke halaman terpilih** | Laporan: "Berhasil diverifikasi (geometri cocok): 2 halaman" |
| Ctrl+Z **sekali** di CorelDRAW | Kedua objek hasil paste hilang **sekaligus** |
| Paste ulang, lalu bandingkan posisi objek di halaman 2 dengan halaman 1 | Posisi & ukuran identik (selisih < 0,5 mm) |
| Ulangi pada `dokumen inch.cdr` (satuan inch) | Tetap "geometri cocok"; setelah operasi, satuan dokumen kembali **inch** (cek penggaris/double-click ruler) |
| Copy di `grup rumit.cdr` (objek berupa grup/powerclip) | Laporan tidak asal bilang "berhasil": bila geometri tak terukur, muncul baris "Terverifikasi lewat jumlah shape saja" |

**Tanda gagal:** perlu Ctrl+Z berkali-kali (undo grouping rusak), atau laporan bilang cocok padahal objek jelas bergeser.

### 3.3 Auto Select mengecualikan watermark (temuan #2)

Lakukan **setelah** §3.4 (watermark sudah ada).

| Langkah | Hasil yang diharapkan |
|---|---|
| Buka dialog Watermark → **Terapkan** di halaman 1 | Layer `WATERMARK` berisi shape bernama `WM_TEXT` / `WM_LOGO` |
| Tutup dialog → pilih **Semua objek** + **Halaman aktif** → **Select otomatis** | Jumlah objek terpilih **tidak termasuk** watermark (bandingkan dengan jumlah shape di luar layer `WATERMARK`) |
| Salin objek terpilih ke halaman 2, lalu periksa halaman 2 | Hanya objek desain yang tersalin; watermark halaman 2 **tidak menumpuk** |

**Tanda gagal:** ada `WM_*` ikut terpilih, atau watermark jadi dobel di halaman tujuan.

### 3.4 Watermark dasar

| Langkah | Hasil yang diharapkan |
|---|---|
| Buka **Watermark…** dari tombol kanan atas | Dialog terbuka; **Warna logo = Hitam**, **Opasitas = 50**, cakupan "Semua halaman" |
| Jenis **Teks saja**, teks `CONTOH`, Terapkan | Tiap halaman punya 1 teks baru di tengah, tanpa outline, di layer `WATERMARK` |
| Buka *Object Properties* → Transparency pada teks tersebut | **Transparency = 50** (karena opasitas 50%) |
| Jenis **Logo saja**, centang "Pakai logo bawaan SETA" varian **Hitam**, lebar 35 mm | Logo tampil dengan **latar transparan** (tidak ada kotak putih/hitam), lebar ≈ 35 mm, proporsi tidak gepeng |
| Ganti varian ke **Putih** lalu Terapkan ulang | Logo jadi putih, tetap transparan |
| Jenis **Teks + logo** | Dua objek (teks & logo), posisi sama-sama mengikuti jangkar |
| Pilih **Posisi = Kanan bawah**, margin 8 mm, rotasi 30° | Objek di sudut kanan bawah, berjarak ±8 mm dari tepi, miring 30° |
| Centang **Mode tile**, jarak 60×45 mm, logo 35 mm | Pola berulang menutupi halaman, jarak antar watermark rapi, **tidak menumpuk**, tepi halaman tetap ter-cover |
| Aktifkan **Kunci layer** lalu coba geser watermark | Layer `WATERMARK` terkunci (tidak bisa digeser/diedit) |

**Cek layer:** *Window ▸ Dockers ▸ Object Manager* → harus ada layer `WATERMARK`, ikon **printable aktif** (bukan ikon printer dicoret).

### 3.5 Idempoten & hapus watermark

| Langkah | Hasil yang diharapkan |
|---|---|
| Klik **Terapkan** 3× berturut-turut | Jumlah objek di layer `WATERMARK` **tetap** (tidak menumpuk) |
| **Hapus watermark** → Yes | Semua `WM_*` hilang dari 12 halaman, objek desain utuh |
| Hapus watermark lagi saat sudah tidak ada | Pesan "Tidak ada watermark yang ditemukan", bukan error |
| Ctrl+Z sekali setelah Hapus | Watermark kembali muncul (satu langkah undo) |

### 3.6 Teks berbeda per halaman (token)

| Langkah | Hasil yang diharapkan |
|---|---|
| Teks: `ORDER {order} · Hal. {page}/{pages}`, Terapkan ke semua halaman `INV-2026-0912 klien budi.cdr` | Halaman 1: `ORDER INV-2026-0912 · Hal. 1/12` … halaman 12: `… Hal. 12/12` |
| Ulangi pada `desain final.cdr` | `ORDER desain final · Hal. 1/3` (tanpa angka → nama file utuh) |
| Teks `{file} — {date} {time}` | Nama file tanpa ekstensi + tanggal `yyyy-MM-dd` dan jam `HH:mm` |
| Teks `Halo {tidak_ada}` | Tetap tertulis `Halo {tidak_ada}` (token asing dibiarkan) |
| Centang **Tambahkan nomor halaman** pada teks yang sudah memakai `{page}` | Nomor halaman **tidak dobel** |

### 3.7 PDF proof

| Langkah | Hasil yang diharapkan |
|---|---|
| Centang **Ekspor PDF proof** → Terapkan | Muncul `INV-2026-0912 klien budi_proof.pdf` di folder yang sama dengan file CDR |
| Buka PDF | **Watermark terlihat** di semua halaman, posisi & opasitas sesuai |
| Terapkan pada dokumen yang **belum pernah disimpan** | PDF ditaruh di `%TEMP%\CorelPageCopier` (tercatat di laporan) |

**Tanda gagal:** PDF ada tapi watermark tidak muncul → cek layer `WATERMARK` printable (§3.4); laporkan sebagai bug prioritas tinggi.

### 3.8 Batch + log CSV + buka folder hasil

| Langkah | Hasil yang diharapkan |
|---|---|
| Arahkan **Folder .cdr** ke `D:\uji-watermark`, centang **Sertakan subfolder**, **Jalankan batch** | Tiap file menghasilkan `*_wm.cdr`; file asli **tidak berubah** (cek tanggal/ukuran file) |
| Centang **Tulis log CSV** sebelum batch | Muncul `watermark-log.csv` di folder batch |
| Buka CSV di Excel | 7 kolom rapi: `Waktu; File; Halaman; Shape; Status; Pesan; Durasi_detik`; header **hanya sekali**; ada baris `RINGKASAN` di akhir |
| Klik **Buka folder hasil** setelah batch | Explorer terbuka, file hasil terakhir **tersorot** |
| Buat 1 file `.cdr` yang sedang dibuka di CorelDRAW, lalu batch lagi | File itu gagal dengan pesan jelas, file lain tetap diproses |
| Batch dengan nama file berisi `;` (mis. `logo; final.cdr`) | Di CSV tetap **satu sel** (kolom tidak bergeser) |

### 3.9 Preset

| Langkah | Hasil yang diharapkan |
|---|---|
| Atur teks `{order}`, logo putih, opasitas 20%, rotasi 25° → **Simpan** preset `Klien A` | Preset muncul di dropdown; file `%APPDATA%\CorelPageCopier\presets.ini` terbentuk |
| **Tutup aplikasi**, buka lagi, pilih preset `Klien A` | Semua field terisi kembali sama persis (termasuk varian logo & opasitas) |
| Simpan preset dengan nama sama (beda kapital, mis. `klien a`) | Preset **diganti**, tidak jadi duplikat |
| **Hapus** preset | Hilang dari dropdown dan dari `presets.ini` |
| Edit manual `presets.ini` (isi `OpacityPercent=abc`) lalu buka aplikasi | Aplikasi tidak crash; nilai kembali ke default 50 |

### 3.10 Pengaman re-entrancy (busy guard)

| Langkah | Hasil yang diharapkan |
|---|---|
| Batch 5+ file, saat berjalan klik tombol lain / klik **Terapkan** lagi | Klik ditolak; tombol nonaktif; status "Masih memproses: …" |
| Saat proses berjalan, klik **X** (tutup jendela) | Jendela tidak tertutup, ada pesan operasi masih berjalan |
| Setelah selesai / gagal di tengah | Semua tombol aktif kembali, kursor normal, aplikasi bisa dipakai lagi |
| Paksa gagal (tutup dokumen saat copy berjalan) | Laporan error muncul, UI tetap responsif — kunci busy **tidak** tertinggal |

---

## 4. Ringkasan ekspektasi angka

| Parameter | Nilai yang benar |
|---|---|
| Opasitas 50% | Transparency objek = **50** |
| Opasitas 25% | Transparency objek = **75** |
| Logo lebar 35 mm | `SizeWidth` objek ≈ **35 mm** (rasio asli terjaga) |
| Margin 8 mm | Jarak tepi objek ke tepi halaman ≈ **8 mm** |
| Toleransi verifikasi copy | selisih posisi/ukuran ≤ **0.5 mm** |
| Layer watermark | nama `WATERMARK`, **printable = on** |
| Nama shape | `WM_TEXT`, `WM_LOGO`, `WM_TEXT_1..n` (mode tile) |
| Folder preset | `%APPDATA%\CorelPageCopier\presets.ini` |
| Nama file log | `watermark-log.csv` (bisa diubah) |
| Suffix PDF | `_proof.pdf` |

---

## 5. Checklist (cetak / tempel di tiket)

```
[ ] 3.1  Koneksi & daftar halaman
[ ] 3.2  Copy + undo tunggal + geometri + dokumen inch
[ ] 3.3  Auto Select mengecualikan watermark
[ ] 3.4  Watermark teks / logo / gabungan / tile / rotasi / jangkar / lock
[ ] 3.5  Idempoten + hapus watermark + undo
[ ] 3.6  Token {page} {pages} {order} {file} {date} {time}
[ ] 3.7  PDF proof (watermark ikut tercetak)
[ ] 3.8  Batch 5 file + log CSV + buka folder hasil
[ ] 3.9  Preset: simpan, restart, timpa, hapus, file rusak
[ ] 3.10 Busy guard: klik ganda, tutup jendela, pemulihan setelah gagal
[ ] 11   Auto Select: kategori teks/bitmap/kurva/vektor/semua (AS-1 … AS-6)
[ ] 11   Auto Select: watermark tidak pernah terpilih; grup utuh; PowerClip = container
[ ] 11   Auto Select: navigasi semua halaman tanpa lompat/ulang, berhenti di hasil terakhir
[ ] 11   Auto Select: tidak ada hasil -> seleksi lama utuh, tidak ada objek terhapus
[ ] 12   Batch: file gagal/dipakai/rusak tidak dihitung sukses; stop -> DIBATALKAN
[ ] 12   PDF: PDF VALID hanya bila struktur+jumlah+ukuran halaman cocok (cek visual)
[ ] 12   CSV: satu baris per file, kolom tidak bergeser, ringkasan = jumlah file
[ ] Semua dokumen uji masih utuh (tidak ada file asli yang berubah)
```

---

## 6. Troubleshooting

| Gejala | Kemungkinan penyebab | Yang dilakukan |
|---|---|---|
| "Tidak dapat mengakses dokumen CorelDRAW" | CorelDRAW belum jalan / tidak ada dokumen aktif / versi bukan 2022 | Buka CorelDRAW + CDR, klik **Muat ulang**. Catat versi CorelDRAW-nya |
| Aplikasi membuka CorelDRAW kedua | Instance aktif tidak terdaftar di ROT | Tutup CorelDRAW kedua, buka dokumen di jendela pertama, lalu Muat ulang |
| Watermark tidak muncul | Layer `WATERMARK` tidak visible, atau objek tertutup objek lain | Cek Object Manager; pastikan layer visible |
| Watermark tidak ada di PDF | Layer tidak **printable**, atau object manager menyembunyikan layer | Set layer printable; ulangi ekspor |
| Logo tampak sebagai kotak (ada latar) | Memakai JPG berlatarbelakang (bukan PNG transparan) | Pakai logo bawaan SETA (sudah PNG transparan) atau PNG transparan sendiri |
| Teks `{page}` tidak terganti | Token ditulis dengan kurung lain, mis. `[page]`, atau salah eja | Gunakan `{page}` (kurung kurawal); cek daftar token di dialog |
| `{order}` berisi seluruh nama file | Memang aturan fallback saat tidak ada angka ≥ 4 digit | Ganti nama file agar memuat nomor order (mis. `INV-2026-0912 …`) |
| Preset tidak muncul setelah restart | `%APPDATA%` tidak bisa ditulis (kebijakan IT) | Cek apakah `presets.ini` terbentuk; laporkan path & pesan error |
| CSV berantakan di Excel | Regional setting memakai koma sebagai pemisah | File memakai `;` + BOM — buka dengan *Data ▸ From Text/CSV* bila tetap bermasalah |
| Tombol "Buka folder hasil" nonaktif | Belum ada output, atau folder sudah dihapus/dipindah | Jalankan Terapkan/batch dulu |
| Copy lambat di dokumen besar | Normal — tiap halaman diukur untuk verifikasi geometri | Bandingkan "Durasi" di laporan |

---

## 7. Kalau menemukan bug

Sertakan (agar bisa direproduksi cepat):

1. Nama & versi CorelDRAW (`Help ▸ About`).
2. Skenario nomor di dokumen ini (mis. §3.6) + langkah persisnya.
3. Isi kotak **laporan** di dialog (copy–paste teksnya).
4. `watermark-log.csv` bila terkait batch.
5. Screenshot Object Manager (layer `WATERMARK`) dan dialog *Object Properties* bila terkait opasitas/ukuran.
6. Nama file dokumen uji (penting untuk kasus `{order}`).

---

## 8. Setelah selesai

1. Jalankan **Hapus watermark** di semua dokumen uji.
2. Hapus file `*_wm.cdr` dan `*_proof.pdf` hasil uji.
3. Pastikan tidak ada file asli yang berubah (bandingkan tanggal modifikasi).
4. Simpan bukti uji (screenshot + CSV) ke tiket/PR.

---

## 9. Lampiran — peta fitur ke kode

| Fitur | File | Bagian analisa |
|---|---|---|
| Salin objek + undo grouping + geometri | `Program.vb` (`CopyClicked`) | `ANALISA-WATERMARK.md` §4.4 |
| Auto Select — klasifikasi objek | `ShapeClassifier.vb` | `docs/AUDIT-AUTOSELECT.md` §3–§4 |
| Auto Select — penelusuran objek (grup/PowerClip) | `ShapeScanner.vb`, `ShapeNode.vb` | idem §3.4–§3.7 |
| Auto Select — navigasi hasil & verifikasi seleksi | `SelectionNavigation.vb` | idem §3.7–§3.9 |
| Auto Select — jembatan UI | `Program.vb` (`AutoSelectClickedCore`, `ApplySelection`) | idem §3 |
| Busy guard | `Program.vb` + `OperationGate` | §4.5 |
| Watermark inti | `WatermarkEngine.vb` | §3, §4 |
| Model & matematika | `WatermarkLayout.vb` | §3.2, §3.4 |
| UI dialog | `WatermarkDialog.vb` | §4.6, §4.7 |
| Preset | `PresetStore.vb` | §4.7 |
| Log CSV | `BatchLog.vb` | §4.7 |
| Token per halaman | `WatermarkTokens.vb` | §4.8 |
| Batch: stop/exception/pembukuan | `BatchRunTracker.vb`, `WatermarkEngine.vb` | `docs/AUDIT-WATERMARK-BATCH.md` §1 |
| Penempatan & verifikasi posisi | `WatermarkPlacementCheck.vb`, `WatermarkLayout.vb` | idem §2 |
| Validasi PDF proof | `PdfProofValidator.vb` | idem §3 |
| Format log CSV | `BatchLog.vb` | idem §4 |
| Pemeriksaan statis | `tools/static-checks.py` | §4.4–§4.8 |

---

## 10. Uji audit reliabilitas (BARU — status: **NOT RUN** di lingkungan audit)

Bagian ini khusus memverifikasi perbaikan audit reliabilitas 2026-10-27
(lihat `docs/AUDIT-RELIABILITY.md`). Belum pernah dijalankan pada CorelDRAW nyata.

### 10.1 Persiapan data uji tambahan

Buat di `D:\uji-audit\`:

| # | Nama file | Isi |
|---|---|---|
| A1 | `satu-objek.cdr` | 3 halaman; halaman 1 berisi **1 persegi** 50×50 mm di (100,100); halaman 2 & 3 **sudah berisi** beberapa objek lain (teks + logo) di area berbeda |
| A2 | `multi-seleksi.cdr` | 2 halaman; halaman 1 berisi 3 objek berjauhan (teks, bitmap, kurva) — untuk multi-selection |
| A3 | `grup-bersarang.cdr` | 2 halaman; halaman 1 berisi grup berisi sub-grup + kurva + teks |
| A4 | `bitmap-teks.cdr` | 2 halaman; halaman 1 berisi bitmap besar + teks kecil (ukuran campuran) |
| A5 | `satuan-inch.cdr` | 2 halaman, **Rulers = inches** (Layout ▸ Document Options), objek persegi 2×2 inch |
| A6 | `satuan-point.cdr` | 2 halaman, **Rulers = points**, objek persegi 144×144 pt |
| B1 | folder `D:\uji-audit\batch\` | `a.cdr`, `b.cdr`, subfolder `sub\c.cdr`, **plus** `a_wm.cdr` & `b_wm.cdr` (output lama dari run sebelumnya) |
| B2 | folder `D:\uji-audit\batch-benturan\` | `x.cdr` dan `x_wm.cdr` yang **dibuat manual** (bukan hasil aplikasi) |

### 10.2 Skenario copy (Program.vb)

| # | Langkah | Hasil yang diharapkan (versi baru) |
|---|---|---|
| 10.2.1 | A1: pilih objek halaman 1, centang halaman 2 & 3, **Salin** | Dialog konfirmasi menampilkan **nama dokumen + jumlah halaman + halaman sumber/tujuan + satuan**. Setelah OK: laporan `Verified: 2`, `Partial: 0`, `Gagal: 0`; halaman 2 & 3 dapat **Verified** meskipun sudah berisi objek lain (bug lama: halaman 2 dilaporkan gagal) |
| 10.2.2 | A2: pilih **3 objek** dengan Shift-klik, salin ke halaman 2 | `Verified: 1`; laporan menyebut geometri & struktur cocok; jumlah objek di halaman 2 bertambah 3 |
| 10.2.3 | A3: pilih grup bersarang, salin | `Verified: 1`; struktur grup di halaman tujuan tetap berupa **satu grup** (cek Object Manager) |
| 10.2.4 | A4: pilih bitmap + teks (ukuran campuran), salin | `Verified: 1`; Δw/Δh kecil (< 0,5 mm) |
| 10.2.5 | A5 (inch): copy antar halaman | `Verified`; laporan menyebut satuan inch + konversi eksplisit **atau** "dinormalkan ke mm"; satuan dokumen **kembali ke inch** setelah operasi |
| 10.2.6 | A6 (point): copy antar halaman | Sama seperti 10.2.5, satuan akhir = points |
| 10.2.7 | A1: setelah paste, geser objek hasil 2 mm secara manual **lalu** ulangi copy pada salinan (atau bandingkan Δ di laporan) | Δx/Δy yang dilaporkan ≈ 0; bila sengaja digeser, status harus **Failed** (bukan Verified) |
| 10.2.8 | Tekan **Ctrl+Z sekali** setelah copy multi-halaman | Semua paste dibatalkan sekaligus; laporan menyebut nama grup undo `Salin objek ke halaman` |
| 10.2.9 | Buka dokumen A, klik **Muat ulang**, aktifkan dokumen lain di CorelDRAW, klik **Salin** | Ditolak: "Dokumen aktif berbeda dari dokumen saat daftar halaman dimuat" (bukan menyalin ke dokumen salah) |
| 10.2.10 | Tutup CorelDRAW, lalu klik **Salin** | Muncul pertanyaan sebelum membuka instance baru (default **No**); bila Yes → ada pemberitahuan bahwa instance baru dibuka |

### 10.3 Skenario batch (WatermarkEngine.vb)

| # | Langkah | Hasil yang diharapkan |
|---|---|---|
| 10.3.1 | B1: jalankan batch (**jangan** centang "Timpa file asli") | Hanya `a.cdr`, `b.cdr`, `sub\c.cdr` diproses. `a_wm.cdr`, `b_wm.cdr` muncul di bagian **"File dilewati"** laporan dengan alasan "output batch sebelumnya". **Tidak ada** `*_wm_wm.cdr` |
| 10.3.2 | Jalankan batch yang sama **lagi** | Hasil identik (idempoten): `a_wm.cdr`/`b_wm.cdr` diperbarui, tetap tidak ada `_wm_wm` |
| 10.3.3 | B2: jalankan batch | `x.cdr` → `x_wm.cdr` **ditimpa** (regenerasi, disebut di laporan). File asli `x.cdr` tidak berubah (bandingkan tanggal modifikasi) |
| 10.3.4 | Buat `d.cdr` di B1, buka & kunci file itu di aplikasi lain (atau set read-only), jalankan batch | Baris CSV untuk `d.cdr` = `GAGAL`, ringkasan = `OK 3, SEBAGIAN 0, GAGAL 1`; **tidak** masuk hitungan sukses; pesan UI sama angkanya; detail exception ada di kolom Pesan |
| 10.3.5 | Dokumen dengan 3 halaman, buat 1 halaman gagal (mis. kunci layer `WATERMARK` + nonaktifkan "Ganti watermark lama") | Baris CSV = `SEBAGIAN`, bukan `OK` |
| 10.3.6 | Centang "Timpa file asli" | Konfirmasi menyebut **file asli DITIMPA**; file read-only ditolak dengan alasan; file `.cdr` di luar folder tidak pernah disentuh |
| 10.3.7 | Periksa `watermark-log.csv` | **Satu** baris per file (bukan dua), kolom Status ∈ {OK, SEBAGIAN, GAGAL}, baris `RINGKASAN` memuat jumlah OK/SEBAGIAN/GAGAL |

### 10.4 Hasil uji (isi saat dijalankan di Windows + CorelDRAW 2022)

| Skenario | Status | Catatan |
|---|---|---|
| 10.2.1 – 10.2.10 | **NOT RUN** | CorelDRAW tidak tersedia di lingkungan audit |
| 10.3.1 – 10.3.7 | **NOT RUN** | idem |
| Build Release x64 (MSBuild Windows) | **NOT RUN** | idem |

Tempelkan bukti (isi kotak laporan + CSV + screenshot Object Manager) ke tiket/PR, lalu ubah
status di atas menjadi PASS/FAIL.

---

## 11. Uji audit Auto Select (BARU — status: **NOT RUN** di lingkungan audit)

Bagian ini memverifikasi perbaikan Auto Select 2026-10-27
(lihat `docs/AUDIT-AUTOSELECT.md`). Regression test otomatis memakai struktur
dataset yang sama, jadi angka di UI bisa dibandingkan langsung dengan angka di
test (`tests/CorelPageCopier.RegressionTests/Tests/ShapeScannerTests.vb`).

**Aturan penting saat menguji:** Auto Select **tidak boleh** menghapus satu objek
pun. Sebelum tiap skenario, simpan salinan dokumen atau periksa jumlah objek di
Object Manager; setelah skenario, jumlah objek harus sama.

### 11.1 Persiapan data uji

Buat folder `D:\uji-autoselect\` dan buat dokumen berikut di CorelDRAW 2022
(late-bound, aplikasi 64-bit).

| # | Nama file | Isi dokumen |
|---|---|---|
| AS-1 | `teks.cdr` | 1 halaman: 1 teks artistik ("Judul"), 1 teks paragraf (3 baris), 1 grup berisi 1 teks paragraf |
| AS-2 | `bitmap.cdr` | 1 halaman: 1 bitmap bebas, 1 grup berisi 2 bitmap, 1 PowerClip (rectangle) berisi 1 bitmap |
| AS-3 | `kurva.cdr` | 1 halaman: 1 kurva terbuka (garis 3 titik, tidak ditutup), 1 kurva tertutup (buat dengan *Auto-Close* / gabungkan ujung), 1 grup bersarang 2 level berisi 1 kurva terbuka |
| AS-4 | `vektor.cdr` | 1 halaman: rectangle, ellipse, polygon, 1 kurva terbuka, 1 kurva tertutup, **plus** 1 garis bantu horizontal & 1 objek OLE (*Insert ▸ OLE Object*, mis. lembar kerja kosong) |
| AS-5 | `campur.cdr` | 1 halaman: gabungan skenario di atas (teks, bitmap, kurva, vektor, 3 grup, 1 PowerClip, 1 objek OLE, 1 guideline) — **tanpa** watermark |
| AS-6 | `tanpa-teks.cdr` | 1 halaman: hanya 2 rectangle. Untuk uji "tidak ada hasil" |
| AS-7 | `watermark-campur.cdr` | 3 halaman: (a) halaman 1 berisi watermark `WM_`+ objek sah, (b) halaman 2 berisi grup campuran (bitmap sah + objek `WM_`), (c) halaman 3 tanpa watermark. Beri watermark lewat aplikasi ini dulu (*Watermark…*) supaya penamaan sesuai |

Semua dokumen: **satuan mm**, jumlah halaman seperti tertulis.

### 11.2 Skenario kategori (cakupan "Halaman aktif")

Jalankan Auto Select dengan cakupan **Halaman aktif** pada tiap dokumen, lalu
bandingkan angka di kotak laporan dengan tabel berikut.

| # | Dokumen | Kategori | Harapan: jumlah "objek cocok" | Harapan: seleksi terverifikasi | Yang tidak boleh terjadi |
|---|---|---|---|---|---|
| 11.2.1 | AS-1 | Teks | **3** (2 level atas + 1 dalam grup) | 3 objek (atau peringatan bila CorelDRAW memilih grupnya) | teks tidak terlewat; bitmap tidak ikut |
| 11.2.2 | AS-1 | Bitmap | 0 | tidak ada operasi | seleksi lama **tidak** hilang, tidak ada dialog error |
| 11.2.3 | AS-2 | Bitmap | **4** | 4 target / 3 bila isi PowerClip dipilih sebagai container (lihat catatan) | isi PowerClip **tidak** diklaim terpilih sendiri |
| 11.2.4 | AS-3 | Kurva terbuka | **2** | 2 | kurva tertutup tidak ikut |
| 11.2.5 | AS-4 | Vektor | **6** | 6 | guideline & objek OLE **tidak** ikut (cek daftar objek di laporan) |
| 11.2.6 | AS-5 | Semua objek | = jumlah objek level atas, **grup dihitung 1** | sama | tidak ada grup yang dipecah; PowerClip dihitung 1 (container) |
| 11.2.7 | AS-6 | Teks | **0** | tidak ada operasi | pesan "Seleksi yang ada tidak diubah"; objek di kanvas tetap seperti sebelumnya |
| 11.2.8 | AS-7 | Bitmap | 3 halaman dipindai; halaman 3 tetap 0 | — | objek `WM_` **tidak pernah** ada di daftar hasil |

**Cara memeriksa "seleksi terverifikasi":** kotak laporan memuat baris
`Target seleksi : N` dan `Terverifikasi : N objek` (atau "TIDAK terverifikasi").
Bila CorelDRAW melaporkan jumlah berbeda, dialog memakai ikon **peringatan** —
itu bukan PASS untuk skenario tersebut, catat angkanya.

### 11.3 Skenario semua halaman & navigasi

| # | Langkah | Hasil yang diharapkan |
|---|---|---|
| 11.3.1 | AS-7, kategori Bitmap, cakupan **Semua halaman**, Auto Select | Pemindaian menyebut progres "halaman 1/3 … 3/3"; laporan memuat **total objek** + rincian per halaman (termasuk halaman dengan 0) dan jumlah halaman dipindai = 3 |
| 11.3.2 | Setelah 11.3.1, klik "Pilih hasil berikutnya" berulang | Urutan halaman hasil **naik, tanpa lompat & tanpa ulang**; status menyebut "hasil ke-2 dari 2"; di hasil terakhir muncul `AtEnd` dan **tidak** kembali ke halaman pertama |
| 11.3.3 | Setelah 11.3.1, ubah ComboBox kategori ke Vektor, lalu klik "Pilih hasil berikutnya" | Ditolak dengan pesan bahwa hasil pemindaian milik kategori lain (minta pindai ulang) — **tidak** memilih objek dengan kategori berbeda |
| 11.3.4 | Setelah 11.3.1, hapus satu halaman hasil secara manual, lalu klik "Pilih hasil berikutnya" | Pesan "halaman hasil tidak ditemukan"; posisi **tidak** dimajukan |
| 11.3.5 | Matikan Auto Select di tengah pemindaian dengan menutup dokumen | Muncul error yang jelas, status busy dilepas, tombol Auto Select bisa dipakai lagi (tidak "menggantung") |
| 11.3.6 | Klik Auto Select dua kali cepat (double click) | Hanya satu pemindaian berjalan (busy guard), tidak ada dua dialog |

### 11.4 Skenario keamanan (tidak destruktif)

| # | Langkah | Hasil yang diharapkan |
|---|---|---|
| 11.4.1 | Pilih 1 objek di kanvas, jalankan Auto Select kategori yang **tidak** punya hasil | Seleksi awal **tetap** terpilih setelah dialog ditutup (tidak ada `ClearSelection`) |
| 11.4.2 | Kunci layer (ikon gembok di Object Manager), jalankan Auto Select kategori vektor pada layer itu | Laporan menyebut objek cocok **dan** jumlah "tidak dapat dipilih" (layer terkunci) — bukan mengklaim semua terpilih |
| 11.4.3 | Bandingkan jumlah objek di Object Manager sebelum & sesudah semua skenario | **Sama persis**; tidak ada objek yang terhapus/berubah posisi |
| 11.4.4 | Tekan Ctrl+Z setelah Auto Select (beberapa kali) | Tidak ada perubahan dokumen yang perlu di-undo dari Auto Select (hanya seleksi) |

### 11.5 Hasil uji (isi saat dijalankan di Windows + CorelDRAW 2022)

| Skenario | Status | Catatan |
|---|---|---|
| 11.2.1 – 11.2.8 | **NOT RUN** | CorelDRAW tidak tersedia di lingkungan audit |
| 11.3.1 – 11.3.6 | **NOT RUN** | idem |
| 11.4.1 – 11.4.4 | **NOT RUN** | idem |
| Regression test otomatis (157 test) | **PASS** | `dotnet run --project tests/CorelPageCopier.RegressionTests -c Release` → `157 lulus, 0 gagal` |
| Pemeriksaan statis | **PASS** | `python3 tools/static-checks.py` → 255 lulus, 0 gagal |
| Kompilasi net48 (lingkungan audit) | **PASS** | 0 error |
| Build Release x64 (MSBuild Windows) | **PASS** | CI run 40446380975 (`windows-latest`, commit `a859858`): `bin\x64\Release\CorelPageCopier.exe`, Build succeeded, 0 warning, 0 error |
| Static checks + regression test di Windows | **PASS** | run yang sama: 255 PASS / 0 FAIL dan `157 lulus, 0 gagal` |

Tempelkan isi kotak laporan Auto Select + screenshot Object Manager (jumlah objek
sebelum/sesudah) ke tiket/PR, lalu ubah status di atas menjadi PASS/FAIL.

---

## 12. Uji audit batch & PDF proof (BARU — status: **NOT RUN** di lingkungan audit)

Bagian ini memverifikasi perbaikan pada `docs/AUDIT-WATERMARK-BATCH.md`.
Dibagi tegas menjadi tiga tingkat bukti yang **tidak boleh dicampur**:

1. **Otomatis** (unit/regression test) — sudah dijalankan, angka di §12.7.
2. **Validasi struktur PDF** — otomatis untuk struktur; hasil pada PDF nyata
   CorelDRAW **belum** diuji (isi tabel §12.6 saat menjalankan).
3. **Pemeriksaan visual** — hanya bisa dilakukan manusia di CorelDRAW /
   pembaca PDF. **Tidak ada klaim "siap cetak"** dari bagian mana pun di sini.

### 12.1 Data uji

Di `D:\uji-batchpdf\`:

| # | Nama | Isi |
|---|---|---|
| W1 | `satu-teks.cdr` | 3 halaman A4 mm, isi objek desain (kotak + teks) — untuk watermark teks |
| W2 | `logo.cdr` | 2 halaman A4, untuk watermark logo (PNG 600×200 px) |
| W3 | `teks-logo.cdr` | 2 halaman A4, untuk kombinasi teks + logo |
| W4 | `campur-ukuran.cdr` | 4 halaman: A4, A5, 100×150 mm, A4 landscape |
| W5 | `inch.cdr` | 2 halaman Letter, **Rulers = inches** |
| W6 | `point.cdr` | 2 halaman A4, **Rulers = points** |
| W7 | `tile.cdr` | 1 halaman A4 untuk mode tile (jarak X 60, Y 45) |
| B1 | folder `batch\` | `a.cdr`, `b.cdr`, `c.cdr` (masing-masing 2–3 halaman), plus `a_wm.cdr` (output lama) |
| B2 | folder `batch-dipakai\` | `x.cdr` yang **sedang dibuka** di CorelDRAW saat batch dijalankan |
| B3 | folder `batch-rusak\` | `y.cdr` yang sudah dirusak (rename acak isinya / 0 byte) |
| B4 | folder `batch-tolak-tulis\` | `z.cdr` + folder read-only (atau jalankan sebagai user tanpa hak tulis) |

### 12.2 Fokus A — batch aman

| # | Langkah | Hasil yang diharapkan |
|---|---|---|
| 12.2.1 | B1, **tanpa** centang "Timpa file asli", jalankan batch | `a_wm.cdr` lama dilewati (alasan "output batch sebelumnya"); `a_wm.cdr`, `b_wm.cdr`, `c_wm.cdr` dibuat/diperbarui; file asli **tidak berubah** (bandingkan tanggal modifikasi) |
| 12.2.2 | Ulangi 12.2.1 | Idempoten: hasil sama, tidak ada `_wm_wm.cdr` |
| 12.2.3 | B2 (file sedang terbuka), jalankan batch | File itu berstatus **GAGAL** dengan alasan jelas; file lain tetap **OK** |
| 12.2.4 | B3 (file rusak), jalankan batch | **GAGAL**, alasan menyebut file tidak dapat dibuka |
| 12.2.5 | B4 (folder tanpa hak tulis), jalankan batch | File diproses, tetapi muncul peringatan **"Baris log CSV TIDAK dapat ditulis"**; ringkasan tetap ditampilkan |
| 12.2.6 | Jalankan batch 10+ file, lalu klik **Hentikan** pada file ke-3 | File ke-3 selesai disimpan; file 4–10 berstatus **DIBATALKAN**; ringkasan: `OK + DIBATALKAN = jumlah file`; jumlah baris CSV = 10 + ringkasan + header |
| 12.2.7 | Setelah 12.2.6, buka Object Manager / Task Manager | Tidak ada dokumen `*_wm.cdr` yang tertinggal terbuka di CorelDRAW |
| 12.2.8 | Paksa exception: cabut media/rename folder saat batch berjalan (atau tutup CorelDRAW) | Muncul pesan error; **ringkasan tetap tertulis**; file yang belum diproses tercatat DIBATALKAN atau GAGAL (tidak ada yang hilang) |
| 12.2.9 | Batch, lalu **Ctrl+Z** satu kali di dokumen yang baru dibuka | Tidak perlu: batch menutup dokumen tanpa menyimpan ke file asli (mode file baru). Pada mode timpa, dokumen sumber berubah dan itu **dinyatakan** di konfirmasi |
| 12.2.10 | Cek file output | Setiap output bukan 0 byte dan dapat dibuka kembali di CorelDRAW |

### 12.3 Fokus B — akurasi watermark (visual)

Untuk tiap baris: terapkan watermark, lalu **periksa dengan mata** dan catat
apakah posisi sesuai. Bila laporan menampilkan peringatan penempatan
(`penempatan ... menyimpang dari perhitungan`), catat angkanya — itu bukan PASS.

| # | Dokumen / setelan | Yang diperiksa |
|---|---|---|
| 12.3.1 | W1, teks, anchor **Center**, margin 10 | Teks tepat di tengah; jarak ke tepi ≥ margin |
| 12.3.2 | W1, sembilan anchor 3×3 (TopLeft … BottomRight) | Tiap posisi menempel tepi sesuai margin 10 mm, tidak keluar halaman |
| 12.3.3 | W1, margin 40 mm | Jarak 40 mm dari tepi; tidak ada bagian terpotong |
| 12.3.4 | W1, teks sangat panjang (200+ karakter) | Bila lebih lebar dari halaman: muncul peringatan **"lebih besar dari halaman … terpotong"**; teks dipusatkan, bukan menghilang ke kiri |
| 12.3.5 | W1, rotasi 30° / 45° / 90° | Sudut sesuai; posisi tetap di dalam halaman |
| 12.3.6 | W2, skala logo 35 mm vs 70 mm | Rasio aspek logo tidak berubah (bandingkan lebar:tinggi) |
| 12.3.7 | W2, opasitas 20% / 50% / 100% | 20% paling transparan; 100% pekat; tidak ada logo yang hilang |
| 12.3.8 | W7, mode tile, jarak X 60 / Y 45 | Pola merata; **tidak ada** watermark terpotong di tepi; jarak antar watermark konsisten; jumlah shape di Object Manager = jumlah posisi di laporan |
| 12.3.9 | W7, jarak X 10 / Y 5 (lebih kecil dari ukuran objek) | Tidak saling menumpuk (jarak minimum otomatis 110%/140%) |
| 12.3.10 | W4 (A4/A5/100×150/landscape) | Setiap halaman memakai ukuran halamannya sendiri; landscape tidak tertukar |
| 12.3.11 | W5 (inch) dan W6 (point) | Posisi & ukuran sama seperti dokumen mm; **satuan dokumen kembali** ke inch/point setelah operasi (cek Rulers) |
| 12.3.12 | W3 (teks + logo) | Keduanya tampil; tidak saling menutupi; rotasi berlaku untuk keduanya |
| 12.3.13 | W1 dengan "Ganti watermark lama" aktif, jalankan 3× | Tetap 1 watermark per halaman (tidak menumpuk) |
| 12.3.14 | W1, lalu **Hapus watermark** | Semua objek `WM_*` hilang; objek desain lain **tidak berubah** (bandingkan posisi/warna) |
| 12.3.15 | Token: teks `ORDER {order} · Hal. {page}/{pages} · {file} · {date} {time}` di 3 halaman | Halaman 1/2/3 benar; `{pages}` = 3; `{order}` dari nama file (mis. `INV-2026-0912.cdr` → `2026-0912`); tanggal/jam masuk akal |

### 12.4 Fokus C — PDF proof

| # | Langkah | Hasil yang diharapkan |
|---|---|---|
| 12.4.1 | W1, centang PDF proof, terapkan | Laporan menyebut **`PDF VALID (n halaman, struktur & ukuran cocok)`** dan file `<nama>_proof.pdf` ada |
| 12.4.2 | Buka PDF di Adobe Acrobat Reader / Edge | PDF **terbuka tanpa error**; jumlah halaman = jumlah halaman dokumen (bandingkan status `halaman x/y`) |
| 12.4.3 | Periksa properti dokumen PDF (Reader: File ▸ Properties) | Ukuran halaman sama dengan dokumen sumber (A4 210×297 mm = 595×842 pt); orientasi landscape untuk halaman landscape |
| 12.4.4 | Lihat tiap halaman | Watermark terlihat di halaman yang ditentukan; halaman lain bersih |
| 12.4.5 | Periksa tepi halaman pada zoom 400% | Tidak ada watermark terpotong (clipping) di tepi |
| 12.4.6 | Bandingkan opasitas watermark di PDF vs di CorelDRAW | Sama (PDF tidak menambah/mengurangi transparansi) |
| 12.4.7 | PDF dari W4 (halaman campur ukuran) | 4 halaman dengan ukuran masing-masing benar; laporan tetap `PDF VALID` |
| 12.4.8 | Rename/rusak file PDF hasil sebelum dibuka, lalu ulangi ekspor ke nama yang sama (folder read-only) | Muncul kegagalan; **tidak** ada klaim "PDF valid" |
| 12.4.9 | Cek laporan bila PDF berstatus **TIDAK DAPAT DIVERIFIKASI** | Status file = **SEBAGIAN**; pesan meminta cek manual — bukan PASS |

> Catatan: validasi struktur PDF **tidak** membuktikan kualitas cetak. Untuk
> klaim siap cetak, perlu uji cetak nyata (di luar cakupan audit ini).

### 12.5 Fokus D — log CSV

| # | Langkah | Hasil yang diharapkan |
|---|---|---|
| 12.5.1 | Buka `watermark-log.csv` di Excel (regional Indonesia) | Terbuka rapi, kolom: `Waktu; File; Halaman; Shape; Status; Pesan; Durasi_detik` |
| 12.5.2 | Batch dengan satu file gagal (B2/B3) | Satu baris per file + satu baris `RINGKASAN`; status di CSV = status di UI = ringkasan |
| 12.5.3 | Batch yang dihentikan (12.2.6) | Baris `DIBATALKAN` ada sebanyak file yang tidak diproses; `OK + SEBAGIAN + GAGAL + DIBATALKAN` = jumlah file |
| 12.5.4 | File dengan nama memuat `;` dan tanda kutip (mis. `a;b"c.cdr`) | Nama utuh dalam satu sel, kolom tidak bergeser |
| 12.5.5 | Pesan error multi-baris | Tetap satu baris fisik di CSV (isi pesan digabung dengan `/`) |
| 12.5.6 | Periksa kolom `Waktu` dan `Durasi_detik` | Waktu masuk akal (format `yyyy-MM-dd HH:mm:ss`), durasi > 0 |

### 12.6 Hasil uji (isi saat dijalankan di Windows + CorelDRAW 2022)

| Skenario | Status | Catatan |
|---|---|---|
| 12.2.1 – 12.2.10 (batch) | **NOT RUN** | CorelDRAW tidak tersedia di lingkungan audit |
| 12.3.1 – 12.3.15 (visual) | **NOT RUN** | idem |
| 12.4.1 – 12.4.9 (PDF) | **NOT RUN** | idem |
| 12.5.1 – 12.5.6 (CSV) | **NOT RUN** di CorelDRAW; format sudah diuji unit | |
| Validasi struktur PDF pada PDF **nyata** CorelDRAW | **NOT RUN** | validator hanya diuji dengan PDF sintetis |
| Cetak fisik / kualitas siap cetak | **NOT RUN** | tidak diklaim |

### 12.7 Bukti otomatis (sudah dijalankan)

| Pemeriksaan | Hasil |
|---|---|
| Regression test (net8.0, tanpa CorelDRAW) | `216 lulus, 0 gagal` — termasuk 59 test baru (PDF 16, penempatan 19, log 10, batch 12) |
| Pemeriksaan statis | `295 pemeriksaan lulus, 0 gagal` |
| Kompilasi net48 | `Build succeeded, 0 warning, 0 error` |
| Build Release x64 Windows | **PASS** — CI run 40447818533 (`windows-latest`, commit `470ae08`): `bin\x64\Release\CorelPageCopier.exe`, 0 warning, 0 error; static checks 295 PASS; test 216 lulus / 0 gagal |

Tempelkan isi kotak laporan + potongan CSV + screenshot PDF/Object Manager ke
tiket/PR, lalu ubah status di tabel §12.6 menjadi PASS/FAIL. Status **hanya**
boleh PASS bila langkah visual benar-benar diperiksa dengan mata.

---

## 13. Uji refactor arsitektur (BARU — status: **NOT RUN** di lingkungan audit)

Refactor 2026-10-27 memindahkan logika koneksi/copy/Auto Select dari
`Program.vb` ke service. Yang harus dibuktikan ulang di mesin ber-CorelDRAW
adalah **perilaku yang sama**, bukan tampilan baru.

### 13.1 Bukti otomatis (sudah dijalankan, tanpa CorelDRAW)

| Pemeriksaan | Hasil |
|---|---|
| Regression test | `253 lulus, 0 gagal` (216 test lama tetap lulus + 37 test service) |
| Pemeriksaan statis | `351 pemeriksaan lulus, 0 gagal` |
| Kompilasi net48 | `Build succeeded, 0 warning, 0 error` |

### 13.2 Uji manual yang WAJIB di CorelDRAW (status awal: NOT RUN)

| No | Langkah | Hasil yang benar | Status |
|---|---|---|---|
| 13.2.1 | Buka aplikasi tanpa CorelDRAW berjalan, klik **Muat ulang** | dialog izin membuka instance baru muncul (default **Tidak**); pilih Tidak → tidak ada jendela CorelDRAW baru | NOT RUN |
| 13.2.2 | CorelDRAW terbuka dengan dokumen, klik **Muat ulang** | daftar halaman terisi, status `Dokumen terdeteksi: ...`, tanpa dialog tambahan | NOT RUN |
| 13.2.3 | Buka dokumen A → Muat ulang → buka dokumen B di jendela yang sama → klik **Salin ke halaman terpilih** | dialog "Dokumen target berubah" muncul, operasi dibatalkan | NOT RUN |
| 13.2.4 | Pilih objek, centang 2 halaman tujuan, klik **Salin ke halaman terpilih** → **Ya** | isi dialog konfirmasi sama seperti sebelumnya (Dokumen/Halaman asal/Halaman tuju/Satuan); setelah selesai laporan hasil sama formatnya; **satu** Ctrl+Z membatalkan seluruh operasi | NOT RUN |
| 13.2.5 | Set satuan dokumen ke inch, ulangi 13.2.4 | satuan dokumen kembali ke inch setelah operasi; status satuan tercatat di laporan | NOT RUN |
| 13.2.6 | Batalkan dialog konfirmasi copy | status `Copy dibatalkan oleh pengguna.`, dokumen tidak berubah | NOT RUN |
| 13.2.7 | Jalankan **Select otomatis**, cakupan **Pindai semua halaman** | status progres `Memindai progres: halaman n/N | total objek cocok sementara: x.`; laporan per halaman sama formatnya; objek terpilih di halaman hasil pertama | NOT RUN |
| 13.2.8 | Klik **Pilih hasil berikutnya** berulang sampai akhir | tidak ada halaman yang terlewat/diulang; di akhir muncul "Anda sudah berada pada hasil terakhir" dan tidak kembali ke hasil pertama | NOT RUN |
| 13.2.9 | Ubah combo kategori setelah pemindaian, lalu klik **Pilih hasil berikutnya** | dialog "Hasil pemindaian tidak sesuai pilihan", navigasi berhenti | NOT RUN |
| 13.2.10 | Jalankan **Select otomatis** pada halaman tanpa objek yang cocok | seleksi lama **tidak** berubah, dialog "Seleksi yang ada tidak diubah." | NOT RUN |
| 13.2.11 | Buka **Watermark...** dari panel 01 | dialog watermark terbuka dan menerima halaman yang dicentang sebagai cakupan | NOT RUN |
| 13.2.12 | Klik tombol saat operasi panjang berjalan (klik ganda) | klik kedua diabaikan, status tetap "… masih berjalan", jendela tidak bisa ditutup | NOT RUN |

### 13.3 Aturan status

Ubah NOT RUN menjadi PASS **hanya** bila langkah di atas benar-benar
dijalankan dengan mata di CorelDRAW 2022. Semua langkah di §13.2 adalah
perilaku yang sudah ada sebelum refactor; kegagalan di sini berarti regresi.
