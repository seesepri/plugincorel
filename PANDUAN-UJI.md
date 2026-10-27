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
| Auto Select (exclude watermark) | `Program.vb` (`CollectMatchingShape`) | §4.4 |
| Busy guard | `Program.vb` + `OperationGate` | §4.5 |
| Watermark inti | `WatermarkEngine.vb` | §3, §4 |
| Model & matematika | `WatermarkLayout.vb` | §3.2, §3.4 |
| UI dialog | `WatermarkDialog.vb` | §4.6, §4.7 |
| Preset | `PresetStore.vb` | §4.7 |
| Log CSV | `BatchLog.vb` | §4.7 |
| Token per halaman | `WatermarkTokens.vb` | §4.8 |
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

## 14. Uji follow-up review PR #2 (BARU — status: **NOT RUN** di lingkungan audit)

Perbaikan 2026-10-27 menutup dua celah kebenaran status:
`CopyMeasurement` (jenis input & kegagalan `Count`) dan verifikasi penyimpanan
(`SaveOutputVerification`). Yang diuji otomatis tanpa CorelDRAW sudah PASS;
langkah di bawah ini **wajib** dijalankan di CorelDRAW 2022.

### 14.1 Bukti otomatis (sudah dijalankan)

| Pemeriksaan | Hasil |
|---|---|
| Regression test | `138 lulus, 0 gagal` pada putaran pertama; **kini 167 lulus, 0 gagal** (lihat §15.1) |
| Pemeriksaan statis | `237 pemeriksaan`, 0 gagal pada putaran pertama; **kini 279**, 0 gagal (lihat §15.1) |
| Kompilasi net48 | `Build succeeded, 0 warning, 0 error` |
| Build Release x64 Windows | **PASS** — CI run 40452213172: `bin\x64\Release\CorelPageCopier.exe`, 0 warning, 0 error |

### 14.2 Uji manual di CorelDRAW (status awal: NOT RUN)

| No | Langkah | Hasil yang benar | Status |
|---|---|---|---|
| 14.2.1 | Pilih **beberapa** objek di halaman sumber (bukan satu), lalu **Salin ke halaman terpilih** | laporan menyebut jumlah objek yang benar; semua objek ikut tersalin; status per halaman **Verified/Partial** sesuai kenyataan | NOT RUN |
| 14.2.2 | Pilih **satu** objek saja, ulangi copy | tetap seperti sebelumnya (tidak ada regresi pada Shape tunggal) | NOT RUN |
| 14.2.3 | Pilih objek yang berada di **layer terkunci** (atau objek yang tidak bisa dipilih) | status halaman **Partial/Failed** dengan alasan, bukan Verified | NOT RUN |
| 14.2.4 | Watermark batch, mode **file baru**, folder berisi file CDR | file `<nama>_wm.cdr` dibuat; ringkasan **OK**; pesan simpan menyebut `save OK (BARU)` | NOT RUN |
| 14.2.5 | Watermark batch, mode **timpa**, file sumber read-only / tidak bisa ditulis | file **tidak** berubah; status **GAGAL** dengan label `TIDAK BERUBAH`/`SAVE GAGAL`; **bukan** OK | NOT RUN |
| 14.2.6 | Setelah 14.2.5, periksa Object Manager & waktu modifikasi file | dokumen tidak berubah; tidak ada dokumen kedua yang terbuka di CorelDRAW (verifikasi tidak membuka ulang file) | NOT RUN |
| 14.2.7 | Lihat CSV log pada 14.2.4 & 14.2.5 | kolom Status = `OK` vs `GAGAL` sesuai kenyataan; tidak ada baris OK palsu | NOT RUN |

Ubah NOT RUN menjadi PASS hanya bila benar-benar diperiksa dengan mata di
CorelDRAW 2022.

---

## 15. Uji follow-up review KEDUA PR #2 (BARU — status: **NOT RUN** di lingkungan audit)

Putaran kedua memperketat dua hal: **bukti jenis input** pada `CopyMeasurement` dan
**bukti penyimpanan** pada `SaveOutputVerification` (bentuk berkas + sidik jari isi).
Laporan lengkap: `docs/REVIEW-PR2-FOLLOWUP-2.md`.

### 15.1 Bukti otomatis (sudah dijalankan, tanpa CorelDRAW)

| Pemeriksaan | Hasil |
|---|---|
| `python3 tools/static-checks.py` | **PASS** — 279 pemeriksaan, 0 gagal |
| Regression test | **PASS** — `167 lulus, 0 gagal` (138 test lama tetap lulus) |
| Kompilasi net48 x64 (`CompileCheck`) | **PASS** — 0 warning, 0 error |
| Build Release x64 via MSBuild Windows | **NOT RUN** di lingkungan audit — jalankan CI (`windows-latest`) dengan push commit |
| Integration CorelDRAW 2022 | **NOT RUN** |

### 15.2 Uji manual di CorelDRAW 2022 (status awal: NOT RUN)

| No | Langkah | Hasil yang benar | Status |
|---|---|---|---|
| 15.2.1 | Watermark batch, mode **file baru**, 3 file CDR 2022 asli | ketiga berkas `<nama>_wm.cdr` dibuat; semua status **OK**; label simpan `BARU`. (Membuktikan bentuk **ZIP** diterima — CorelDRAW 2022 menulis paket ZIP, bukan RIFF) | NOT RUN |
| 15.2.2 | Pilih **satu** objek kurva saja, lalu **Salin ke halaman terpilih** | status halaman **Verified** (bukan Partial). Bila Partial muncul, catat pesan `Count`/`Shape.Type` — berarti build ini melaporkan Shape tunggal secara tak terduga | NOT RUN |
| 15.2.3 | Pilih objek di **layer terkunci** / grup yang terkunci sebagian, ulangi copy | status **Partial/Failed** dengan alasan; tidak pernah Verified | NOT RUN |
| 15.2.4 | Watermark batch, mode **timpa**, berkas sumber diset **read-only** | berkas **tidak berubah**; status **GAGAL** (`TIDAK BERUBAH` atau `SAVE GAGAL`); **tidak ada** peringatan kerusakan sumber (karena berkas utuh) | NOT RUN |
| 15.2.5 | Watermark batch, mode **timpa**, hentikan CorelDRAW paksa tepat setelah menyimpan (simulasi simpan gagal) — atau pakai berkas di folder tanpa hak tulis lalu periksa berkas | bila berkas sumber menjadi kosong/hilang, ringkasan **wajib** memuat peringatan “berkas sumber kemungkinan TIDAK LENGKAP atau rusak”; status **GAGAL** | NOT RUN |
| 15.2.6 | Buka `<nama>_wm.cdr` hasil 15.2.1 **di CorelDRAW** | terbuka normal, watermark terlihat, jumlah halaman benar (bukti tidak ada klaim validitas palsu) | NOT RUN |
| 15.2.7 | Periksa CSV: baris file & baris ringkasan | setiap berkas muncul **satu** baris; teks ringkasan (`… OK n, SEBAGIAN n, GAGAL n`) muncul **tepat sekali** per baris ringkasan | NOT RUN |
| 15.2.8 | Selama dan sesudah 15.2.1 & 15.2.4, lihat jendela CorelDRAW | **tidak ada** dokumen kedua yang terbuka sendiri (verifikasi tidak membuka ulang berkas) dan dokumen aktif tidak berpindah | NOT RUN |

Ubah NOT RUN menjadi PASS hanya bila benar-benar diperiksa dengan mata di
CorelDRAW 2022 dan Windows. Build Release x64 Windows juga harus PASS di CI sebelum merge.
