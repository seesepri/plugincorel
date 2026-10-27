# Audit — Reliabilitas Batch Watermark & Validasi PDF Proof

Tanggal: 2026-10-27
Cakupan: `WatermarkEngine.vb`, `WatermarkLayout.vb`, `WatermarkTokens.vb`,
`BatchLog.vb`, `WatermarkDialog.vb`, plus modul baru `PdfProofValidator.vb`,
`WatermarkPlacementCheck.vb`, `BatchRunTracker.vb`, dan perubahan kecil di
`BatchFilePolicy.vb`.
Dokumen audit sebelumnya: `docs/AUDIT-RELIABILITY.md` (copy/undo/satuan),
`docs/AUDIT-AUTOSELECT.md` (Auto Select).

---

## 0. Ringkasan temuan

| # | Fokus | Temuan | Dampak | Status |
|---|---|---|---|---|
| A1 | Batch | Tidak ada jalur **berhenti** sama sekali; menutup jendela saat batch berjalan meninggalkan proses & laporan tak terkendali | Pengguna tidak bisa membatalkan batch panjang dengan aman | Diperbaiki |
| A2 | Batch | Exception di luar blok per-file (mis. handler UI `Progress` melempar) membuat **baris ringkasan CSV tidak pernah ditulis** dan sisa file hilang dari laporan | Jumlah OK/GAGAL tidak menjumlah ke jumlah file | Diperbaiki |
| A3 | Batch | Dokumen yang tidak bisa ditutup hanya dianggap "warning kecil", tanpa angka; tidak ada deteksi **kebocoran dokumen** | Memori CorelDRAW penuh pada batch besar | Diperbaiki |
| A4 | Batch | Status per file hanya OK/SEBAGIAN/GAGAL; file yang **tidak sempat diproses** tidak punya status | File hilang dari laporan | Diperbaiki (status `DIBATALKAN`) |
| B1 | Penempatan | Objek yang lebih lebar dari halaman (teks panjang, margin besar pada halaman kecil) menghasilkan koordinat **negatif** → watermark terpotong | Watermark salah/terpotong tetapi dilaporkan sukses | Diperbaiki (`ClampAxis`) |
| B2 | Penempatan | Mode tile sebelumnya bisa menempatkan pola **keluar halaman** (`Ceiling(page/step)`), sehingga baris/kolom terakhir terpotong | Watermark terpotong di tepi | Diperbaiki |
| B3 | Penempatan | Tidak ada verifikasi: setelah objek diletakkan, posisi/ukuran **tidak pernah dibaca kembali** | Kesalahan CorelDRAW (unit, rotasi, halaman landscape) tidak terdeteksi | Diperbaiki (`VerifyPlacement`) |
| B4 | Penempatan | Rotasi tidak punya model bounding box eksplisit; jarak tile memakai ukuran sebelum/sesudah rotasi tanpa dokumentasi | Sulit diuji, rawan tumpang-tindih | Diperbaiki (`RotatedBounds`, didokumentasikan) |
| C1 | PDF | Keberhasilan ekspor hanya dinilai dari "tidak ada exception" — bahkan file yang **tidak pernah dibuat** dilaporkan `PDF proof: <path>` | Klaim sukses tanpa bukti | Diperbaiki |
| C2 | PDF | Tidak ada pemeriksaan jumlah halaman, ukuran halaman, orientasi, header/`%%EOF`/`startxref` | PDF rusak/terpotong lolos sebagai sukses | Diperbaiki (`PdfProofValidator`) |
| C3 | PDF | Fallback `ExportEx` menelan pesan error (`Catch` kosong) | Penyebab kegagalan hilang | Diperbaiki |
| D1 | Logging | Pesan exception multi-baris/karakter kontrol bisa memecah baris CSV | Jumlah baris CSV ≠ jumlah file | Diperbaiki (`NormalizeCellText`) |
| D2 | Logging | Kegagalan menulis CSV (folder read-only) tidak dilaporkan ke pengguna | Pengguna mengira log ada | Diperbaiki |
| D3 | Logging | Ringkasan batch tidak bisa membedakan file **dibatalkan** dari **gagal** | Angka ringkasan menyesatkan | Diperbaiki |

Temuan yang **sudah** diperbaiki pada audit sebelumnya dan tetap dijaga di sini:
output default berupa file baru `nama_wm.cdr`, output lama (`*_wm.cdr`) tidak
dijadikan sumber, benturan nama output, validasi mode timpa, serta verifikasi
"file output ada & tidak 0 byte" (lihat `docs/AUDIT-RELIABILITY.md` §D/E).

---

## 1. Fokus A — batch yang aman

### 1.1 Penghentian batch (baru)

* `WatermarkEngine.RequestStop()` menyalakan penanda berhenti.
* Penanda diperiksa **di antara file** dan **di antara halaman**:
  file yang sedang diproses diselesaikan (tidak ada dokumen setengah jadi),
  halaman tersisa tidak disentuh → status file menjadi `SEBAGIAN` karena
  `PagesApplied < PagesIntended`.
* Sisa file **tidak hilang**: masing-masing mendapat outcome
  `DIBATALKAN` + baris CSV (`BatchRunTracker.BuildCancelledOutcome`).
* UI: tombol **"Hentikan"** di panel batch (aktif hanya saat batch berjalan).
  Tombol ini tidak mematikan proses secara paksa — ia meminta berhenti rapi.

### 1.2 Exception di tengah batch

Seluruh loop batch dibungkus `Try/Catch/Finally`:

* exception di luar blok per-file (mis. handler `Progress` milik UI melempar)
  dicatat sebagai `Errors`, sisa file dicatat `DIBATALKAN`;
* **baris ringkasan CSV selalu ditulis** di `Finally`, termasuk saat batch
  dihentikan atau gagal;
* `BatchRunTracker.IsConsistent()` memeriksa `OK + SEBAGIAN + GAGAL + DIBATALKAN
  = jumlah file yang direncanakan`; bila tidak, peringatan muncul di laporan.

### 1.3 File yang bermasalah

| Kondisi | Perilaku |
|---|---|
| File sedang terbuka / tidak dapat diakses | `OpenDocument` melempar → status `GAGAL` dengan alasan, batch lanjut |
| File rusak | idem |
| `SaveAs` gagal atau tidak menulis file | `GAGAL` (file tidak pernah dihitung sukses) |
| Output 0 byte | `GAGAL` |
| Dokumen gagal ditutup | pesan per file + peringatan; **jumlah dokumen terbuka** dibandingkan sebelum/sesudah batch (`DocumentsLeftOpen`) |

### 1.4 Output & nama file

Tidak diubah dari audit sebelumnya (sudah benar): default file baru, output lama
dilewati, benturan nama diselesaikan dengan akhiran unik, mode timpa divalidasi
ketat dan menolak file read-only / di luar folder.

---

## 2. Fokus B — akurasi watermark

### 2.1 Masalah yang diperbaiki

1. **Clamp halaman.** `WatermarkLayout.AnchorLeftBottom` sekarang menjepit
   posisi (`ClampAxis`) sehingga objek tidak pernah keluar halaman. Objek yang
   lebih besar dari halaman dipusatkan dan dilaporkan sebagai risiko clipping.
2. **Tile di dalam halaman.** Jumlah kolom/baris tile dihitung dari ruang yang
   benar-benar tersedia (`Floor((page − objek)/step) + 1`), bukan
   `Ceiling(page/step)`, sehingga tidak ada pola yang terpotong di tepi.
3. **Verifikasi penempatan.** Setelah objek diletakkan, `CenterX/CenterY/
   SizeWidth/SizeHeight` dibaca kembali dan dibandingkan dengan perhitungan
   (`WatermarkPlacementCheck.Verify`, toleransi 0,5 mm). Selisih → peringatan
   per halaman di laporan. Jumlah tile yang berhasil ditempati juga dibandingkan
   dengan jumlah posisi yang diharapkan.
4. **Rotasi.** `RotatedBounds(width, height, angle)` menghitung bounding box
   sejajar sumbu; dipakai sebagai model untuk menilai jarak tile/objek yang
   diputar (diuji: 45° dan 90°).
5. **Satuan.** Ukuran halaman dibaca setelah `Document.Unit` dinormalkan ke mm
   (perilaku lama dipertahankan), dan untuk perbandingan PDF dipakai
   `WatermarkPlacementCheck.PageSizeMillimeters` → `ToMillimeters` eksplisit
   (mm/inch/point), bukan asumsi mm. Satuan yang tidak dapat dikonversi
   menghasilkan NaN → tidak dipakai (bukan ditebak).

### 2.2 Yang diuji oleh unit test

* 3×3 anchor + margin 10 mm pada A4 (sembilan posisi diperiksa titiknya).
* Margin negatif → 0; margin besar pada halaman kecil tidak menghasilkan
  koordinat negatif.
* Objek lebih besar dari halaman → dipusatkan, `FitsOnPage` = False.
* Tile: tidak menumpuk, seluruh posisi berada di dalam halaman, objek
  raksasa tetap 1 posisi di tengah.
* Rotasi 0°/45°/90°/360° terhadap bounding box.
* Konversi mm ↔ inch ↔ point, dan hasil penempatan yang identik dari ukuran
  inch vs mm.
* Verifikasi: objek tepat → lolos; bergeser 2 mm → gagal; ukuran tak terbaca →
  tidak diklaim cocok.

### 2.3 Yang TIDAK bisa dibuktikan tanpa CorelDRAW

* Apakah CorelDRAW benar-benar menaruh objek di posisi yang diminta (yang
  diperiksa hanyalah **selisih yang dilaporkan**; bila CorelDRAW menyimpang,
  kini muncul peringatan, bukan klaim sukses).
* Tampilan visual opasitas 50%, rotasi, dan clipping pada dokumen nyata.
* Semua kombinasi teks+logo, tile pada halaman landscape, dan dokumen
  multi-ukuran → masuk daftar uji manual `PANDUAN-UJI.md` §12.

---

## 3. Fokus C — PDF proof

### 3.1 Alur baru

```
PublishToPDF(pdf)            ─┐
  gagal → ExportEx(pdf)      ─┤  dua pesan kegagalan DICATAT (tidak ada Catch kosong)
file tidak ada               → "PDF proof TIDAK dibuat" (Invalid)
PdfProofValidator.Validate   → header / %%EOF / startxref / jumlah halaman / MediaBox
status Valid                 → dilaporkan valid (halaman & ukuran cocok)
status Invalid / Unverifiable→ ProofPdfInvalid = True, status file = SEBAGIAN
```

### 3.2 Yang divalidasi otomatis (`PdfProofValidator`)

| Pemeriksaan | Bukti |
|---|---|
| Header `%PDF-1.x` | 1024 byte pertama |
| Trailer `%%EOF` | 2048 byte terakhir |
| `startxref` | 8192 byte terakhir |
| Jumlah halaman | objek `/Type /Page` (bukan `/Pages`) |
| Ukuran halaman | `/MediaBox`, dibandingkan dengan ukuran dokumen sumber (mm → pt, toleransi 1 pt) |
| Orientasi | lebar/tinggi, termasuk `/Rotate 90/270` |
| Ukuran berbeda per halaman | dibandingkan halaman per halaman |

### 3.3 Yang **tidak** diklaim

* Watermark benar-benar terlihat, opasitas, clipping, dan posisi visual —
  tidak dapat dibaca dari struktur PDF; karena itu dilaporkan sebagai
  **uji manual** (`PANDUAN-UJI.md` §12), bukan bagian dari klaim otomatis.
* PDF dengan objek terkompresi (`/ObjStm`) → status **Unverifiable** (bukan
  Valid), dengan instruksi cek manual.
* Kualitas cetak (warna, font tertanam, overprint) — di luar cakupan audit ini.

---

## 4. Fokus D — logging

* Satu sumber status: `BatchFileOutcome.Status` → dipakai oleh UI, ringkasan,
  dan baris CSV. Tidak ada lagi perhitungan ulang di UI.
* Status: `OK`, `SEBAGIAN`, `GAGAL`, `DIBATALKAN` (+ baris `RINGKASAN`).
* Kolom CSV: `Waktu; File; Halaman; Shape; Status; Pesan; Durasi_detik`
  (tidak berubah, kompatibel dengan log lama).
* `NormalizeCellText` mengganti CR/LF/TAB/NUL sehingga **satu baris log =
  satu baris fisik**; kutip ganda & pemisah `;` tetap di-escape RFC 4180.
* Kegagalan menulis CSV (folder read-only/tidak dapat diakses) → peringatan
  di laporan, batch tetap berjalan.
* Ringkasan memuat `OK n, SEBAGIAN n, GAGAL n, DIBATALKAN n` dan jumlah file
  yang direncanakan; konsistensi diperiksa oleh `BatchRunTracker`.

---

## 5. Bukti pengujian

| Jenis bukti | Hasil |
|---|---|
| Regression/unit test (net8.0, tanpa CorelDRAW) | **216 lulus, 0 gagal** (sebelumnya 157; **+59 test baru**) |
| Kompilasi produksi net48 (`Option Strict Off`, semua `CorelPageCopier/*.vb`) | **Build succeeded, 0 warning, 0 error** |
| Pemeriksaan statis (`tools/static-checks.py`) | **295 pemeriksaan lulus, 0 gagal** |
| Build Release x64 (MSBuild Windows) | **PASS** — CI run [40447818533](https://github.com/seesepri/plugincorel/actions/runs/40447818533) (commit `470ae08`): `CorelPageCopier -> bin\x64\Release\CorelPageCopier.exe`, **Build succeeded, 0 Warning, 0 Error**; static checks 295 PASS, test `216 lulus, 0 gagal` di Windows (angka identik dengan Linux) |
| Validasi PDF pada file PDF **nyata** hasil CorelDRAW | **NOT RUN** — tidak ada CorelDRAW di lingkungan ini |
| Pemeriksaan visual (posisi, opasitas, clipping, rotasi) | **NOT RUN** — butuh CorelDRAW + mata manusia |
| Uji cetak / kualitas siap cetak | **NOT RUN** — di luar cakupan, tidak diklaim |

Test baru:

| Suite | Jumlah | Isi |
|---|---|---|
| `PdfProofValidatorTests` | 16 | struktur PDF, jumlah halaman, ukuran, orientasi, object stream |
| `WatermarkPlacementTests` | 19 | anchor 3×3, clamp, rotasi, tile, satuan, verifikasi |
| `BatchLogFormatTests` | 10 | escaping CSV, newline, kontrol, ringkasan, konsistensi |
| `BatchReliabilityTests` | 12 | stop, exception, gagal buka/simpan/tutup, PDF valid/rusak/hilang/salah ukuran, CSV gagal |

Perintah:

```bash
dotnet run --project tests/CorelPageCopier.RegressionTests -c Release
python3 tools/static-checks.py
msbuild CorelPageCopier\CorelPageCopier.vbproj /p:Configuration=Release /p:Platform=x64   # Windows
```

---

## 6. Keterbatasan yang masih tersisa

1. **Belum ada bukti PDF nyata.** Validator diuji dengan PDF sintetis. PDF
   hasil CorelDRAW 2022 harus diuji manual (langkah di `PANDUAN-UJI.md` §12);
   bila CorelDRAW memakai object stream, status akan `Unverifiable` dan itu
   berarti "perlu cek manual", bukan "valid".
2. **Penghentian bersifat kooperatif.** File yang sedang diproses tetap
   diselesaikan; tidak ada pembatalan paksa di tengah `SaveAs`.
3. **Deteksi kebocoran dokumen** bergantung pada `Application.Documents.Count`;
   bila nilai itu tidak terbaca (`-1`), pemeriksaan dilewati dan hal ini tidak
   diklaim sebagai "tidak ada kebocoran".
4. **Verifikasi penempatan** hanya seakurat properti yang dilaporkan
   CorelDRAW (`CenterX/CenterY/SizeWidth/SizeHeight`). Kesalahan visual yang
   tidak tercermin di properti itu tidak terdeteksi.
5. **Kualitas cetak** (font tertanam, overprint, profil warna) tidak diperiksa.
6. **Log CSV** tetap satu baris per file (bukan per halaman); jumlah halaman
   dilaporkan di kolom `Halaman` + pesan (`n/m diterapkan`, `k halaman gagal`).
7. **Token** (`{page}`, `{pages}`, `{file}`, `{order}`, `{date}`, `{time}`)
   sudah punya test sendiri dari audit sebelumnya; audit ini tidak mengubah
   perilakunya dan tidak mengklaim ulang pengujian visualnya.
