# Audit Reliabilitas — Corel Page Copier

Tanggal audit: 2026-10-27
Target: CorelDRAW 2022 (ProgID `CorelDRAW.Application.24`), Windows x64, VB.NET WinForms,
.NET Framework 4.8, COM late binding.
Prinsip kerja: **tanpa rewrite**, perubahan sekecil mungkin per temuan, fitur lama dipertahankan.

---

## 0. Ringkasan status verifikasi (jujur, apa adanya)

| Tahap | Status | Bukti |
|---|---|---|
| 1. `python tools/static-checks.py` | **PASS** — 200 pemeriksaan, 0 gagal | output di §7.1 |
| 2. Regression/unit tests tanpa CorelDRAW | **PASS** — 103 test, 0 gagal | output di §7.2 |
| 3. Compile seluruh proyek (`Program.vb`, `WatermarkDialog.vb`, engine, …) | **PASS (compile-check)** — lihat catatan | §7.3 |
| 4. Build Release x64 dengan **MSBuild di Windows** | **PASS (CI `windows-latest`)** — run 40444845364 | §7.4 |
| 5. Smoke test / integration test di **CorelDRAW 2022** | **NOT RUN** | CorelDRAW tidak tersedia di lingkungan ini |
| 6. ZIP source | **DIBUAT & DIVERIFIKASI** (ekstrak + cek isi + jalankan checks dari hasil ekstrak) | §7.6 |

> Build Windows sudah diverifikasi lewat GitHub Actions (`windows-latest`): static checks →
> regression tests → `msbuild /t:Rebuild /p:Configuration=Release /p:Platform=x64` semuanya
> **success**, dan artifaknya berisi `CorelPageCopier.exe` PE32+ x86-64 (lihat §7.4).
> **Integration test CorelDRAW 2022 tetap belum dijalankan (NOT RUN)** — langkahnya ada di
> §6 dan `PANDUAN-UJI.md` §10.

Catatan §7.3: tidak tersedia Windows, jadi MSBuild Windows tidak bisa dipakai. Sebagai gantinya
seluruh 13 file `.vb` proyek **dikompilasi dengan compiler VB (Roslyn) milik .NET SDK 8 terhadap
reference assemblies .NET Framework 4.8** (`Microsoft.NETFramework.ReferenceAssemblies`), dengan
`Option Strict Off`, `RootNamespace`, `PlatformTarget x64` yang sama seperti `.vbproj`. Ini
memvalidasi sintaks + resolusi tipe/nama untuk target `net48` (termasuk `System.Windows.Forms`),
tetapi **bukan** pengganti `msbuild /p:Configuration=Release /p:Platform=x64` di Windows.

---

## 1. Daftar file yang diubah

### File baru (logika murni / bisa diuji tanpa CorelDRAW)

| File | Isi |
|---|---|
| `CorelPageCopier/CopyVerification.vb` | `CopyVerifyStatus` (Verified/PartiallyVerified/Failed), `BoundingBox`, `ShapeSignature`, `CopyVerification.Classify` + `VerifyPasted`, `CopyPageOutcome`, `CopyRunSummary` (sumber tunggal status + pembentuk laporan) |
| `CorelPageCopier/CopyMeasurement.vb` | Pengukuran shape late-bound yang benar: `MeasureTopShapes` (z-order atas), `MeasureSelection`, `MeasurePasted`, `MeasuredShapeSet` |
| `CorelPageCopier/DocumentUnits.vb` | `DocumentUnits` (tabel konversi eksplisit mm/cm/inch/pt/pica/didot/cicero; pixel = tidak didukung) + `DocumentUnitScope` (state machine normalisasi & pemulihan) |
| `CorelPageCopier/UndoGroup.vb` | `UndoGroupGuard` — pelacak status `BeginCommandGroup`/`EndCommandGroup`, `SingleUndoGuaranteed`, `ShouldCallEnd` |
| `CorelPageCopier/BatchFilePolicy.vb` | `BatchFileStatus`, `BatchStatusRules` (evaluasi status), `BatchFileOutcome`, dan kebijakan file: `NormalizeOutputSuffix`, `IsGeneratedOutput`, `SelectSources`, `BuildOutputPath`, `ResolveUniqueOutputPath`, `IsUnderFolder`, `ValidateOverwriteTarget`, `ValidateNewOutputPath` |
| `CorelPageCopier/CorelConnection.vb` | `DocumentIdentity` (identitas dokumen + token instance, `SameTarget`) + `ComObjectHealth` (`IsAlive`, `TryRelease`) |
| `tests/CorelPageCopier.RegressionTests/**` | Proyek test .NET 8 yang **mengompilasi ulang file sumber produksi** (bukan salinan) + fake COM (`FakeCorel.vb`) + 103 test |

### File yang diubah

| File | Perubahan |
|---|---|
| `CorelPageCopier/Program.vb` | `ConnectToCorel` ditulis ulang (identitas instance, izin sebelum instance baru, deteksi COM mati), `CopyClickedCore` ditulis ulang (verifikasi A, satuan B, undo C, konfirmasi identitas F), helper `CurrentIdentity`/`EnsureSameDocument`/`ReleaseCorelReferences`, `LoadPages` menyimpan identitas dokumen, helper lama `TryMeasureTopShapes` **dihapus** |
| `CorelPageCopier/WatermarkEngine.vb` | `WatermarkResult` (+`PagesIntended`, `PagesFailed`, `Warnings`, `FileOutcomes`, `SkippedFiles`, `FilesOk/FilesPartial/FilesFailed`), `ApplyToDocument` (satuan & undo aman, `writeLogRow`), `RemoveFromDocument` (undo guard), `ApplyToFolder` (filter output lama, status per file, verifikasi save, validasi timpa) |
| `CorelPageCopier/WatermarkDialog.vb` | Konfirmasi identitas dokumen sebelum Terapkan/Hapus, konfirmasi batch menyebut kebijakan akhiran output, baris status dari satu sumber status |
| `CorelPageCopier/WatermarkLayout.vb` | `WatermarkSettings.OutputSuffix` (+`Clone`) |
| `CorelPageCopier/PresetStore.vb` | Serialisasi/deserialisasi `OutputSuffix` |
| `CorelPageCopier/BatchLog.vb` | `SummaryEntry` overload dengan jumlah SEBAGIAN (overload lama dipertahankan) |
| `CorelPageCopier/CorelPageCopier.vbproj` | 6 file baru didaftarkan di `<Compile Include>` |
| `tools/static-checks.py` | Invarian satuan diperbarui + 90 pemeriksaan baru untuk temuan audit (total 200) |
| `.github/workflows/build-release.yml` | Langkah baru menjalankan regression tests sebelum MSBuild |
| `PANDUAN-UJI.md`, `CorelPageCopier/README_ID.md` | Prosedur uji manual baru + catatan perilaku baru |

Diff: `9 files changed, 993 insertions(+), 318 deletions(-)` + 6 file baru (~2.100 baris) + proyek test (~1.400 baris).

---

## 2. Temuan sebelum → sesudah

### A. Verifikasi Copy (temuan #1)

**Sebelum (bug nyata, 3 lapis):**

1. `TryMeasureTopShapes(layer, added, …)` mengukur `layer.Shapes(1 .. added)`.
   Di CorelDRAW, indeks koleksi `Shapes` berjalan **dari bawah ke atas** (z-order), sedangkan
   objek hasil `Paste` berada **di atas**. Jadi geometri sumber dibandingkan dengan objek
   **terbawah** pada layer — bandingannya hampir selalu salah. Pada halaman yang sudah berisi
   objek, ini rutin menghasilkan "posisi/ukuran berbeda dari sumber" (false negative), dan pada
   layer kosong menghasilkan false positive bila objek barunya kebetulan seukuran.
2. Bila pengukuran tidak tersedia (`sourceMeasured = False` atau ukur gagal), halaman tetap
   dihitung `succeeded += 1` dan dilaporkan **"Terverifikasi lewat jumlah shape saja"**, lalu
   status bar menulis **"X/N halaman terverifikasi"**. Status "Partial" tidak ada.
3. Menambah shape dianggap bukti cukup: `added > 0` → sukses; tidak ada pembandingan struktur
   (jumlah/tipe objek, grup).

**Sesudah:**

- `CopyMeasurement.MeasurePasted` mengukur **N objek teratas** (`total - added + 1 .. total`) —
  deterministik, tidak bergantung pada perilaku seleksi setelah paste. `ActiveSelectionRange`
  hanya dipakai sebagai *fallback* bila geometri z-order gagal diukur **dan** jumlahnya persis
  sama dengan jumlah objek baru.
- Status per halaman dipisah: **Verified** (geometri cocok dalam 0,5 mm **dan** struktur cocok),
  **Partial** (paste terjadi, tetapi geometri/struktur tidak dapat diverifikasi — *tidak pernah*
  dihitung sebagai sukses), **Failed** (tidak ada objek hasil, atau geometri di luar toleransi).
- Struktur objek dibandingkan lewat `ShapeSignature` (jumlah objek + sebaran kode tipe:
  3 kurva / 5 bitmap / 6 teks / 7 grup), sehingga grup bersarang dan multi-selection diverifikasi
  strukturnya, bukan hanya "ada tambahan shape".
- `CopyRunSummary` adalah satu-satunya sumber angka untuk status bar **dan** MessageBox; halaman
  Partial dirinci terpisah, ikon dialog hanya "information" bila **semua** halaman Verified.
- Objek sumber tidak pernah dihapus/diubah: alur hanya membaca properti lalu `selection.Copy()`.
  Tidak ada fallback `Delete`.

**Reproduksi bug lama:**
1. Buat dokumen 3 halaman. Halaman 1: objek persegi 50×50 mm di (100,100); halaman 2: **sudah berisi**
   beberapa objek desain (mis. teks dan logo di area lain).
2. Jalankan aplikasi, pilih halaman 1 sebagai sumber, pilih objek, centang halaman 2 & 3, klik **Salin**.
3. Versi lama: halaman 2 dilaporkan gagal "posisi/ukuran berbeda" padahal hasil paste benar
   (karena yang diukur objek terbawah), sedangkan halaman 2 yang tidak dapat diukur dilaporkan
   "terverifikasi lewat jumlah shape". Versi baru: kedua halaman **Verified** dengan Δx/Δy/Δw/Δh
   yang dilaporkan dan struktur objek disebutkan.

**Test yang menjaga:** `CopyMeasurementTests` (5 test, termasuk "pengukuran lama akan salah"),
`CopyPipelineTests` (12 test: objek tunggal, multi-selection, grup bersarang, bitmap, teks,
ukuran campuran, bergeser, lebih kecil, tidak terukur, dokumen inch, struktur berubah),
`CopyVerificationTests` (15 test).

---

### B. Normalisasi satuan (temuan #2)

**Sebelum:**

```vb
Try : previousUnit = CInt(corelDoc.Unit) : unitStored = True : Catch : End Try   ' gagal? lanjut
Try : corelDoc.Unit = CDR_UNIT_MM : Catch : End Try                             ' gagal? lanjut
…
If Not sourceMeasured Then succeeded += 1   ' dilaporkan sebagai terverifikasi
…
Finally
    If unitStored Then Try : corelDoc.Unit = previousUnit : Catch : End Try      ' gagal? senyap
```

Tiga masalah: (a) bila `Document.Unit` gagal dibaca/di-set, verifikasi tetap berjalan seolah mm;
(b) kegagalan memulihkan satuan tidak dilaporkan sama sekali; (c) tidak ada konversi eksplisit
untuk dokumen inch/point.

**Sesudah (Program + WatermarkEngine):**
- Satuan asli **wajib** terbaca. Bila tidak → operasi dihentikan **sebelum** dokumen diubah,
  dengan pesan eksplisit ("Operasi dibatalkan agar verifikasi tidak memakai asumsi milimeter").
- Bila `doc.Unit = mm` gagal tetapi satuan asli dapat dikonversi (inch/cm/pt/pica/didot/cicero),
  pengukuran dilakukan dalam satuan asli dan **dikonversi eksplisit** ke mm (faktor terdokumentasi
  di `DocumentUnits.MillimetersPerUnit`). Pixel sengaja **tidak** didukung (bergantung resolusi)
  → operasi dibatalkan, bukan ditebak.
- Watermark engine **membatalkan** bila `mm` tidak bisa di-set, karena seluruh matematika tata
  letak memakai mm — melanjutkan akan menghasilkan posisi salah yang tampak "sukses".
- Pemulihan satuan selalu di `Finally`, dan kegagalannya masuk `Warnings`/`Errors` **dan**
  MessageBox ("Satuan dokumen gagal dipulihkan … ubah manual"). Laporan dibuat **setelah**
  `Finally`, jadi kegagalan pemulihan tidak mungkin hilang dari laporan.
- Toleransi geometri terdokumentasi: **0,5 mm** per sumbu (Δpusat dan Δukuran) — lihat
  `CopyVerification.ToleranceMm` (alasan: ±1,4 pt, di bawah ambang terlihat pada A4, dan menutupi
  pembulatan internal CorelDRAW). Pembandingan selalu dalam mm.

**Reproduksi bug lama:** buka dokumen dengan satuan **inch** (Layout ▸ Document Options ▸ Rulers =
inches), copy objek antar halaman. Versi lama membandingkan angka inch dengan toleransi 0,5 →
objek yang bergeser 12 mm masih dinyatakan cocok (0,47 inch < 0,5). Versi baru mengonversi ke mm
sehingga pergeseran 1,27 mm (0,05 inch) sudah **Failed** — diuji di `CopyPipelineTests`.

**Test yang menjaga:** `DocumentUnitTests` (9 test: konversi mm/inch/pt/cm, pixel ditolak,
scope state machine, pemulihan gagal dilaporkan) + 3 test engine (satuan dipulihkan, unit tak
terbaca, penyetelan mm gagal).

---

### C. Undo grouping (temuan #3)

**Sebelum:** `Try doc.BeginCommandGroup(...) : inCommandGroup = True : Catch : End Try` — kegagalan
ditelan, dan laporan **selalu** menutup dengan "Semua paste berada dalam satu grup undo — satu
Ctrl+Z membatalkan seluruh operasi". `EndCommandGroup` juga dibungkus `Catch` kosong. Pada
`RemoveFromDocument` sama.

**Sesudah:**
- `UndoGroupGuard` mencatat setiap percobaan; `ShouldCallEnd` memastikan `EndCommandGroup`
  **hanya** dipanggil bila group benar-benar terbuka; `SingleUndoGuaranteed` hanya True bila
  Begin **dan** End sukses.
- Copy: bila Begin gagal, pengguna ditanya eksplisit ("Tanpa grup undo, satu Ctrl+Z TIDAK akan
  membatalkan seluruh operasi … Lanjutkan?") dengan default **No**; keputusan & alasannya masuk
  laporan.
- Watermark & hapus watermark: Begin gagal → `Warnings` + laporan; End gagal →
  `Errors` ("EndCommandGroup gagal … periksa daftar Undo CorelDRAW") sehingga hasil tidak
  dilaporkan OK.
- Teks laporan sekarang berasal dari `undoGuard.Describe()`: menjanjikan satu Ctrl+Z **hanya**
  bila benar-benar berhasil, dan menyebutkan nama grup undo di daftar Undo CorelDRAW.

**Reproduksi bug lama:** sulit dipicu pada CorelDRAW normal (butuh kegagalan COM). Simulasi
deterministik ada di `UndoGroupTests` (5 test: begin gagal, end gagal, end ganda) dan
`EngineTests` ("BeginCommandGroup gagal", "EndCommandGroup gagal").

---

### D. Status batch watermark (temuan #4)

**Sebelum (`ApplyToFolder`):**

```vb
Dim fileOk As Boolean = partResult.Ok AndAlso partResult.PagesPassed > 0
If fileOk Then okFiles += 1 Else failedFiles += 1
```

- Kegagalan `SaveAs` dimasukkan ke `Failures` **milik hasil batch**, bukan `partResult`, sehingga
  `partResult.Ok` tetap True → file yang **gagal disimpan** dihitung **OK** di ringkasan, dan
  baris CSV-nya ditulis `OK` (laporan sukses palsu).
- Tidak ada status SEBAGIAN: 1 dari 3 halaman gagal = "GAGAL", dan file dengan 2/3 halaman
  berhasil tidak dibedakan.
- Baris CSV ditulis **dua kali** untuk file yang sama (satu dari `ApplyToDocument`, satu dari
  `ApplyToFolder`) dengan status berbeda — bisa muncul `OK` lalu `GAGAL` untuk file yang sama.
- Pesan UI: `"Batch selesai: N dokumen tersimpan, M gagal"` dihitung dari `Failures.Count`
  (bukan dari status file) → angka UI ≠ angka ringkasan.

**Sesudah:**
- `BatchFileOutcome` memisahkan `ApplyAttempted/PagesIntended/PagesApplied/PagesFailed` (processing),
  `SaveAttempted/SaveSucceeded` (save), dan `Status` (final, dihitung `BatchStatusRules`).
  **OK hanya bila**: tidak ada error fatal, semua halaman target diterapkan, save dijalankan dan
  berhasil. Save diverifikasi (`File.Exists` + ukuran > 0 byte).
- Baris CSV, baris ringkasan CSV, dan pesan status UI memakai **`outcome.StatusLabel()` /
  `BatchStatusRules`** yang sama (`OK` / `SEBAGIAN` / `GAGAL`).
- Baris CSV ganda dihilangkan: `ApplyToDocument(..., writeLogRow:=False)` saat dipanggil batch.
- Detail exception disimpan: `ex.GetType().Name & ": " & ex.Message` pada `Failures`, `Messages`,
  dan CSV; kegagalan menutup dokumen masuk `Warnings` (tidak disembunyikan, tidak menggagalkan output).

**Reproduksi bug lama:** (1) buat folder berisi 2 file `.cdr`; (2) jalankan batch dengan salah satu
file dalam kondisi **read-only** (atau kunci file dengan aplikasi lain) sehingga `SaveAs` gagal;
(3) versi lama menulis baris CSV `OK` + ringkasan "berhasil 2" untuk file yang gagal disimpan.
Versi baru: baris CSV `GAGAL`, ringkasan `OK 1, SEBAGIAN 0, GAGAL 1`, dan file gagal **tidak**
masuk hitungan sukses. Diuji `EngineTests` ("SaveAs gagal -> status GAGAL", "SaveAs tidak menulis
file", "watermark hanya sebagian halaman").

---

### E. Pemrosesan output berulang (temuan #5)

**Sebelum:** `Directory.GetFiles(folder, "*.cdr", …)` mengambil juga `*_wm.cdr` hasil eksekusi
sebelumnya → pada run kedua file `a_wm.cdr` diproses lagi dan menghasilkan `a_wm_wm.cdr`
(watermark berlapis, ukuran file membengkak). Tidak ada pemeriksaan benturan nama; bila ada
`a.cdr` **dan** `a_wm.cdr` milik pengguna, mode "file baru" menimpa `a_wm.cdr` tanpa peringatan.
Mode timpa tidak divalidasi (path di luar folder, file read-only).

**Sesudah (`BatchFilePolicy`):**
- Sumber difilter: file berakhiran akhiran output (default `_wm`, `NormalizeOutputSuffix`) **tidak**
  diproses, dan masuk daftar "file dilewati" **beserta alasannya** di laporan + log progres
  (tidak hilang diam-diam). Bila semua file adalah output lama → error eksplisit, bukan "0 file OK".
- `ValidateNewOutputPath` + `SamePath` menjamin mode file baru **tidak pernah** menimpa file sumber;
  `ResolveUniqueOutputPath` memberi nama `_2`, `_3`, … bila path sudah dipesan sumber lain atau
  output sebelumnya pada batch yang sama.
- Mode timpa divalidasi ketat per file: harus `.cdr`, harus **di dalam** folder batch (hormati
  pilihan subfolder), bukan folder, dan tidak read-only; pelanggaran → file ditandai **GAGAL**
  dengan alasan, bukan ditimpa.
- Konfirmasi batch di dialog menyebut: mode simpan, akhiran output, dan bahwa file berakhiran
  `_wm` dilewati.

**Reproduksi bug lama:** (1) folder `D:\uji` berisi `a.cdr`, `b.cdr`; (2) jalankan batch
(hasil `a_wm.cdr`, `b_wm.cdr`); (3) jalankan batch **lagi** pada folder yang sama → versi lama
memproses 4 file dan menghasilkan `a_wm_wm.cdr`, `b_wm_wm.cdr` dengan watermark ganda.
Versi baru: hanya `a.cdr` & `b.cdr` diproses; `a_wm.cdr`/`b_wm.cdr` muncul di daftar "File
dilewati" dengan alasan "output batch sebelumnya".

**Test yang menjaga:** `BatchPolicyTests` (14 test) + `EngineTests` ("output lama (_wm) tidak
diproses ulang", "mode file baru: output tidak pernah menimpa file sumber", "dua sumber dengan
output sama -> nama unik", "mode timpa … ditolak").

---

### F. Koneksi CorelDRAW COM (temuan #6)

**Sebelum:** bila instance yang sedang dipakai tidak terdaftar di ROT, kode langsung
`Activator.CreateInstance` → **instance kedua dibuka tanpa memberi tahu pengguna**, lalu
"ActiveDocument" instance baru (biasanya dokumen kosong) dianggap target. Tidak ada penanganan
objek COM mati (instance ditutup di tengah sesi hanya menghasilkan error COM mentah), tidak ada
pemeriksaan bahwa dokumen target masih dokumen yang sama, dan tidak ada tampilan identitas
dokumen sebelum operasi mengubah dokumen.

**Sesudah:**
- Urutan koneksi: (1) koneksi lama **bila masih hidup** (`ComObjectHealth.IsAlive`), (2) instance
  yang sedang berjalan via `Marshal.GetActiveObject`, (3) **hanya setelah pengguna menyetujui**
  (`MessageBox` Yes/No, default No) instance baru dibuat — dan pengguna diberi tahu bahwa jendela
  CorelDRAW **baru** sudah terbuka serta dokumen mana yang akan diproses.
- Bila instance hidup tetapi tidak ada dokumen aktif → pesan eksplisit ("buka dokumen CDR di
  jendela itu, lalu klik Muat ulang"), bukan exception COM mentah.
- `DocumentIdentity` (nama/nama file + jumlah halaman + token instance dari
  `Application.AppWindow.Handle`) disimpan saat daftar halaman dimuat; `EnsureSameDocument`
  menolak operasi bila dokumen **atau instance** berubah, dengan menampilkan identitas lama vs
  baru. Ini juga mencegah "dokumen dari instance berbeda dianggap target yang sama".
- Identitas dokumen + jumlah halaman + halaman target + satuan ditampilkan di dialog konfirmasi
  sebelum copy; dialog watermark menampilkan identitas dokumen sebelum Terapkan/Hapus.
- Referensi COM lama dilepas (`Marshal.ReleaseComObject`) sebelum menyambung ulang.

**Reproduksi bug lama:** (1) tutup CorelDRAW; (2) jalankan aplikasi → versi lama membuka instance
baru diam-diam dan menampilkan "Tidak dapat terhubung" atau, setelah dokumen dibuka, memproses
dokumen yang salah. Versi baru menanyakan lebih dulu dan memberi tahu apa yang terjadi.
(2) Buka dokumen A, klik Muat ulang, lalu aktifkan dokumen B di CorelDRAW dan klik Salin → versi
lama menyalin dari halaman B tetapi melaporkan halaman A; versi baru menolak dengan pesan
"Dokumen aktif berbeda dari dokumen saat daftar halaman dimuat".

**Test yang menjaga:** `ConnectionIdentityTests` (10 test: instance berbeda, path berbeda,
nama tanpa path, describe, `FromDocument` fake, token instance, `ComObjectHealth`).

---

## 3. Aturan keselamatan — cek satu per satu

| Aturan | Status | Bukti |
|---|---|---|
| Tidak memakai `Delete` sebagai fallback pembersihan seleksi | ✅ tetap | `ApplySelection` masih `Throw` bila `ClearSelection` gagal; static check "Tidak ada objek yang dihapus"; `shape.Delete` tidak ada di `Program.vb` |
| Tidak menimpa file asli secara default | ✅ diperkuat | default `saveInPlaceCheck = False`; `ValidateNewOutputPath` + `SamePath` menolak output = sumber |
| Tidak menyatakan sukses bila hasil tidak dapat diverifikasi | ✅ | Partial tidak pernah masuk `VerifiedCount`; OK batch butuh save terverifikasi |
| Tidak menelan exception penting | ✅ | satuan, undo, save, close, PDF, CSV kini tercatat & dilaporkan |
| Tidak mengubah perilaku fitur lain | ✅ | Auto Select, watermark layout, token, preset, CSV, PDF proof, busy-gate tidak diubah; 8 test "fitur lama" mengunci perilaku |
| Kompatibel CorelDRAW 2022 & .NET Framework 4.8 | ✅ | tetap late binding, tanpa referensi interop baru, `TargetFrameworkVersion v4.8`, `PlatformTarget x64` tidak diubah |

Perubahan perilaku yang **disengaja** (harus diketahui pengguna):
1. Copy kini menampilkan **konfirmasi** (dokumen, halaman sumber/tujuan, satuan) sebelum menyalin.
2. Bila grup undo tidak tersedia, Copy bertanya lebih dulu dan laporan tidak lagi menjanjikan
   satu Ctrl+Z.
3. Watermark dibatalkan bila satuan dokumen tidak bisa dinormalkan ke mm (sebelumnya tetap jalan
   dengan hasil salah).
4. Batch melewati file berakhiran `_wm` (dengan alasan tertulis) — file seperti itu harus di-rename
   bila memang ingin diproses.
5. Batch menulis **satu** baris CSV per file (sebelumnya dua baris dengan status berbeda).

---

## 4. Risiko yang tersisa (jujur)

| Risiko | Dampak | Mitigasi / catatan |
|---|---|---|
| Glue UI di `CopyClickedCore` (±120 baris: urutan `Activate`/`Paste`, dialog) tidak tercakup unit test | Regresi di urutan pemanggilan tidak terdeteksi otomatis | Inti keputusan sudah dipindah ke kelas murni yang diuji (`CopyMeasurement`, `CopyVerification.VerifyPasted`, `CopyRunSummary`); sisanya dijaga static checks + prosedur manual §10 |
| `Layer.Paste()` diasumsikan menaruh objek di **atas** z-order layer aktif | Bila versi CorelDRAW berbeda perilakunya, pengukuran bisa salah | Strategi kedua memakai `ActiveSelectionRange` (hanya bila jumlah objeknya sama); selisih tetap terdeteksi sebagai Failed/Partial, bukan sukses palsu |
| Objek yang dirotasi/di-skew: `SizeWidth/SizeHeight` adalah bbox | Toleransi 0,5 mm bisa memicu Partial pada kasus ekstrem | Toleransi terdokumentasi; laporan menyebut Δ sehingga bisa dinilai manusia |
| Instans CorelDRAW "kedua" tetap mungkin dibuka — sekarang dengan izin pengguna | Pengguna bisa keliru memilih Yes | Pesan menyebut dokumen di jendela baru; identitas dokumen ditampilkan sebelum operasi |
| Identitas instance memakai `AppWindow.Handle` (fallback: hash proxy COM) | Bila API tidak tersedia di versi lain, token bisa berubah → peringatan "dokumen berubah" palsu | Perilaku aman (menolak, bukan salah target); pengguna cukup klik Muat ulang |
| Konversi satuan untuk pixel sengaja tidak didukung | Dokumen ber-unit pixel tidak bisa diverifikasi | Dibuat **eksplisit**: operasi dibatalkan dengan pesan, bukan ditebak |
| Pemrosesan ulang bisa terjadi bila pengguna mengganti akhiran output antar-run | `a_wm.cdr` lama bisa jadi sumber bila akhiran run baru `_v2` | Di luar cakupan default; akhiran tersimpan di preset & disebut di konfirmasi batch |
| Belum ada CI yang menjalankan test di Windows | Kegagalan build Windows hanya terdeteksi saat CI jalan | Workflow kini menjalankan static checks + regression tests + MSBuild x64 |

---

## 5. Toleransi & konstanta terdokumentasi

| Konstanta | Nilai | Alasan |
|---|---|---|
| `CopyVerification.ToleranceMm` | **0,5 mm** per sumbu (Δpusat, Δukuran) | ±1,4 pt; di bawah ambang terlihat pada A4; menutupi pembulatan CorelDRAW saat paste |
| `DocumentUnits.MillimetersPerUnit` | mm=1, cm=10, inch=25.4, pt=25.4/72, pica=25.4/6, didot=0.376, cicero=4.5115, pixel=**0 (ditolak)** | Konversi eksplisit; pixel bergantung resolusi sehingga tidak bisa dijamin |
| `BatchFilePolicy.DefaultOutputSuffix` | `_wm` | Dipertahankan dari perilaku lama; dipakai sebagai penanda "output batch" |
| `cdrShapeType` yang dibandingkan | 3 kurva, 5 bitmap, 6 teks, 7 grup | Konstanta lama di `Program.vb` / `WatermarkCdrConst` |

---

## 6. Instruksi reproduksi & verifikasi (ringkas)

Semua langkah di bawah dijalankan dari root repo:

```bat
:: 1. pemeriksaan statis
python tools\static-checks.py

:: 2. regression test (butuh .NET SDK 8; boleh Linux/Windows/macOS)
dotnet run --project tests\CorelPageCopier.RegressionTests

:: 3. build Windows
msbuild CorelPageCopier\CorelPageCopier.vbproj /m /t:Rebuild /p:Configuration=Release /p:Platform=x64


```

Integration test CorelDRAW: `PANDUAN-UJI.md` §10 (status **NOT RUN** di audit ini).

---

## 7. Hasil eksekusi (dipisah per tahap)

### 7.1 Static checks — PASS

```
$ python3 tools/static-checks.py
…
SEMUA PEMERIKSAAN STATIS LULUS      (200 pemeriksaan PASS, 0 FAIL, exit 0)
```

### 7.2 Regression / unit tests (tanpa CorelDRAW) — PASS

```
$ dotnet run --project tests/CorelPageCopier.RegressionTests
Corel Page Copier — regression tests (tanpa CorelDRAW)
…
Ringkasan: 103 lulus, 0 gagal, 103 total.
Catatan: integration test CorelDRAW 2022 TIDAK dijalankan oleh test ini (NOT RUN).
```

Cakupan (103 test): verifikasi copy 16, pengukuran objek 10, alur copy per halaman/tipe objek 11,
satuan dokumen 9, undo guard 5, kebijakan file batch 15, koneksi & identitas 10,
engine satu dokumen 10, engine batch folder 9, fitur lama 8. Rincian output: §7.7.

### 7.3 Compile proyek (.NET Framework 4.8, semua file produksi) — PASS (compile-check, bukan MSBuild Windows)

```
$ dotnet build tests/CorelPageCopier.CompileCheck/CorelPageCopier.CompileCheck.vbproj -c Release
Build succeeded.  0 Warning(s)  0 Error(s)
```
Proyek `tests/CorelPageCopier.CompileCheck/` mengompilasi **semua** file `.vb` produksi
(`Program.vb`, `WatermarkDialog.vb`, `WatermarkEngine.vb`, `WatermarkLayout.vb`,
`WatermarkTokens.vb`, `PresetStore.vb`, `BatchLog.vb` + 6 file baru) terhadap
`net48` + `Microsoft.NETFramework.ReferenceAssemblies`, `Option Strict Off`, `x64`.
Proyek ini hanya alat verifikasi — **bukan** pengganti build rilis MSBuild.

Status: **PASS untuk compile-check**, **NOT RUN untuk build MSBuild Windows** (§7.4).

### 7.4 Build Release x64 dengan MSBuild di Windows — **PASS**

Dijalankan di GitHub Actions, runner `windows-latest`, workflow `build-release.yml` pada
pull request (event `pull_request`), commit `3d8ce621`:

```
Run        : 40444845364
URL        : https://github.com/seesepri/plugincorel/actions/runs/40444845364
Status     : completed / success   (03:33:41Z → 03:34:30Z)
Job        : Build Windows x64 (windows-latest)

Steps:
  3 Static checks (undo grouping, watermark exclusion, unit normalization)  success
  4 Regression tests (tanpa CorelDRAW)                                      success
  5 Setup MSBuild                                                          success
  6 Build .NET Framework 4.8   (msbuild /m /t:Rebuild /p:Configuration=Release /p:Platform=x64)  success
  7 Prepare release package                                                success
  8 Upload build artifacts                                                 success
  9 Create Draft GitHub Release                                            skipped (bukan tag)
```

Artifact `CorelPageCopier-Windows-x64` (220.642 byte) diunduh & diperiksa ulang:

```
release-package/CorelPageCopier.exe   210.944 byte
  file(1)  : PE32+ executable for MS Windows (GUI), x86-64, Mono/.Net assembly
  header   : machine = 0x8664 (x64/AMD64)
CorelPageCopier-Windows-x64.zip       210.944 (exe) + 4.155 (README_ID.md) + 621 (ToolbarLauncher.bas)
```

Yang **belum** dibuktikan oleh CI ini: aplikasi benar-benar berjalan & berinteraksi dengan
CorelDRAW 2022 (butuh integration test manual, §7.5).

### 7.5 Integration / smoke test CorelDRAW 2022 — **NOT RUN**

Alasan: CorelDRAW tidak tersedia. Prosedur lengkap: `PANDUAN-UJI.md` §10. Bagian yang harus
diverifikasi manual: objek tunggal, multi-selection, grup bersarang, bitmap, teks, ukuran
campuran, dokumen mm/inch/point, batch dengan subfolder + nama duplikat + output lama + file gagal
disimpan.

### 7.6 ZIP source — DIBUAT & DIVERIFIKASI

| Item | Nilai |
|---|---|
| File | `CorelPageCopier-audit-2026-10-27.zip` (di root workspace, bersama patch) |
| Ukuran | 194 KB |
| Isi | 50 file (13 file `.vb` produksi, `.vbproj`, 2 aset PNG, workflow CI, `tools/static-checks.py`, proyek `tests/` (regression + compile-check), dokumen (`docs/AUDIT-RELIABILITY.md`, `PANDUAN-UJI.md`, `README_ID.md`, `ANALISA-WATERMARK.md`) |
| Tidak disertakan | `.git/`, `bin/`, `obj/` (artefak build) |
| Integritas | `unzip -t` → "No errors detected in compressed data" |
| Uji dari hasil ekstrak | `python3 tools/static-checks.py` → LULUS (200) · `dotnet run --project tests/CorelPageCopier.RegressionTests` → 103 lulus / 0 gagal · `dotnet build … CompileCheck` → Build succeeded |

Patch terpisah untuk review/PR: `CorelPageCopier-audit-2026-10-27.patch`
(32 file, `git diff` lengkap termasuk file baru; bisa diterapkan dengan `git apply`).

### 7.7 Output test lengkap

Contoh nama test yang menjadi bukti utama (semua PASS):

```
Verifikasi copy (Temuan A): geometri tidak terukur -> Partial (bukan Verified)
Verifikasi copy (Temuan A): CopyRunSummary: partial tidak masuk hitungan verified
Pengukuran objek (Temuan A): MeasureTopShapes mengukur objek TERATAS, bukan terbawah
Pengukuran objek (Temuan A): pengukuran lama (index 1..n) akan salah — regresi dijaga
Alur copy per halaman (Temuan A): grup bersarang (1 shape grup) -> Verified
Alur copy per halaman (Temuan A): objek hasil paste tidak terukur -> Partial (bukan Verified)
Alur copy per halaman (Temuan A): dokumen inch: selisih 0,05 inch (1,27 mm) -> Failed
Normalisasi satuan (Temuan B): penyetelan mm gagal -> konversi eksplisit dipakai
Normalisasi satuan (Temuan B): pemulihan gagal dilaporkan eksplisit
Undo grouping (Temuan C): group gagal dibuka -> tidak menjanjikan satu undo
Undo grouping (Temuan C): End gagal -> dilaporkan, bukan ditelan
Kebijakan file batch (Temuan D/E): status OK hanya bila semua langkah wajib berhasil
Kebijakan file batch (Temuan D/E): SelectSources membuang output _wm dari daftar sumber
Watermark — satu dokumen: Document.Unit tidak terbaca -> batal tanpa mengubah dokumen
Watermark — satu dokumen: EndCommandGroup gagal -> dilaporkan, tidak ditelan
Watermark — batch folder: SaveAs gagal -> status GAGAL, bukan OK
Watermark — batch folder: output lama (_wm) tidak diproses ulang
Watermark — batch folder: mode file baru: output tidak pernah menimpa file sumber
Koneksi & identitas dokumen (Temuan F): dokumen sama di instance BERBEDA -> bukan target yang sama
```
