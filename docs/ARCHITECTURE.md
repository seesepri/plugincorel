# Arsitektur — Corel Page Copier

Dokumen ini mencatat pembagian tanggung jawab setelah refactor arsitektur
2026-10-27, beserta **baseline** yang diambil sebelum kode dipindahkan.

Target platform tidak berubah: VB.NET Windows Forms, .NET Framework 4.8, x64,
CorelDRAW 2022, **late-bound COM** (tanpa referensi interop, tanpa NuGet).

---

## 1. Baseline (diambil SEBELUM refactor)

Bukti awal, pada commit `2b67ccb` (PR #4), sebelum satu baris pun dipindahkan:

| Ukuran | Nilai |
|---|---|
| Regression test | **216 lulus, 0 gagal** (`dotnet run --project tests/... -c Release`) |
| Pemeriksaan statis | **295 lulus, 0 gagal** (`python3 tools/static-checks.py`) |
| Kompilasi net48 (lokal) | 0 warning, 0 error |
| Build Release x64 Windows (CI) | PASS — run 40447818533 / 40447928301 |
| `Program.vb` | 1342 baris, satu kelas `CopierForm` + `PageItem` |
| Modul pendukung | 17 file, 9027 baris total |

### 1.1 Inventaris UI yang WAJIB dipertahankan

Teks tombol, urutan panel, dan nama handler di bawah ini dipatok oleh
`tools/static-checks.py` (bagian "Arsitektur & baseline UI") supaya refactor
tidak bisa mengubah UI tanpa ketahuan.

| Panel | Kontrol | Teks / isi | Handler |
|---|---|---|---|
| Header | tombol | `Watermark...` | `OpenWatermarkDialog` |
| 01 Salin ke halaman | judul | `01  Salin ke halaman` | — |
| | tombol | `Muat ulang` | `RefreshClicked` |
| | tombol | `Pilih semua` | lambda → `SetAllChecks(True)` |
| | tombol | `Kosongkan` | lambda → `SetAllChecks(False)` |
| | tombol utama | `Salin ke halaman terpilih` | `CopyClicked` |
| 02 Auto Select objek | judul | `02  Auto Select objek` | — |
| | combo kategori | `Teks (Artistik + Paragraf)`, `Gambar / Bitmap`, `Garis / Kurva terbuka`, `Objek vektor (tanpa teks/bitmap/grup)`, `Semua objek` | — |
| | combo cakupan | `Halaman aktif`, `Pindai semua halaman` | — |
| | tombol utama | `Select otomatis` | `AutoSelectClicked` |
| | tombol | `Pilih hasil berikutnya` | `NextMatchingPageClicked` |
| Status | label | `Status: siap. Buka dokumen di CorelDRAW.` | — |

### 1.2 Perilaku yang dikunci (characterization)

| Perilaku | Dijaga oleh |
|---|---|
| Busy guard: klik ganda diabaikan, UI dipulihkan di `Finally`, jendela tidak bisa ditutup saat sibuk | static checks + `OperationGate` unit test |
| Konfirmasi sebelum mengubah dokumen (copy, watermark) | static checks (teks dialog) |
| Batal ≠ gagal: setiap jalur batal punya alasan eksplisit | test `CopyAbortReason` (baru) |
| Tidak ada `Delete` sebagai fallback | static checks |
| Watermark tidak pernah ikut Auto Select | static checks + test |
| Nomor halaman/urutan cek halaman tidak berubah | static checks (`PageItem` tetap ada) |

---

## 2. Struktur setelah refactor

```
CopierForm (Program.vb)            UI & event handler saja
   ├─ CorelConnectionService       koneksi, identitas instance, validasi dokumen
   ├─ CopyService                  copy + verifikasi + laporan
   ├─ AutoSelectService            pemindaian, navigasi, penerapan seleksi
   └─ WatermarkDialog              UI watermark
        └─ WatermarkEngine         operasi watermark & batch
```

| Kelas | File | Tanggung jawab | Menyentuh COM? | Diuji tanpa CorelDRAW? |
|---|---|---|---|---|
| `CopierForm` | `Program.vb` | UI, event handler, status bar, busy guard, prompt WinForms | tidak langsung (lewat service) | tidak (WinForms) |
| `CorelConnectionService` | `CorelConnectionService.vb` | ambil instance (ROT), izin buka instance baru, identitas dokumen, `EnsureSameDocument`, release COM | ya (late bound) | **ya** (resolver/factory/prompt disuntik) |
| `CopyService` | `CopyService.vb` | validasi target, normalisasi satuan, undo group, paste per halaman, verifikasi, laporan | ya | **ya** (`FakeCorel`) |
| `CopyTargetValidation` | `CopyService.vb` | validasi daftar halaman tujuan (murni) | tidak | **ya** |
| `AutoSelectService` | `AutoSelectService.vb` | pindai halaman, snapshot navigasi, terapkan seleksi + verifikasi jumlah | ya | **ya** (`FakeCorel`) |
| `ShapeClassifier` / `ShapeScanner` / `ShapeNode` / `SelectionNavigation` | file masing-masing | klasifikasi & traversal murni | hanya `ShapeNode` (adapter) | **ya** |
| `WatermarkEngine` | `WatermarkEngine.vb` | watermark per dokumen & batch | ya | **ya** (`FakeCorel`) |
| `WatermarkLayout` / `WatermarkPlacementCheck` | file masing-masing | matematika posisi & geometri | tidak | **ya** |
| `BatchLog` / `BatchFilePolicy` / `BatchRunTracker` / `PdfProofValidator` | file masing-masing | log CSV, kebijakan file, pembukuan, validasi PDF | tidak | **ya** |
| `PresetStore` | `PresetStore.vb` | baca/tulis preset `%APPDATA%` | tidak | **ya** |

Aturan batas yang dijaga static checks:

1. `Program.vb` **tidak boleh** memuat lagi orkestrasi COM tingkat lanjut
   (`Marshal.GetActiveObject`, `BeginCommandGroup`, `Paste()`, `AddToSelection`,
   `PublishToPDF`) — semuanya harus lewat service.
2. Service **tidak boleh** memakai `System.Windows.Forms` (kecuali
   `WatermarkDialog`, yang memang UI). Interaksi pengguna lewat interface
   `ICorelPrompt` / `ICopyPrompt` / `IProgressSink`.
3. Semua file baru terdaftar di `CorelPageCopier.vbproj` (non-SDK, `<Compile Include>`
   eksplisit) dan di `.vbproj` test sebagai `<Compile ... Link>`.

---

## 3. Alur

### 3.1 Koneksi (`CorelConnectionService.Connect`)

```
1. bila koneksi lama masih hidup & punya ActiveDocument  -> pakai
2. ROT: Marshal.GetActiveObject("CorelDRAW.Application.24", "CorelDRAW.Application")
      ada dokumen  -> pakai
      tanpa dokumen -> catat alasan, lepas objek, lanjut
3. tidak ada  -> TANYA pengguna (ICorelPrompt.ConfirmOpenNewInstance)
      tidak setuju -> throw (tidak ada instance baru, tidak ada dokumen berubah)
4. buat instance -> Visible = True
      ada ActiveDocument -> pakai + beri tahu pengguna
      tidak ada          -> beri tahu + lepas objek + throw
```

Semua pesan kegagalan dipertahankan apa adanya (dipatok static checks).

### 3.2 Copy (`CopyService.Run`)

```
validasi target (murni)         -> NoTargetPages / NoValidTargets
pastikan dokumen masih sama     -> SourcePageChanged / DocumentMismatch
pastikan ada seleksi            -> NoSelection
baca satuan asli                -> UnitUnreadable
tanya konfirmasi (ICopyPrompt)  -> UserCancelled
normalisasi satuan ke mm        -> UnitNotNormalizable
buka grup undo (boleh gagal)    -> UndoGroupDeclined (tanya pengguna)
Copy() lalu Paste() per halaman -> verifikasi per halaman (Verified/Partial/Failed)
Finally: EndCommandGroup, pulihkan satuan, aktifkan halaman sumber
laporan: CopyRunSummary.BuildReport(...)
```

Semua jalur batal mengembalikan `CopyAbortReason` + pesan; UI menampilkan
dialog yang sama seperti sebelumnya.

### 3.3 Auto Select (`AutoSelectService`)

```
ScanAllPages(category)   -> ShapeScanResult per halaman + total + halaman berisi hasil
                          (progress lewat IProgressSink, DoEvents di UI)
Navigator.Start(...)     -> snapshot kategori + cakupan
SelectPageAtCurrentMatch -> aktivasi halaman, ApplySelection (ClearSelection hanya
                            bila ada target), verifikasi lewat ActiveSelectionRange
MoveNext(categoryNow)    -> StaleSnapshot / PageMissing / AtEnd / NotStarted
```

---

## 4. Yang TIDAK berubah

* Reaksi tombol, teks tombol, urutan panel, dan nama handler.
* Late-bound COM: `Marshal.GetActiveObject` + `Type.GetTypeFromProgID` +
  `Activator.CreateInstance`, tanpa interop baru.
* `Option Strict Off` dan gaya late binding yang sudah ada.
* Tidak ada dependency eksternal baru (hanya BCL .NET Framework 4.8).

## 5. Batasan & hal yang belum diverifikasi

* Integration test CorelDRAW 2022 tetap **NOT RUN** (tidak ada CorelDRAW di
  lingkungan ini) — lihat `PANDUAN-UJI.md` §10–§12.
* `CopierForm` sendiri tidak diuji unit (butuh WinForms); yang diuji adalah
  service di belakangnya + pemeriksaan statis yang mematok permukaan UI.
* Service COM diuji dengan objek palsu (`FakeCorel.vb`), bukan COM nyata.

---

## 6. Verifikasi SETELAH refactor (bukti regresi)

Diukur pada branch `refactor/safe-architecture` setelah semua pemindahan selesai.

| Ukuran | Sebelum (`2b67ccb`) | Sesudah | Selisih |
|---|---|---|---|
| Regression test | 216 lulus / 0 gagal | **253 lulus / 0 gagal** | +37 test baru untuk service |
| Pemeriksaan statis | 295 lulus / 0 gagal | **351 lulus / 0 gagal** | +56 pemeriksaan (batas lapisan & UI) |
| Kompilasi net48 (lokal) | 0 warning / 0 error | **0 warning / 0 error** | tetap |
| `Program.vb` | 1342 baris | **983 baris** | −359 (−27 %) |
| Modul pendukung | 17 file / 9027 baris | 25 file / 9826 baris | +8 file (5 service/port + tests) |
| `Marshal.GetActiveObject` di `Program.vb` | ada | **tidak ada** (pindah ke `CorelComProbe.vb`) | 0 |
| `BeginCommandGroup` / `ActiveSelectionRange` / `AddToSelection` di `Program.vb` | ada | **tidak ada** | 0 |
| `MessageBox` di dalam logika operasi | ada | **tidak ada** (hanya di port UI) | 0 |

Arti angka di atas:

* Tidak ada satu pun test lama yang dihapus atau dilonggarkan — 216 test lama
  tetap lulus, 37 test baru menambah cakupan untuk service hasil pemindahan.
* 56 pemeriksaan statis tambahan memaku **baseline UI** (teks tombol, handler,
  urutan panel, status awal) dan **batas lapisan** (service tanpa WinForms,
  `Program.vb` tanpa orkestrasi COM).

Rincian lengkap per perubahan: lihat `docs/REFACTOR-REPORT.md`.
