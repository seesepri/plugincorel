# Laporan Refactor Arsitektur — Corel Page Copier

Tanggal: 2026-10-27
Branch: `refactor/safe-architecture` (dari `2b67ccb`, PR #4)
Target platform **tidak berubah**: VB.NET WinForms, .NET Framework 4.8, x64,
CorelDRAW 2022, **late-bound COM**, tanpa NuGet, tanpa interop baru.

---

## 1. Ringkasan

Refactor ini memisahkan `Program.vb` (dulu 1342 baris: UI + COM + verifikasi +
laporan sekaligus) menjadi UI tipis + tiga service yang bisa diuji tanpa
CorelDRAW. **Tidak ada penulisan ulang**: kode operasi dipindahkan, bukan
ditulis baru; teks dialog dan perilaku tombol dipertahankan.

| Sebelum | Sesudah |
|---|---|
| `CopierForm` = UI + koneksi COM + copy + verifikasi + Auto Select + laporan | `CopierForm` = UI + event handler + port dialog (`ICorelPrompt`, `ICopyPrompt`, `IProgressSink`) |
| 1342 baris di satu file | 983 baris di `Program.vb` + 5 modul baru |
| Logika COM tidak bisa diuji tanpa CorelDRAW | 37 test baru menjalankan seluruh cabang service memakai objek COM palsu |

---

## 2. Daftar perubahan

### 2.1 File baru (produksi)

| File | Isi | Menyentuh COM? |
|---|---|---|
| `CorelPageCopier/ServicePorts.vb` | `IProgressSink` (`Report`, `Pump`), `NullProgressSink`, `RecordingProgressSink` | tidak |
| `CorelPageCopier/CorelComProbe.vb` | `DefaultProgramIds()` (`CorelDRAW.Application.24`, `CorelDRAW.Application`), `GetRunningObject` (reflection atas `Marshal.GetActiveObject`), `CreateInstance` (reflection atas `Type.GetTypeFromProgID` + `Activator.CreateInstance`) | ya (satu-satunya titik) |
| `CorelPageCopier/CorelConnectionService.vb` | `ICorelPrompt`, `ConnectionOutcome`, `Connect`, `Release`, `EnsureDocument`, `RememberLoadedIdentity`, `EnsureSameDocument`, `CurrentIdentity`, `IsAppAlive` | ya |
| `CorelPageCopier/CopyService.vb` | `CopyAbortReason` (10 nilai), `CopyPlan`, `ICopyPrompt`, `CopyServiceResult`, `CopyTargetValidation` (murni), `CopyService.Run` | ya |
| `CorelPageCopier/AutoSelectService.vb` | `AutoSelectScanReport`, `ScanPage`/`ScanAllPages`/`ScanSinglePage`, `SelectPageAtCurrentMatch`, `MoveNext`, `ApplySelection`, `Reset` | ya |

### 2.2 File produksi yang diubah

| File | Perubahan |
|---|---|
| `CorelPageCopier/Program.vb` | −359 baris. Dihapus dari form: `ConnectToCorel` (isi), `ReleaseCorelReferences`, `CurrentIdentity`, `EnsureSameDocument` (isi), isi `CopyClickedCore`, `ScanSinglePage`/`ScanPageForUi`/`SelectPageAtCurrentMatch`/`ApplySelection`/`NextMatchingPageClickedCore` (isi). Ditambah: 3 field service, implementasi `ICorelPrompt` / `ICopyPrompt` / `IProgressSink`, `RenderCopyResult`. Field `corelApp`/`corelDoc`/`loadedIdentity`/`matchNavigator`/`lastScanByPage` dihapus (state dipindah ke service). `Imports System.Runtime.InteropServices` dihapus. |
| `CorelPageCopier/CorelPageCopier.vbproj` | 5 `<Compile Include>` baru (proyek non-SDK, eksplisit). |
| `tools/static-checks.py` | Pemeriksaan yang dulu membaca `Program.vb` sekarang membaca modul barunya (mis. `BeginCommandGroup` di `CopyService.vb`), ditambah blok **batas lapisan** + **baseline UI**. |

### 2.3 Test

| File | Isi |
|---|---|
| `tests/.../Tests/CorelConnectionServiceTests.vb` | 10 test: instance berjalan diutamakan, tolak/setuju instance baru, instance baru tanpa dokumen, dokumen berbeda, `EnsureSameDocument` (file sama + jumlah halaman berubah), `Release` lalu `Connect`, ProgID 2022 |
| `tests/.../Tests/CopyServiceTests.vb` | 15 test: validasi target murni (rentang, halaman sumber, duplikat), tiap `CopyAbortReason`, urutan operasi (konfirmasi sebelum mengubah dokumen), satuan naik-turun, kegagalan pemulihan satuan, objek sumber tidak dihapus, exception per halaman |
| `tests/.../Tests/AutoSelectServiceTests.vb` | 12 test: pemindaian semua halaman + snapshot, progres per halaman, kategori spesifik, tanpa hasil (navigasi tidak dimulai), `ApplySelection` tanpa target (tidak `ClearSelection`), verifikasi jumlah, `AddToSelection` tidak berefek, `ClearSelection` gagal (tanpa fallback `Delete`), navigasi maju, `StaleSnapshot`, `Reset` |
| `tests/.../FakeCorel.vb` | Ditambah: `FakeClipboard` (dipakai bersama), `FakeSelectionRange` (`Count`, `Copy`), `FakeDocument.Selection`/`ClearSelection`/`ActiveSelectionRange`/`CopySelection`, `FakeLayer.Paste`, `FakePage.Shapes`, `FakeShape.Copied`/`Selected`/`AddToSelection`, flag kegagalan terarah (`FailActivate`, `FailAddToSelection`, `FailClearSelection`, `FailUnitRestoreTo`) |
| `tests/.../TestMain.vb` + `.vbproj` | 3 suite baru didaftarkan; 5 sumber produksi dikompilasi ulang (`<Compile ... Link>`) supaya test menguji **sumber yang sama** |

---

## 3. Perbandingan sebelum vs sesudah (per tanggung jawab)

### 3.1 Koneksi CorelDRAW

| Aspek | Sebelum | Sesudah |
|---|---|---|
| Lokasi | `CopierForm.ConnectToCorel` (~110 baris) | `CorelConnectionService.Connect` |
| Pencarian instance | `Marshal.GetActiveObject` langsung di `Program.vb` | `CorelComProbe.GetRunningObject` (reflection, `TargetInvocationException` dibuka) |
| Izin instance baru | `MessageBox.Show` di dalam logika | `ICorelPrompt.ConfirmOpenNewInstance` |
| State | field `corelApp` / `corelDoc` di form | `_app` / `_document` di service, dibaca UI lewat property |
| Uji tanpa CorelDRAW | tidak bisa | 10 test (resolver & factory disuntik) |

### 3.2 Copy

| Aspek | Sebelum | Sesudah |
|---|---|---|
| Lokasi | `CopierForm.CopyClickedCore` (~240 baris) | `CopyService.Run` + `RenderCopyResult` (~40 baris UI) |
| Alasan batal | `Return` tersebar, pesan berbeda-beda di dalam satu method | `CopyAbortReason` (10 nilai) dikembalikan ke UI |
| Validasi target | inline di method | `CopyTargetValidation.Validate` — **murni**, diuji terpisah |
| Dialog | `MessageBox.Show` di tengah pipeline | `ICopyPrompt` (`ConfirmCopy`, `ConfirmContinueWithoutUndoGroup`, `NotifyUnitRestoreFailed`) |
| Progres | `statusLabel.Text` + `Application.DoEvents` | `IProgressSink.Report` / `Pump` |
| Uji tanpa CorelDRAW | tidak bisa | 15 test (sukses, semua jalur batal, pemulihan satuan, per-halaman gagal) |

### 3.3 Auto Select

| Aspek | Sebelum | Sesudah |
|---|---|---|
| Lokasi | 5 method di form (`AutoSelectClickedCore`, `ScanSinglePage`, `ScanPageForUi`, `SelectPageAtCurrentMatch`, `ApplySelection`, `NextMatchingPageClickedCore`) | `AutoSelectService` + laporan di form |
| State pemindaian | `matchNavigator` + `lastScanByPage` di form | `AutoSelectService.Navigator` / `.Scans`, `Reset()` |
| Penerapan seleksi | `ApplySelection` di form | `AutoSelectService.ApplySelection` (aturan sama: tanpa target → seleksi lama utuh; `ClearSelection` gagal → batal; tanpa `Delete`) |
| Uji tanpa CorelDRAW | tidak bisa | 12 test |

### 3.4 Yang **tidak** berubah

* Semua teks tombol, judul panel, urutan combo, status awal, nama handler.
* Teks setiap `MessageBox` (dipatok `tools/static-checks.py`).
* Late-bound COM (tanpa interop baru, tanpa `Microsoft.Office.Interop`).
* `Option Strict Off`, gaya `Object` + `CInt(...)` seperti sebelumnya.
* `PageItem`, busy guard (`OperationGate`), `CopierFormClosing`.
* Tidak ada dependency baru; semua file terdaftar di `.vbproj` non-SDK.

---

## 4. Bukti tidak ada regresi

### 4.1 Otomatis (dijalankan di lingkungan ini, tanpa CorelDRAW)

| Pemeriksaan | Baseline `2b67ccb` | Sesudah refactor | Hasil |
|---|---|---|---|
| Regression test (`dotnet run --project tests/CorelPageCopier.RegressionTests -c Release`) | 216 lulus / 0 gagal | **253 lulus / 0 gagal** | PASS |
| Pemeriksaan statis (`python3 tools/static-checks.py`) | 295 lulus / 0 gagal | **351 lulus / 0 gagal** | PASS |
| Kompilasi net48 semua modul (`/tmp/cc`, SDK-style glob) | 0 warning / 0 error | **0 warning / 0 error** | PASS |
| Build Release x64 Windows (MSBuild non-SDK, CI `windows-latest`) | PASS (run 40447818533) | lihat §4.3 | — |

216 test lama **tetap ada dan tetap lulus**; 37 test baru menambah cakupan
service. 56 pemeriksaan statis baru memaku UI & batas lapisan.

### 4.2 Pemeriksaan statis baru yang mengunci refactor

| Pemeriksaan | Tujuan |
|---|---|
| `Program.vb` tidak memuat `Marshal.GetActiveObject` | pelamaran COM hanya di `CorelComProbe.vb` |
| `Program.vb` tidak memuat `BeginCommandGroup` / `EndCommandGroup` | grup undo hanya di `CopyService.vb` (dan `WatermarkEngine`) |
| `Program.vb` tidak memuat `ActiveSelectionRange` / `AddToSelection` | seleksi hanya di `AutoSelectService.vb` |
| Service tidak `Imports System.Windows.Forms`, tidak memanggil dialog | UI hanya lewat port |
| Port `ICorelPrompt` / `ICopyPrompt` / `IProgressSink` dipakai | interaksi teruji pengganti |
| `CopyTargetValidation` ada & murni | bagian copy yang paling rawan diuji tanpa COM |
| Semua file baru terdaftar di `CorelPageCopier.vbproj` & test `.vbproj` | proyek non-SDK tetap build di Windows |
| Teks tombol/handler/panel/status dipatok | UI tidak berubah tanpa ketahuan |

### 4.3 Build Release x64 Windows

Dijalankan oleh GitHub Actions (`windows-latest`, MSBuild, non-SDK `.vbproj`,
platform x64) pada commit refactor ini. Hasil run dicatat di komentar PR dan
di `docs/ARCHITECTURE.md` §6 — **hanya** diklaim PASS bila log MSBuild
menunjukkan `0 Warning(s) / 0 Error(s)` dan `bin\x64\Release\CorelPageCopier.exe`.

### 4.4 Yang tetap **NOT RUN**

| Uji | Status | Alasan |
|---|---|---|
| Integration test CorelDRAW 2022 (klik nyata di aplikasi) | **NOT RUN** | CorelDRAW tidak tersedia di lingkungan ini |
| Verifikasi visual UI setelah refactor (layout, teks di layar) | **NOT RUN** | idem; yang tersedia hanya pemeriksaan statis atas teks & handler |
| Copy/watermark pada dokumen CDR nyata | **NOT RUN** | idem |
| Kualitas cetak / PDF nyata | **NOT RUN** | tidak diklaim |

Status di atas **tidak boleh** diubah menjadi PASS tanpa menjalankan
`PANDUAN-UJI.md` §10–§13 di mesin ber-CorelDRAW.

---

## 5. Risiko & catatan teknis

1. **Urutan operasi copy dipertahankan persis**: validasi target → pastikan
   dokumen → pastikan seleksi → baca satuan → konfirmasi → normalisasi mm →
   grup undo → copy/paste → verifikasi → `Finally` (tutup group, pulihkan
   satuan, aktifkan halaman sumber). Test mengunci urutan ini.
2. **`CopyService` memanggil `_connection.Connect()`** di awal `Run`,
   meniru `ConnectToCorel()` di awal setiap operasi versi lama. Bila pengguna
   menolak instance baru, `Connect` melempar dan kegagalan muncul sebagai
   "Copy tidak berjalan" + detail — sama seperti sebelumnya.
3. **`EnsureSameDocument` membaca ulang dokumen aktif** lewat `Connect` yang
   dipanggil tiap operasi, sehingga pindah dokumen tetap terdeteksi.
4. **Dialog tetap di UI**: service tidak pernah menampilkan dialog sendiri,
   jadi seluruh cabang bisa diuji, dan teks dialog bisa dipatok statis.
5. **`WatermarkDialog` tidak dipecah** pada refactor ini: ia sudah terpisah dari
   `Program.vb`, dan memecahnya tidak dibutuhkan untuk memutus coupling UI ↔ COM
   pada alur copy/Auto Select. Ini sengaja dicatat sebagai pekerjaan lanjutan
   yang terpisah, bukan diklaim selesai.
