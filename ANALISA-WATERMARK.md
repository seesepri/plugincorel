# Analisa Repo `seesepri/plugincorel` + Rencana Fitur Watermark

Dibuat: 26 Oktober 2026 · Basis: commit `968b08a` ("Improve application UI layout")
Status: repo awalnya 404 (private), analisa dilakukan setelah repo bisa di-clone.

---

## 1. Ringkasan repo

| Item | Nilai |
|---|---|
| Nama produk | **Corel Page Copier** |
| Bahasa | VB.NET (Windows Forms) |
| Target | .NET Framework **4.8**, **x64**, WinExe |
| Ketergantungan | Hanya BCL (`System`, `System.Drawing`, `System.Windows.Forms`), **tanpa** interop CorelDRAW |
| Interop | COM **late binding** via `Marshal.GetActiveObject` / `Activator.CreateInstance` → tipe `Object` |
| Jumlah file | 7 (6 source + config) |
| Ukuran kode | `Program.vb` = **691 baris** (seluruh UI + logika dalam 1 file) |
| CI/CD | GitHub Actions `windows-latest` → MSBuild → ZIP artifact; draft Release saat push tag `v*` |
| Dokumentasi | `GITHUB_SETUP_ID.md`, `CorelPageCopier/README_ID.md` |
| Lisensi | **belum ada file LICENSE** |

Fitur yang sudah ada:

1. **Salin objek** dari halaman aktif ke halaman-halaman yang dicentang (`CopyClicked`).
2. **Auto Select** objek berdasarkan kategori (teks / bitmap / kurva terbuka / vektor / semua) di halaman aktif atau seluruh halaman (memindai + navigasi hasil).
3. Launcher makro VBA (`ToolbarLauncher.bas`) yang memanggil EXE.

### Arsitektur sekarang (satu file)

```
Program.vb
└── CopierForm (Form)
    ├── BuildUi()                     → susun kontrol manual (TableLayoutPanel + kartu)
    ├── ConnectToCorel()              → ROT lookup "CorelDRAW.Application.24" → fallback CreateInstance
    ├── LoadPages()                   → isi CheckedListBox halaman
    ├── CopyClicked()                 → copy → paste per halaman + laporan
    ├── AutoSelectClicked()           → scan + seleksi objek
    ├── GetMatchingShapes/CollectMatchingShape/MatchesCategory
    ├── ApplySelection()
    └── PageItem (inner class)
```

Kalau watermark ditambahkan langsung ke `CopierForm`, file itu akan melewati ±1.200 baris dengan 3 tanggung jawab sekaligus (copy, select, watermark). Karena itu kode watermark saya pisahkan ke 3 file baru (lihat §4).

---

## 2. Temuan teknis (hasil review)

Prioritas: **P1** = berpotensi merusak data / salah dokumen, **P2** = bug / risiko perilaku, **P3** = kebersihan kode.

| # | Temuan | Dampak | Prioritas | Status |
|---|---|---|---|---|
| 1 | **Tidak ada undo grouping.** Copy & watermark tidak dibungkus `doc.BeginCommandGroup`/`EndCommandGroup`, jadi user harus Ctrl+Z berkali-kali untuk membatalkan | Salah klik = repot, terutama kalau nanti batch 50 file | **P1** | ✅ diperbaiki (§4.4) |
| 2 | **`Auto Select` "Semua objek" akan ikut memilih watermark.** `CollectMatchingShape` untuk kategori 4 menambahkan semua shape. Setelah watermark ada, "Pilih semua + Salin" akan menyalin watermark juga | Watermark dobel/berantakan di halaman tujuan | **P1** | ✅ diperbaiki (§4.4) |
| 3 | **Watermark harus punya penanda.** Tanpa nama shape/layer khusus, watermark tidak bisa diganti/dihapus terprogram dan tidak bisa dikecualikan dari Auto Select | Watermark jadi sampah permanen di dokumen | **P1** | ✅ diperbaiki (§4.4) |
| 4 | **Satuan dokumen tidak pernah dinormalisasi.** Semua posisi objek CorelDRAW dibaca/ditulis dalam `Document.Unit` (bisa inch/mm/pt tergantung dokumen). Fitur watermark butuh mm absolut | Watermark bisa salah skala per dokumen | **P1** | ✅ diperbaiki (§4.4) |
| 5 | `ConnectToCorel` bisa membuka **instance CorelDRAW kedua** (fallback `Activator.CreateInstance`), lalu user mengedit dokumen di instance lain → aplikasi menulis ke dokumen yang salah | Perubahan hilang / user bingung | **P2** |
| 6 | `Application.DoEvents()` di dalam loop paste/scan → re-entrancy (user bisa klik tombol lain saat proses berjalan); tidak ada flag `busy` | State kacau saat proses panjang | **P2** |
| 7 | Verifikasi hasil copy hanya berbasis **jumlah shape tingkat atas** (`afterCount > beforeCount`) — grup/layer khusus bisa lolos tanpa perubahan jumlah shape | False positive "berhasil" | **P2** | ✅ diperbaiki (§4.4) |
| 8 | `Option Strict Off` + late binding: typo nama properti COM baru ketahuan saat runtime. Wajar untuk COM tanpa interop, tapi butuh pola `Try/Catch` per-API (sudah saya pakai di kode watermark) | Crash runtime | **P2** |
| 9 | Tidak ada penyimpanan preset/setting (ukuran watermark, logo brand, dsb.) → user mengisi ulang tiap kali | UX | **P3** |
| 10 | `ToolbarLauncher.bas` hardcode `C:\Program Files\Corel Page Copier\CorelPageCopier.exe` | Gagal kalau dipasang di path lain | **P3** |
| 11 | Belum ada LICENSE (repo publik) & belum ada versi/tag | Legal & distribusi | **P3** |
| 12 | Proyek **non-SDK style**: daftar `<Compile Include>` eksplisit. File baru yang tidak didaftarkan **tidak ikut ter-build**, dan CI tidak akan error — hanya fitur hilang | Jebakan saat menambah fitur | **P3** (sudah saya tangani) |

Catatan yang **sudah benar** dan patut dipertahankan: `Pages` 1-based, tidak memakai `Pages(0)` sebagai halaman; validasi indeks target; `ApplySelection` tidak memakai fallback hapus (aman); COM exception per halaman tidak menggagalkan seluruh proses; workflow CI membuat ZIP + draft release hanya saat tag.

---

## 3. Desain fitur watermark

### 3.1 Kebutuhan (sesuai permintaan: teks + logo, sekaligus bisa batch)

| Kebutuhan | Keputusan desain |
|---|---|
| Watermark teks (brand/nama) | `Layer.CreateArtisticText` + font/ukuran/warna/opasitas |
| Watermark logo (PNG transparan) | `Layer.Import` lalu skala berdasarkan **lebar mm** (rasio dijaga) |
| Teks + logo sekaligus | Dua objek terpisah di layer yang sama, posisi sama (logo di atas/bawah teks) |
| Batch banyak file | `ApplyToFolder` → buka tiap `.cdr`, terapkan, simpan |
| Aman untuk dokumen klien | Default simpan ke `nama_wm.cdr`, **bukan** menimpa file asli |
| Bisa diulang / dibatalkan | Shape dinamai `WM_*` di layer `WATERMARK`; `ReplaceExisting` menghapus versi lama dulu |

### 3.2 Model pengaturan (`WatermarkSettings`)

Jenis (`Teks / Logo / Teks+Logo`), teks, font + ukuran, warna, path logo + lebar mm, opasitas %, rotasi °, jangkar 3×3 + margin, mode tile + jarak X/Y, cakupan halaman (semua / halaman tertentu), nama layer, opsi idempoten (`ReplaceExisting`), nomor halaman otomatis, dan kunci layer. Satu objek = satu preset → mudah diserialisasi (JSON) di iterasi berikutnya.

### 3.3 API CorelDRAW yang dipakai (semua late binding, tanpa referensi interop)

| Kebutuhan | API | Catatan |
|---|---|---|
| Satuan mm | `Document.Unit = cdrMillimeter` (**= 3**) | disimpan & dipulihkan setelah proses |
| Undo 1 langkah | `Document.BeginCommandGroup` / `EndCommandGroup` | selalu dipasangkan di `Finally` |
| Layer khusus | `Page.CreateLayer(nama)`, `Layer.Editable`, `Layer.Visible` | layer watermark di atas objek desain |
| Teks | `Layer.CreateArtisticText(kiri, bawah, teks, languageID, charSet, font, size)` | posisi dalam satuan dokumen |
| Logo | `Layer.Import(file)` → fallback `ImportEx(file, 0, Nothing)` | PNG dengan alpha |
| Skala | `Shape.SizeWidth` / `SizeHeight` (setter) | tinggi dihitung otomatis dari lebar → rasio aman |
| Posisi | `Shape.CenterX` / `CenterY` | origin CorelDRAW di **kiri-bawah** |
| Rotasi | `Shape.Rotate(derajat)` | |
| Transparansi | `Shape.Transparency.ApplyUniformTransparency(nilai)` + set `.Uniform` (workaround resmi Corel), `.AppliedTo = cdrApplyToFillAndOutline` (**= 2**), `.MergeMode = cdrMergeNormal` (**= 0**) | opasitas 25% → transparency 75% |
| Cari/hapus | `Shape.Name = "WM_*"` lalu iterasi `Layer.Shapes` | idempoten |
| Batch | `Application.OpenDocument`, `Document.SaveAs`, `Document.Close` | simpan non-destruktif default |

Semua panggilan dibungkus `Try/Catch` per langkah supaya API yang tidak tersedia di versi CorelDRAW tertentu tidak menggagalkan seluruh proses (mis. `MoveToTop`, `Transparency.Uniform`).

### 3.4 Algoritma penempatan (bagian yang bisa diuji tanpa CorelDRAW)

- **Jangkar tunggal:** 9 posisi (3×3) + margin, dihitung dalam koordinat kiri-bawah.
- **Tile (pola berulang):** jumlah kolom/baris = `ceil(halaman / jarak)`, distribusi **simetris** (margin kiri = kanan, atas = bawah), objek yang sebagian keluar halaman tetap ditaruh supaya tepi halaman tetap ter-watermark. Jarak minimum dipaksa 110% lebar / 140% tinggi objek agar tidak menumpuk.
- **Fallback:** objek lebih besar dari halaman → 1 watermark di tengah.
- **Konversi opasitas:** `transparency = 100 − opasitas` (di-clamp 0..100).

---

## 4. Kode yang sudah saya siapkan

Sudah ditambahkan ke repo (dan `CorelPageCopier.vbproj` sudah diperbarui):

| File | Isi | Baris |
|---|---|---|
| `CorelPageCopier/WatermarkLayout.vb` | `WatermarkKind`, `WatermarkAnchor`, `WatermarkSettings` (+ validasi & deskripsi), `WatermarkLayout` (matematika jangkar/tile/opasitas) — **tanpa** COM & WinForms | ±300 |
| `CorelPageCopier/WatermarkEngine.vb` | `WatermarkEngine`: `ApplyToDocument`, `RemoveFromDocument`, `ApplyToFolder` (batch), helper layer/shape/gaya, `WatermarkResult` + laporan | ±660 |
| `CorelPageCopier/WatermarkDialog.vb` | Dialog UI lengkap (jenis, teks, font, warna, logo, opasitas, rotasi, jangkar, margin, tile, cakupan halaman, layer, batch folder, area laporan) | ±614 |
| `CorelPageCopier/Program.vb` (dimodifikasi) | undo grouping + normalisasi satuan + verifikasi geometri (§4.4), flag `busy` global (§4.5), Auto Select mengecualikan watermark | 691 → 997 |
| `tools/static-checks.py` | 41 pemeriksaan regresi untuk #1/#2/#3 + busy guard (dipanggil CI sebelum build) | 190 |

### 4.1 Verifikasi yang sudah dilakukan

| Uji | Hasil |
|---|---|
| Kompilasi seluruh proyek (Program.vb + 3 file baru) sebagai `.NET Framework 4.8` WinExe x64 | **Build succeeded, 0 error** |
| Kompilasi setelah sisipan tombol watermark di `CopierForm` (kode integrasi di §4.2) | **Build succeeded, 0 error** |
| Uji unit matematika (jangkar, tile, simetri, fallback, mapping opasitas, validasi setting) | **52 assertion, semua lulus** (31 dasar + 21 untuk perbaikan §4.4) |
| Pemeriksaan statis `tools/static-checks.py` (menjaga #1/#2/#3 saat refactor) | **22 pemeriksaan, semua lulus** |
| Uji integrasi nyata ke CorelDRAW 2022 | **belum bisa** — tidak ada Windows/CorelDRAW di lingkungan ini (sama seperti catatan di `README_ID.md` repo) |

### 4.2 Cara memasang (2 langkah, ±2 menit)

**Langkah 1 — `CorelPageCopier.vbproj`** (sudah saya lakukan; kalau mau ulang manual):

```xml
<Compile Include="Program.vb">
  <SubType>Form</SubType>
</Compile>
<Compile Include="WatermarkLayout.vb" />
<Compile Include="WatermarkEngine.vb" />
<Compile Include="WatermarkDialog.vb" />
```

**Langkah 2 — hook tombol di `CopierForm`** (di `Program.vb`, dua sisipan kecil):

Di dalam `BuildUi()`, tepat sebelum blok `header.Controls.AddRange(...)`, tambahkan:

```vb
' --- TOMBOL WATERMARK (tambahan) ---
Dim wmButton As Button = CreateButton("Watermark...", False, 120)
wmButton.Location = New Point(430, 24)
wmButton.Anchor = AnchorStyles.Top Or AnchorStyles.Right
AddHandler wmButton.Click, AddressOf OpenWatermarkDialog
```

lalu ubah baris kontrol header menjadi:

```vb
header.Controls.AddRange(New Control() {
    title, subtitle, sourceLabel, wmButton
})
```

dan tambahkan method berikut (mis. tepat sebelum `ShowError`):

```vb
Private Sub OpenWatermarkDialog(sender As Object, e As EventArgs)
    Try
        ConnectToCorel()
        Using dlg As New WatermarkDialog(corelApp)
            dlg.ShowDialog(Me)
        End Using
        Try
            LoadPages()
        Catch
        End Try
    Catch ex As Exception
        ShowError("Fitur watermark tidak dapat dibuka: " & ex.Message)
    End Try
End Sub
```

Tidak ada perubahan lain yang diperlukan. Workflow CI tidak perlu diubah (file baru otomatis ikut karena terdaftar di `.vbproj`).

### 4.3 Perilaku fitur

- **Terapkan**: watermark ke semua halaman / halaman aktif saja, satu kali Ctrl+Z untuk membatalkan.
- **Hapus watermark**: sapu semua halaman, hapus shape bernama `WM_*` (aman kalau tidak ada).
- **Jalankan batch**: rekap semua `.cdr` dalam folder (+subfolder), simpan sebagai `nama_wm.cdr` (default) atau timpa file asli (opt-in dengan konfirmasi).
- **Idempoten**: tombol Terapkan boleh ditekan berulang tanpa menumpuk watermark.
- **Laporan**: jumlah halaman diproses, shape dibuat, daftar kegagalan per halaman — sama pola-nya dengan laporan `CopyClicked` yang sudah ada.

### 4.4 Perbaikan Temuan #1–#3 (selesai)

Ketiganya dikerjakan di `Program.vb` + helper bersama di `WatermarkLayout.vb`. Tidak ada perubahan pada `.vbproj`, workflow CI, atau UI (kecuali tombol Watermark yang opsional di §4.2).

#### #1 — Undo grouping

```vb
Try
    corelDoc.BeginCommandGroup("Salin objek ke halaman")
    inCommandGroup = True
Catch
End Try
' ...seluruh loop paste ke semua halaman tujuan...
Finally
    If inCommandGroup Then Try : corelDoc.EndCommandGroup() : Catch : End Try
End Try
```

- Karena `EndCommandGroup` ada di blok `Finally`, grup undo **selalu** ditutup walau ada exception/kegagalan per halaman — tidak mungkin "menggantung".
- Efek: **satu Ctrl+Z membatalkan seluruh paste** ke semua halaman tujuan, bukan satu per satu.
- Watermark (`WatermarkEngine`) sudah memakai pola yang sama sejak awal.
- Posisi grup dimulai **setelah** `selection.Copy()`, karena operasi clipboard tidak perlu masuk undo stack.

#### #2 — Auto Select mengecualikan watermark

Di awal `CollectMatchingShape` (berlaku untuk semua kategori, bukan hanya "Semua objek"):

```vb
If IsWatermarkShape(shape, depth) Then Return
```

`IsWatermarkShape` menganggap objek sebagai watermark jika:

1. `Shape.Name` berawalan `WM_`, **atau**
2. shape berada di layer bernama `WATERMARK` (case-insensitive, trim), **atau**
3. shape adalah grup yang salah satu anggotanya memenuhi (1) atau (2) — rekursif, kedalaman maksimum 32 (sama seperti `CollectMatchingShape`).

Logika nama/layer dipusatkan di `WatermarkLayout.IsWatermarkIdentity(shapeName, layerName)` supaya **satu sumber kebenaran** dipakai bersama oleh Auto Select, `WatermarkEngine.IsWatermarkShape` (hapus/ganti watermark), dan pembuatan watermark.

Catatan perilaku: grup yang **campuran** (berisi objek desain + watermark) akan dilewati seluruhnya saat Auto Select "Semua objek" — konservatif dengan sengaja, karena menyalin grup campuran berarti menyalin watermark juga.

#### #3 — Normalisasi satuan dokumen

```vb
previousUnit = CInt(corelDoc.Unit)   ' disimpan
corelDoc.Unit = CDR_UNIT_MM          ' = 3, selama operasi
' ...ukur sumber, paste, ukur hasil, bandingkan dalam mm...
Finally
    corelDoc.Unit = previousUnit     ' selalu dipulihkan
```

- `CenterX`/`CenterY`/`SizeWidth`/`SizeHeight` dibaca & ditulis dalam `Document.Unit`, jadi tanpa normalisasi perbandingan bisa meleset 25.4× antar dokumen (inch vs mm).
- Verifikasi hasil copy kini **membandingkan geometri**, bukan hanya jumlah shape:
  - ukur objek sumber (gabungan `ShapeRange`) → center + ukuran dalam mm,
  - setelah `Paste`, ukur `N` shape teratas pada layer aktif (`N` = jumlah shape yang bertambah) lalu hitung kotak pembatas gabungannya,
  - cocok bila selisih x, y, lebar, dan tinggi semuanya ≤ **0.5 mm** (`GEOMETRY_TOLERANCE_MM`).
- Kasus tepi yang ditangani eksplisit: objek terlalu rumit/tidak terukur → dihitung "terverifikasi lewat jumlah shape saja" dan **dipisahkan di laporan**, jadi tidak lagi tercampur dengan hasil yang benar-benar terverifikasi.
- Laporan sekarang memisahkan: *geometri cocok* / *hanya jumlah shape* / *gagal*, dan menyebutkan toleransi + fakta bahwa satuan dipulihkan otomatis.

#### Verifikasi perbaikan

| Uji | Hasil |
|---|---|
| Kompilasi proyek (net48, WinExe x64) | **Build succeeded, 0 error, 0 warning** |
| Uji unit `WatermarkLayout` (jangkar, tile, opasitas, validasi, **deteksi watermark**, **geometri/Δ**, simulasi filter Auto Select) | **52 assertion, semua lulus** |
| `python3 tools/static-checks.py` — menjaga #1/#2/#3 tidak hilang saat refactor | **22 pemeriksaan, semua lulus** |
| Uji integrasi CorelDRAW (Ctrl+Z tunggal, exclude watermark, dokumen bersatuan inch) | belum bisa di sini — jalankan checklist §7 no. 3, 5, 10 di Windows |

`tools/static-checks.py` sengaja dibuat tanpa dependensi (Python 3 stdlib) supaya bisa dipanggil manual atau ditambahkan sebagai step CI (`runs-on: ubuntu-latest`) tanpa mengubah job build Windows.


### 4.5 Perbaikan Fase 1.5 #3 — flag `busy` global (selesai)

**Masalahnya:** `Application.DoEvents()` dipakai di dalam loop panjang (paste ke banyak halaman, pindai semua halaman) supaya status progres ter-update. Tapi `DoEvents` juga **memproses klik berikutnya**, jadi user bisa menekan tombol lain (atau tombol yang sama) saat operasi pertama masih berjalan → dua operasi COM berjalan bersamaan pada dokumen yang sama.

**Solusinya** — tiga lapis:

1. **`OperationGate`** (kelas murni di `WatermarkLayout.vb`, tanpa COM/WinForms):
   `TryBegin(nama)` mengembalikan `False` bila sudah sibuk, `Release()` melepas kunci, plus `IsBusy` / `CurrentOperation` / `BusyMessage`. Karena murni, seluruh perilakunya bisa diuji unit (15 assertion) tanpa Windows.
2. **Guard di setiap handler panjang** — pola seragam, badan lama dipindah ke `<Nama>Core()`:

   ```vb
   Private Sub CopyClicked(sender As Object, e As EventArgs)
       If Not TryBeginOperation("Salin objek") Then Return
       Try
           CopyClickedCore()
       Finally
           EndOperation()
       End Try
   End Sub
   ```

   Diterapkan ke `CopyClicked`, `AutoSelectClicked`, `NextMatchingPageClicked`, `RefreshClicked` (paling lama: pindai semua halaman). `EndOperation()` di `Finally` = kunci **tidak mungkin** tertinggal walau ada exception, sehingga UI tidak bisa "terkunci permanen".
3. **Penonaktifan kontrol + proteksi tutup jendela:**
   - `ApplyBusyState(False)` menonaktifkan semua `Button`, `ComboBox`, `CheckedListBox` secara rekursif; status awal tiap kontrol diingat di `busyControlState` sehingga kontrol yang memang sudah nonaktif tetap nonaktif setelah operasi.
   - Kursor menunggu (`UseWaitCursor` + `Cursor`), dan status bar menampilkan nama operasi + "berjalan…".
   - `CopierFormClosing` membatalkan penutupan jendela (`e.Cancel = True`) bila ada operasi berjalan, dengan pesan jelas — mencegah aplikasi mati di tengah loop COM.
   - Klik yang tetap lolos (mis. lewat keyboard sebelum kontrol sempat disable) akan ditolak gate dan status berubah jadi *"Masih memproses: Salin objek. Tunggu hingga selesai."*

**Bonus kecil:** durasi operasi (`Stopwatch`) kini ikut di laporan copy & pemindaian ("Durasi: 3.4 detik"), jadi user tahu apakah operasi memang lambat atau macet.

#### Verifikasi

| Uji | Hasil |
|---|---|
| Build net48 WinExe x64 | **succeeded, 0 error, 0 warning** |
| Uji unit `OperationGate` (klik kedua ditolak, release idempoten, kunci lepas saat exception, pesan status) | **15 assertion lulus** |
| Total uji unit `WatermarkLayout` + `OperationGate` | **67 assertion lulus** |
| `tools/static-checks.py` — kini juga menjaga guard busy (wrapper + `Finally` + `FormClosing`) | **41 pemeriksaan lulus** |
| Verifikasi nyata: klik dobel saat proses berjalan | belum bisa di sini — jalankan checklist §7 no. 16–18 |


---

## 5. Yang perlu kamu putuskan (open questions)

1. **Watermark per halaman vs master page.** Versi ini membuat watermark di **setiap halaman** (layer `WATERMARK` per halaman) → konsisten walau tiap halaman beda ukuran. Alternatif "master layer" lebih hemat objek tapi berisiko pada dokumen dengan ukuran halaman campur. Saran: tetap per halaman (sudah diimplementasikan).
2. **Ekspor PDF ber-watermark** (mis. proof untuk klien): perlu fase 2 (`Document.PublishToPDF` / `ExportEx` per halaman). Belum termasuk.
3. **Logo ber-DPI**: `Layer.Import` menaruh gambar sesuai ukuran aslinya; lebar mm dikontrol, tapi resolusi hasil cetak bergantung file sumber. Minimal 300 DPI pada lebar target.
4. **Lock layer** default `Off` supaya user tetap bisa menggeser manual; bisa dinyalakan.
5. **Penamaan watermark** (`WM_`) sebaiknya juga dipakai untuk mengecualikan watermark dari `Auto Select` (lihat §6).

---

## 6. Saran lanjutan (roadmap)

**Fase 1 — sekarang (sudah ada di repo)**
- Watermark teks + logo, jangkar/tile, opasitas, rotasi, layer khusus, hapus watermark, batch folder, undo grouping, laporan.

**Fase 1.5 — perbaikan kecil**
1. ✅ `CollectMatchingShape` melewati shape watermark (`WM_*` / layer `WATERMARK`), termasuk yang tersembunyi di dalam grup — *selesai (§4.4)*.
2. ✅ `CopyClicked` dibungkus `BeginCommandGroup` / `EndCommandGroup` di `Finally` — *selesai (§4.4)*.
3. ✅ Flag `busy` global di `CopierForm` (`OperationGate`) untuk mencegah klik ganda saat proses berjalan — *selesai (§4.5)*.
4. ✅ `corelDoc.Unit` dinormalisasi ke mm di `CopyClicked` + verifikasi geometri menggantikan verifikasi jumlah shape — *selesai (§4.4)*.

**Fase 2**
- Preset watermark tersimpan di `%APPDATA%\CorelPageCopier\watermark.json` (brand klien A, B, …).
- Log CSV batch (file, halaman, jumlah watermark, status, durasi) + tombol "buka folder hasil".
- Ekspor PDF proof ber-watermark.
- Watermark per-halaman berbeda teks (mis. nomor order dari nama file / `Document.FileName`).

**Fase 3**
- Pindah ke **SDK-style project** (`net48` + `UseWindowsForms`) supaya tidak perlu daftar `<Compile>` manual lagi.
- Pecah `Program.vb` menjadi `CopierForm.vb` + `CorelConnection.vb` + `ShapeScanner.vb` (logika COM yang sama, tanpa UI).
- Opsional: pakai interop resmi `Corel.Interop.VGCore` (compile-time checking, autocomplete) — konsekuensinya butuh referensi DLL Corel di mesin build; late binding sekarang lebih portabel.

---

## 7. Checklist QA manual (sebelum rilis)

| # | Skenario | Ekspektasi |
|---|---|---|
| 1 | Dokumen 1 halaman, watermark teks tengah, opasitas 25% | 1 shape `WM_TEXT`, layer `WATERMARK`, tampak transparan, tidak ada outline |
| 2 | Tekan **Terapkan** 3× | tetap 1 watermark (idempoten) |
| 3 | Ctrl+Z sekali | seluruh watermark hilang |
| 4 | Dokumen 12 halaman ukuran campur (A4 + A5 + custom) | watermark pas di tiap halaman, tidak meleset skalanya |
| 5 | Dokumen bersatuan inch | hasil sama dengan dokumen mm |
| 6 | Logo PNG transparan + teks | logo proporsional, latar tetap transparan |
| 7 | Mode tile, jarak 60×45 mm di A4 | pola rapi, tepi halaman tetap ter-cover, tidak menumpuk |
| 8 | Rotasi 30°, jangkar kanan bawah, margin 8 mm | posisi & sudut benar |
| 9 | **Hapus watermark** | semua `WM_*` hilang, objek desain utuh |
| 10 | Auto Select "Semua objek" setelah watermark ada | hanya objek desain terpilih; `WM_*` & layer `WATERMARK` **tidak** terpilih |
| 10b | Ctrl+Z setelah copy ke 5 halaman | kelima paste hilang sekaligus dalam satu langkah undo |
| 10c | Copy di dokumen bersatuan inch/pixel | laporan "geometri cocok" muncul, satuan dokumen kembali seperti semula setelah operasi |
| 11 | Batch 5 file, mode `_wm.cdr` | 5 file baru, file asli tidak berubah (cek timestamp/ukuran) |
| 12 | Batch dengan 1 file sedang dibuka di CorelDRAW | file itu gagal dengan pesan jelas, file lain tetap diproses |
| 13 | Batch, mode timpa file asli | konfirmasi muncul, isi file ter-update |
| 14 | CorelDRAW ditutup saat dialog terbuka | pesan "Tidak ada dokumen aktif", aplikasi tidak crash |
| 15 | Build CI: push ke `main` | job *Static checks* lulus lalu artifact ZIP berisi EXE versi baru |
| 16 | Saat copy 20 halaman berjalan, klik tombol lain / klik dua kali | klik ditolak, tombol nonaktif, status: "Masih memproses: Salin objek…" |
| 17 | Saat proses berjalan, tekan X (tutup jendela) | jendela tidak tertutup, muncul pesan bahwa operasi masih berjalan |
| 18 | Setelah operasi selesai / gagal di tengah | semua tombol aktif kembali, kursor normal, aplikasi bisa dipakai lagi |
| 19 | Batalkan dengan exception di tengah (mis. tutup dokumen saat copy) | laporan error muncul, UI tetap responsif (kunci busy tidak tertinggal) |

---

## 8. Referensi API (CorelDRAW)

- `cdrUnit` — `cdrMillimeter = 3`, `cdrCentimeter = 4`, `cdrInch = 1`: https://community.coreldraw.com/sdk/api/draw/17/e/cdrunit
- `Document.Unit` (satuan dokumen untuk posisi/ukuran objek): https://community.coreldraw.com/sdk/api/draw/18/p/ivgdocument.unit
- `Layer.CreateArtisticText`: https://community.coreldraw.com/sdk/api/draw/21/m/layer.createartistictext
- `Transparency.ApplyUniformTransparency`: https://community.coreldraw.com/sdk/api/draw/17/m/transparency.applyuniformtransparency
- `Transparency.Uniform` (workaround agar nilai transparansi "nempel"): https://community.coreldraw.com/sdk/api/draw/17/p/shape.transparency
- `cdrTransparencyAppliedTo` (`cdrApplyToFillAndOutline = 2`): https://community.coreldraw.com/sdk/api/pp/23/e/cdrtransparencyappliedto
- `Shapes.FindShape` / `Shape.Name` (identifikasi watermark): https://community.coreldraw.com/sdk/api/draw/23/m/shapes.findshape · https://community.coreldraw.com/sdk/api/draw/17/p/shape.name
- `cdrMergeMode` — `cdrMergeNormal = 0`: https://community.coreldraw.com/talk/coreldraw-2018/f/coreldraw-2018/62542/macro-for-find-transparency-doesn-t-work-for-powerclipped-objects

---

## 9. Catatan penutup

- Semua kode baru **kompilasi bersih** untuk target .NET Framework 4.8 WinExe x64 dan **tidak menambah referensi NuGet/DLL** — jadi CI yang sekarang tidak perlu diubah.
- Yang belum terverifikasi hanya hal yang mustahil diuji di luar Windows + CorelDRAW: perilaku COM sebenarnya (import logo, transparansi, `MoveToTop`, `SaveAs` batch). Semua panggilan itu sudah defensif (`Try/Catch` per langkah + laporan kegagalan per halaman), jadi kegagalan satu API tidak menggagalkan semuanya.
- Sebelum dipakai ke dokumen produksi: jalankan checklist §7 pada **salinan** dokumen, dan mulai batch dari mode `_wm.cdr` (non-destruktif).
