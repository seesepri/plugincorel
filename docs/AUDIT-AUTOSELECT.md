# Audit & Perbaikan Auto Select — Corel Page Copier

Tanggal: 2026-10-27
Cakupan: `CorelPageCopier/Program.vb` (bagian Auto Select), modul baru
`ShapeClassifier.vb`, `ShapeNode.vb`, `ShapeScanner.vb`, `SelectionNavigation.vb`.
Kaitannya dengan audit sebelumnya: `docs/AUDIT-RELIABILITY.md` (Copy & batch).
Dokumen ini **hanya** membahas Auto Select.

---

## 1. Ringkasan

Auto Select lama bekerja dengan satu fungsi rekursif (`CollectMatchingShape`) yang
mencampur empat hal sekaligus: pembacaan COM, keputusan kategori, penyaringan
watermark, dan pemilihan objek. Akibatnya beberapa kesalahan tidak bisa diuji tanpa
CorelDRAW, dan sebagian keputusan salah tanpa terlihat.

Perbaikan memisahkan tanggung jawab menjadi:

| Lapisan | File | Sifat |
|---|---|---|
| Klasifikasi objek | `ShapeClassifier.vb` (`ShapeCategory`, `ShapeFacts`, `Classify`) | murni, tanpa COM, unit-tested |
| Penelusuran pohon objek | `ShapeScanner.vb` (`Scan`, `Visit`, `VisitPowerClip`) | murni, lewat antarmuka `IShapeNode` |
| Adapter COM | `ShapeNode.vb` (`ShapeNodeAdapter`) | late-bound, satu-satunya bagian yang menyentuh CorelDRAW |
| Navigasi hasil | `SelectionNavigation.vb` (`MatchNavigator`, `SelectionOutcome`) | murni, unit-tested |
| Jembatan UI | `Program.vb` (`AutoSelectClickedCore`, `ScanPageForUi`, `ApplySelection`) | COM + WinForms saja |

Prinsip yang dipegang: **Auto Select tidak pernah menghapus atau mengubah objek**
(tidak ada `Delete`, tidak ada perubahan properti), watermark tidak pernah
terpilih, dan setiap kegagalan baca menjadi diagnostik yang dilaporkan — bukan
`Catch` kosong.

---

## 2. Metode & batasan

* Audit dilakukan dengan membaca kode lama, lalu setiap temuan diberi perbaikan
  minimum (tanpa menulis ulang fitur lain) dan satu/lebih regression test yang
  gagal pada perilaku lama.
* Klasifikasi dan penelusuran dipindahkan ke kode murni supaya bisa diuji pada
  lingkungan **tanpa** CorelDRAW (net8.0, objek palsu `FakeShapeNode`).
* Kompatibilitas target: CorelDRAW 2022, Windows x64, .NET Framework 4.8,
  late-bound COM (tanpa referensi interop). Tidak ada API baru yang dipakai.

**Batasan yang harus dinyatakan jujur:** tidak ada CorelDRAW 2022 di lingkungan
audit, jadi **integration test di aplikasi nyata berstatus NOT RUN**. Bukti yang
ada bersifat kompilasi + unit/regression test + pemeriksaan statis (lihat §7).

---

## 3. Temuan dan perbaikan

Nomor mengikuti daftar temuan pada tugas audit.

### 3.1 Kategori "vektor" menelan apa saja (#1/#4)

**Gejala lama:** `MatchesCategory` untuk kategori vektor memakai `Case Else`,
yaitu "semua yang bukan teks, bukan bitmap, bukan grup". Garis bantu
(*guideline*), objek OLE, dan *selection shape* ikut terpilih.

**Perbaikan:** vektor didefinisikan eksplisit dan terdokumentasi di
`ShapeClassifier`:

* **termasuk**: rectangle, ellipse, polygon, kurva (terbuka maupun tertutup), dan
  kode tipe tak dikenal yang dilaporkan sebagai "vektor khusus" **dengan catatan
  verifikasi**;
* **tidak termasuk**: `noShape`, bitmap, teks, grup, guideline, objek OLE,
  selection shape, objek yang tipenya tidak terbaca.

Test: *vektor mencakup rectangle, ellipse, polygon, kurva*,
*vektor TIDAK mencakup guideline / OLE / teks / bitmap / grup*,
*kode tipe tak dikenal → vektor khusus + catatan verifikasi*.

### 3.2 Teks artistik vs paragraf tidak dibedakan (#1)

**Gejala lama:** `shape.Text.Type` tidak pernah dibaca; semua teks dianggap sama,
sehingga pilihan pengguna tidak bisa dibatasi ke salah satunya.

**Perbaikan:** `ShapeFacts.TextKind` menyimpan `TextArtistic` / `TextParagraph`
(dibaca dari `Text.Type`), kategori **Teks** mencakup keduanya, dan label hasil
menyebut jenisnya ("Teks artistik" / "Teks paragraf"). Bila `Text.Type` tidak
terbaca, objek tetap dihitung sebagai teks **dan** diberi catatan — tidak
ditebak diam-diam.

### 3.3 Kurva terbuka menebak dari bounding box (#3)

**Gejala lama:** status `Curve.Closed` dibaca di dalam `Try` dengan `Catch` kosong;
objek yang gagal dibaca dianggap kurva terbuka/tertutup secara acak.

**Perbaikan:** `ShapeFacts.CurveClosed` + `CurveClosedKnown`. Kurva terbuka hanya
dipilih bila `Curve.Closed = False` **terbaca**. Bila tidak terbaca, objek tidak
dimasukkan ke kategori kurva terbuka dan alasannya muncul di catatan pemindaian
("status Curve.Closed tidak dapat dibaca"). Tidak ada perhitungan geometri
pengganti.

### 3.4 Nested group: kedalaman & kegagalan baca anak (#6)

**Gejala lama:** `depth > 32` langsung `Return` tanpa jejak; `Catch` pada
pembacaan anak kosong.

**Perbaikan:** `ShapeNodeAdapter.MaxDepth = 32` (eksplisit, diteruskan lewat
`ShapeScanOptions.MaxDepth`); pemotongan menaikkan `DepthTruncated` dan menulis
diagnostik "melebihi batas 32 — isi objek ini tidak dipindai". Kegagalan membaca
anak menaikkan `ChildrenFailures` + diagnostik. Pemindaian **lanjut** ke objek
lain; tidak ada exception yang lolos ke UI, dan tidak ada exception yang hilang.

### 3.5 PowerClip tidak pernah dipindai (#7)

**Gejala lama:** hanya `shape.Shapes` yang dibaca; isi PowerClip tidak ada di
`Shapes`, tetapi laporan lama tetap menyatakan "sudah dipindai".

**Perbaikan:** `ShapeNodeAdapter` membuka isi PowerClip berlapis
(`PowerClip.Shapes` → `PowerClip.Contents.Shapes` → `PowerClip.Contents`).
`ShapeScanner.VisitPowerClip` menelusuri isinya, dan **target seleksinya adalah
container** (isi PowerClip tidak dapat dipilih sendiri). Kegagalan membuka PowerClip
menaikkan `PowerClipFailures` dan menulis "isi PowerClip TIDAK dipindai" — angka
tidak diklaim lengkap.

### 3.6 Watermark (#8)

| Perilaku lama | Perilaku baru |
|---|---|
| Grup yang memuat watermark dilewati **seluruhnya** pada semua kategori → saudara yang sah ikut hilang | Kategori spesifik: watermark (dan isinya) dilewati satu per satu; saudara yang sah tetap terpilih |
| Mode "semua objek" menambahkan grup tanpa memeriksa watermark di dalamnya | Mode "semua objek": grup yang memuat watermark dilewati utuh + dicatat (karena grup dipilih sebagai satu kesatuan, memilihnya berarti memilih watermark) |
| Tidak ada angka/diagnostik | `WatermarksSkipped` + diagnostik per objek |

Identitas watermark memakai satu helper bersama (`WatermarkLayout.IsWatermarkIdentity`),
sehingga aturan `WM_*`/layer `WATERMARK` sama dengan fitur watermark lain dan
tidak digandakan. Pemeriksaan bersifat case-insensitive.

### 3.7 Navigasi hasil memakai kategori saat ini (#9)

**Gejala lama:** tombol "hasil berikutnya" menghitung ulang dengan
`categoryBox.SelectedIndex` **saat itu**, bukan kategori yang dipakai ketika
memindai; `matchingPageIndices` bisa basi. Akibatnya halaman terlewat atau
terulang.

**Perbaikan:** `MatchNavigator` menyimpan **snapshot** (kategori + cakupan +
daftar halaman hasil) saat pemindaian selesai. Perubahan ComboBox setelahnya
menghasilkan `StaleSnapshot` (pengguna diminta memindai ulang), bukan lompatan
diam-diam. `MoveNext` berhenti di hasil terakhir (`AtEnd`) tanpa kembali ke awal;
halaman hasil yang hilang → `PageMissing` dan posisi tidak dimajukan.

### 3.8 Seleksi kosong tetap menghapus seleksi lama (#10)

**Gejala lama:** `ApplySelection` selalu memanggil `ClearSelection()` lebih dulu,
lalu mengembalikan **jumlah panggilan** `AddToSelection` sebagai "jumlah terpilih".
Bila tidak ada yang cocok, seleksi pengguna hilang tanpa hasil apa pun; bila
sebagian gagal, angkanya tetap terlihat benar.

**Perbaikan:**

* Tidak ada target → `SelectionOutcome.NotRun(...)`: `ClearSelection` **tidak**
  dipanggil, seleksi lama utuh.
* Jumlah terpilih diverifikasi dari `corelApp.ActiveSelectionRange.Count`.
  Bila berbeda dari target, laporan menyebut selisihnya dan dialog memakai ikon
  peringatan; bila jumlah tidak dapat dibaca, status "tidak terverifikasi"
  (bukan klaim sukses).
* Kegagalan `ClearSelection` membatalkan operasi (tidak ada fallback `Delete`).

### 3.9 Objek pada layer terkunci diklaim terpilih

Objek pada layer yang tidak dapat diedit tetap **dihitung sebagai cocok**
(angka di laporan jujur: "N cocok"), tetapi **dikeluarkan dari target seleksi**
(`LockedLayerSkipped`), karena CorelDRAW tidak akan memilihnya. Test:
*layer terkunci: objek tidak diklaim terpilih + dicatat*.

### 3.10 Busy guard dan pemulihan UI

`AutoSelectClicked` membungkus pekerjaan dengan `TryBeginOperation("Auto Select")`
/ `EndOperation()` di `Finally`, sehingga klik ganda dan re-entrancy `DoEvents`
tidak mungkin terjadi, dan status busy selalu dilepas meskipun pemindaian gagal.
Kegagalan apa pun dijalankan lewat `ShowError` dan mengosongkan snapshot navigasi
(`matchNavigator.Reset()`), jadi tidak ada sisa hasil setengah jadi.

---

## 4. Definisi kategori (dipakai UI & test)

| Kategori (ComboBox) | Isi | Catatan yang tampil di UI |
|---|---|---|
| Teks | teks artistik + teks paragraf | jenis teks disebut per objek |
| Bitmap | objek bitmap, termasuk di dalam grup dan di dalam PowerClip | target untuk isi PowerClip = container |
| Kurva terbuka | kurva dengan `Curve.Closed = False` | status tak terbaca → tidak dipilih + catatan |
| Vektor | rectangle, ellipse, polygon, kurva (terbuka/tertutup), kode tak dikenal ("vektor khusus") | guideline/OLE/selection shape dikecualikan; kode tak dikenal diberi catatan verifikasi |
| Semua objek | objek level atas; **grup utuh sebagai satu kesatuan** | grup ber-watermark dilewati utuh |

Semua kategori: watermark (`WM_*` / layer `WATERMARK`) **tidak pernah** terpilih.

---

## 5. Bukti kode: apa yang tidak lagi dilakukan

Diperiksa juga oleh `tools/static-checks.py` (bagian "Auto Select"):

* `CollectMatchingShape` tidak ada lagi di `Program.vb`;
* `ApplySelection` tidak memuat panggilan `.Delete(`;
* `ClearSelection` hanya muncul **setelah** pemeriksaan `targets.Count = 0`;
* jumlah terpilih diverifikasi lewat `ActiveSelectionRange`;
* `AutoSelectClicked` memakai `TryBeginOperation` **dan** `EndOperation`.

---

## 6. Dataset uji

### 6.1 Dataset otomatis (`FakeShapeNode.MixedDataset`)

16 node akar pada satu halaman:

| # | Objek | Keterangan |
|---|---|---|
| 1–2 | teks artistik, teks paragraf | AS-1 |
| 3 | bitmap level atas | AS-2 |
| 4–5 | kurva terbuka, kurva tertutup | AS-3 |
| 6–8 | rectangle, ellipse, polygon | AS-4 |
| 9–10 | guideline, objek OLE | harus dikecualikan dari vektor |
| 11–12 | watermark `WM_TEXT`, bitmap di layer `WATERMARK` | tidak boleh terpilih |
| 13 | grup: 2 bitmap + 1 teks paragraf | bitmap di dalam grup |
| 14 | PowerClip: 1 bitmap | target = container |
| 15 | grup bersarang 2 level: kurva terbuka | kedalaman |
| 16 | objek dengan tipe tak terbaca | dilaporkan, tidak ditebak |

Angka yang diharapkan (dan diuji):

| Dataset | Kategori | Hasil |
|---|---|---|
| AS-1 | Teks | 3 objek (2 level atas + 1 dalam grup) |
| AS-2 | Bitmap | 4 objek, 1 di antaranya di dalam PowerClip → target container |
| AS-3 | Kurva terbuka | 2 objek (kurva tertutup **tidak** ikut) |
| AS-4 | Vektor | 6 objek; guideline & OLE **tidak** ikut |
| AS-5 | Semua objek | 13 objek (16 − 2 watermark − 1 tak terbaca), 3 grup utuh |
| AS-6 | Teks pada dokumen tanpa teks | 0 objek, 0 target, tidak ada operasi destruktif |

### 6.2 Dataset manual (CorelDRAW 2022)

Langkah membuat dokumen uji nyata beserta angka yang diharapkan ada di
`PANDUAN-UJI.md` §11 (dataset **AS-1 … AS-6**). Test otomatis memakai struktur
yang sama supaya angka di UI dan angka di test dapat dibandingkan langsung.

---

## 7. Status verifikasi

| Verifikasi | Status | Bukti |
|---|---|---|
| Kompilasi produksi .NET Framework 4.8 (net48, `Option Strict Off`) | **PASS** | `dotnet build` proyek kompilasi setara (semua `CorelPageCopier/*.vb`), 0 error / 0 warning |
| Regression test (net8.0, 157 test) | **PASS** | `Ringkasan: 157 lulus, 0 gagal, 157 total` |
| Pemeriksaan statis (`tools/static-checks.py`) | **PASS** | 255 pemeriksaan lulus, 0 gagal |
| Build Windows x64 (MSBuild, `CorelPageCopier.vbproj`) | **NOT RUN di lokal** — dijalankan lewat CI GitHub Actions `windows-latest` | lihat log run di PR |
| Integration test CorelDRAW 2022 (dokumen nyata) | **NOT RUN** | CorelDRAW tidak tersedia di lingkungan audit; ikuti `PANDUAN-UJI.md` §11 |

Tidak ada status "PASS" untuk hal yang tidak dijalankan. Angka 157/255 di atas
adalah keluaran nyata dari perintah yang tercantum di §8.

---

## 8. Cara menjalankan ulang

```bash
# regression test (Linux/macOS/Windows dengan .NET 8 SDK)
dotnet run --project tests/CorelPageCopier.RegressionTests -c Release

# pemeriksaan statis
python3 tools/static-checks.py

# kompilasi net48 (opsional, butuh paket reference assemblies)
dotnet build <proyek-kompilasi-net48>.vbproj -c Release
```

Build Windows x64 yang sebenarnya:

```powershell
msbuild CorelPageCopier\CorelPageCopier.vbproj /p:Configuration=Release /p:Platform=x64
```

---

## 9. Risiko sisa / catatan kompatibilitas

1. **Seleksi objek di dalam grup.** Bila kategori spesifik menemukan objek di
   dalam grup, target seleksinya adalah objek anak. CorelDRAW 2022 umumnya
   menerima `AddToSelection` pada anak grup, tetapi bila ternyata yang terpilih
   adalah grupnya, `ActiveSelectionRange.Count` akan berbeda dan laporan
   menandainya sebagai **tidak terverifikasi** (ikon peringatan) — bukan klaim
   sukses. Butuh konfirmasi pada uji manual AS-2b/AS-1.
2. **Deteksi PowerClip.** Jalur `PowerClip.Shapes` → `Contents.Shapes` →
   `Contents` mencakup struktur yang dikenal di CorelDRAW 2018–2024; struktur
   eksotis (mis. efek `Clip` pada objek non-container) dilaporkan sebagai
   `PowerClipFailures`, tidak dipaksakan.
3. **Kode tipe objek tak dikenal** dipetakan ke "vektor khusus" **dengan catatan
   verifikasi**, bukan ditebak sebagai vektor biasa. Bila hasil uji manual
   menunjukkan objek tertentu salah kategori, tambahkan pemetaan eksplisit di
   `ShapeClassifier` beserta test-nya.
4. **Batas kedalaman 32** dipertahankan (sama dengan perilaku lama) agar tidak
   ada perubahan perilaku pada dokumen normal; yang berubah hanya pelaporannya.

---

## 10. Peta test → temuan

| Test suite | Temuan yang dijaga |
|---|---|
| `AutoSelectClassifierTests` (19 test) | #1, #3, #4, #8, definisi kategori, label UI |
| `ShapeScannerTests` (20 test) | #2, #5, #6, #7, #8, #10, layer terkunci, PowerClip |
| `SelectionNavigationTests` (14 test) | #9, #10, verifikasi seleksi, berhenti di hasil terakhir |
| `tools/static-checks.py` | struktur pemisahan lapisan, tidak ada `Delete`, urutan `ClearSelection` |
