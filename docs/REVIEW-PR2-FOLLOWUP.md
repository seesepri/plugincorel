# Follow-up Review PR #2 — Pengukuran Input & Verifikasi Simpan

Tanggal: 2026-10-27
Branch: `audit/reliability-copy-verification`
PR: [#2](https://github.com/seesepri/plugincorel/pull/2) (base `main`) — **tidak di-merge otomatis**
Commit kode: `a711c14`

Perbaikan reliabilitas yang sudah ada **dipertahankan** (verifikasi copy,
normalisasi satuan, undo guard, kebijakan status batch, filter output lama).
Tidak ada penulisan ulang modul; perubahan dibatasi pada tiga temuan review.

---

## 1. Ringkasan temuan

| # | Temuan | Tingkat | Status |
|---|---|---|---|
| 1 | `CopyMeasurement.MeasureSelection` memperlakukan **setiap** kegagalan membaca `Count` sebagai "ini satu Shape" | **Tinggi** — bisa mengklaim Verified padahal jumlah objek tidak diketahui | Diperbaiki |
| 2 | `WatermarkEngine` menilai keberhasilan simpan hanya dengan `File.Exists` + `Length > 0` | **Tinggi** — pada mode timpa, file lama otomatis "lulus"; SaveAs yang tidak menulis apa pun tetap berstatus OK | Diperbaiki |
| 3 | Teks ringkasan CSV (`entry.Message`) dirakit di beberapa tempat | Rendah | Diperbaiki (satu perakit) |
| 4 | `MeasureTopShapes` melaporkan "jumlah objek tak terbaca" sebagai "koleksi hanya berisi -1 objek" | Rendah (pesan menyesatkan) | Diperbaiki |

---

## 2. Butir 1 — `CopyMeasurement.vb`

### 2.1 Masalah

```vb
' SEBELUM
Dim total As Integer = SafeCount(rangeOrShape)
If total < 0 Then
    ' Bukan range (Shape tunggal tanpa Count): ukur sebagai satu objek.
    Return MeasureSingleShape(rangeOrShape, result.Source)
End If
```

`SafeCount` mengembalikan `-1` untuk **dua** kondisi yang sangat berbeda:

| Kondisi | Arti | Tindakan yang benar |
|---|---|---|
| objek tidak punya anggota `Count` | memang Shape tunggal | ukur sebagai satu objek |
| `Count` ada tetapi gagal dibaca (COM mati, RPC sibuk, akses ditolak) | jumlah objek **tidak diketahui** | jangan mengukur; hasil = tidak terverifikasi → **Partial** |

Pada kasus kedua, objek yang sama masih bisa membalas `CenterX`/`SizeWidth`
(union bounding box). Kode lama memakainya sebagai geometri "sumber", sehingga
bila geometri union itu kebetulan sama dengan hasil paste, verifikasi bisa
menyimpulkan **Verified** padahal banyak objek tidak pernah diukur.

### 2.2 Perbaikan

| Tambahan | Fungsi |
|---|---|
| `Enum CountReadStatus` (`Ok`, `MemberMissing`, `Failed`) + `TryReadCount` | memisahkan "tidak punya Count" dari "Count gagal dibaca"; `HResult &H80020003`/`&H80020006` dan `MissingMemberException` (termasuk yang bersarang di `TargetInvocationException`) dikenali sebagai member hilang |
| `Enum MeasurementInputKind` (`Empty`, `SingleShape`, `ShapeRange`, `ShapeCollection`, `Unknown`) + `ClassifyInput`, `DescribeInputKind` | jenis input dicatat & dilaporkan; `Unknown` memaksa jalur "tidak terverifikasi" |
| `MeasuredShapeSet.InputKind` | ikut disalin oleh `Scaled()`, jadi jenis input tidak hilang saat konversi satuan |
| `MeasureSelection` ditulis ulang | `MemberMissing` → ukur satu objek; `Ok` → ukur seluruh koleksi; `Failed`/`Unknown` → **tanpa geometri + struktur ditandai `MarkUnreadable()`** dan `Reason` eksplisit |
| `MeasureTopShapes` | jumlah objek layer tak terbaca → `Reason` "jumlah objek pada layer tidak dapat dibaca (...)"; struktur ditandai tidak terbaca |
| `MeasureSingleShape` | tipe/geometri yang gagal dibaca masuk ke `Reason` (tidak lagi hanya "geometri tidak dapat diukur") |
| `SafeCount` | **API lama dipertahankan** (kompatibel), sekarang memakai `TryReadCount` di dalamnya |

Efek pada status: `CopyVerification.VerifyPasted` memberi **Partial** ketika
geometri tidak tersedia (aturan yang sudah ada di PR #2, sekarang benar-benar
terpicu pada kasus ini) — bukan Verified.

### 2.3 Test baru (`Tests/CopyInputKindTests.vb`, 15 test)

| Test | Yang dibuktikan |
|---|---|
| `TryReadCount`: Shape tunggal → `MemberMissing` | "tidak punya Count" ≠ "gagal dibaca" |
| `TryReadCount`: Count melempar → `Failed` | jumlah tidak diketahui, bukan Shape |
| exception bersarang (`TargetInvocationException` → `MissingMemberException`) | deteksi member hilang tahan wrapper |
| `ClassifyInput` | Empty / SingleShape / ShapeRange / ShapeCollection / Unknown dibedakan |
| ShapeRange 2 objek | **dua** objek diukur (bukan satu), jenis dicatat |
| koleksi COM 3 objek | tiga objek diukur |
| Shape tunggal | perilaku lama tetap: geometri + 1 tipe |
| Count gagal + geometri terbaca | geometri **tidak** diklaim (`HasGeometry = False`) |
| Count gagal → `VerifyPasted` | **Partial**, bukan Verified |
| Count & geometri gagal | Partial, tanpa klaim |
| Shape tunggal tanpa tipe/geometri | Partial + `Reason` |
| koleksi kosong | "selection kosong" |
| `MeasureTopShapes` count tak terbaca | alasan eksplisit, tidak mengarang hasil |
| `SafeCount` = -1 untuk kedua kegagalan | kompatibilitas API |
| regresi: ShapeRange yang cocok | tetap **Verified** (perilaku lama tidak rusak) |

---

## 3. Butir 2 — verifikasi penyimpanan (`WatermarkEngine.vb`)

### 3.1 Masalah

```vb
' SEBELUM (ApplyToFolder)
outcome.SaveAttempted = True
Try
    doc.SaveAs(targetPath)
    If Not File.Exists(targetPath) Then Throw ...
    If New FileInfo(targetPath).Length <= 0 Then Throw ...
    outcome.SaveSucceeded = True          ' <-- pemanggilan & bukti dicampur
Catch saveEx ...
End Try
```

Dua akibat:

1. **Mode timpa** (`saveInPlace := True`, `targetPath = sourceFile`): file selalu
   ada dan tidak kosong **sebelum** SaveAs. Jadi SaveAs yang tidak menulis apa
   pun (COM menelan kegagalan, atau file ditolak sistem) tetap menghasilkan
   `SaveSucceeded = True` → status **OK**. Status Verified/OK menjadi tidak jujur.
2. **Mode file baru** dengan output sisa eksekusi sebelumnya: `File.Exists`
   juga tidak membuktikan apa pun tentang simpan terbaru.
3. Keberhasilan **pemanggilan** SaveAs dan keberhasilan **penulisan** tidak bisa
   dibedakan pada laporan/CSV.

### 3.2 Perbaikan — modul baru `SaveOutputVerification.vb`

| Bagian | Isi |
|---|---|
| `SaveOutputSnapshot.Capture(path)` | keadaan file (ada/tidak, ukuran, `LastWriteTimeUtc`) — tidak pernah melempar, kegagalan metadata dicatat di `CaptureError` |
| `SaveOutputMode` | `NewFile` / `OverwriteInPlace` |
| `SaveOutputStatus` | `VerifiedNew`, `VerifiedChanged`, `Missing`, `Empty`, `Unchanged`, `Unreadable` |
| `SaveOutputVerification.Verify(before, after, mode)` | aturan berurutan: keadaan sesudah terbaca → file ada → > 0 byte → keadaan sebelum diketahui → (mode timpa: file sumber harus ada sebelumnya) → file baru muncul = **BARU** → atau ukuran/waktu berubah = **DIPERBARUI** → selain itu **TIDAK BERUBAH** (gagal) |

Integrasi di `ApplyToFolder`:

```
SaveAttempted = True
before = SaveOutputSnapshot.Capture(targetPath)      ' SEBELUM SaveAs
Try doc.SaveAs(targetPath) : SaveCallSucceeded = True
Catch -> pesan "gagal menyimpan: ..."                ' pemanggilan gagal
If SaveCallSucceeded Then
    after = SaveOutputSnapshot.Capture(targetPath)
    check = SaveOutputVerification.Verify(before, after, mode)
    If check.IsVerified -> SaveSucceeded = True, "save OK (BARU/DIPERBARUI)"
    Else                -> pesan "penyimpanan tidak terverifikasi (LABEL): alasan" + Failures
```

Tambahan pada `BatchFileOutcome`: `SaveCallSucceeded`, `SaveVerificationLabel`,
`SaveVerificationReason`; `MessageText()` sekarang membedakan
`save OK (LABEL)` / `SAVE TIDAK TERVERIFIKASI (LABEL)` / `SAVE GAGAL`.
Status akhir tetap dihitung `BatchStatusRules` (satu sumber kebenaran), jadi
file dengan penyimpanan tak terbukti otomatis **GAGAL**, bukan OK.

**Dokumen tidak pernah dibuka ulang.** Verifikasi 100 % memakai metadata sistem
file, sehingga `ActiveDocument` CorelDRAW dan instance COM milik pengguna tidak
tersentuh (diuji: `app.Opened.Count` = jumlah file sumber).

### 3.3 Matriks perilaku

| Mode | Keadaan sebelum | Kejadian SaveAs | Status | Label |
|---|---|---|---|---|
| file baru | tidak ada | menulis | OK | BARU |
| file baru | tidak ada | tidak menulis | GAGAL | TIDAK DITEMUKAN |
| file baru | ada (sisa run lama) | tidak menulis | **GAGAL** | TIDAK BERUBAH |
| file baru | ada (sisa run lama) | menulis | OK | DIPERBARUI |
| **timpa** | ada (file sumber) | tidak menulis | **GAGAL** | TIDAK BERUBAH |
| **timpa** | ada (file sumber) | menulis | OK | DIPERBARUI |
| keduanya | apa pun | menulis 0 byte | GAGAL | KOSONG |
| keduanya | apa pun | melempar | GAGAL | SAVE GAGAL |

### 3.4 Test baru (`Tests/SaveOutputVerificationTests.vb`, 19 test)

Murni (12): baru muncul, ukuran berubah, waktu berubah, tidak berubah, hilang,
0 byte, keadaan sesudah tak terbaca, keadaan sebelum tak terbaca, mode timpa
tanpa file sebelumnya, mode timpa berubah, mode timpa tidak berubah,
`Capture` pada file nyata/path kosong.

WatermarkEngine + FakeCorel (7):

| Test | Yang dibuktikan |
|---|---|
| SaveAs melempar | GAGAL; `SaveCallSucceeded = False`, `SaveSucceeded = False` |
| mode file baru, SaveAs tidak menulis | GAGAL ("tidak ditemukan") |
| mode file baru, output lama masih ada & tidak berubah | GAGAL, label TIDAK BERUBAH, file lama tidak tersentuh, CSV GAGAL |
| **mode TIMPA, SaveAs tidak menulis, file lama masih ada** | **GAGAL** (sebelumnya OK) — inti temuan |
| mode TIMPA menulis | OK, label DIPERBARUI |
| output 0 byte | GAGAL, label KOSONG |
| verifikasi tidak membuka ulang dokumen | `OpenDocument` hanya untuk sumber |

`FakeCorel.vb` ditambah `SaveWritesEmptyFile` untuk skenario 0 byte.

---

## 4. Butir 3 — `BatchLog.vb`

**Apa yang ditemukan:** pada `HEAD` (`4f17da1`) **tidak ada** assignment
`entry.Message` yang benar-benar dobel — hanya satu, di `SummaryEntry`. Yang
memang terduplikasi adalah **perakitan teksnya**: rantai overload
(`SummaryEntry` 5-arg → 6-arg) masing-masing menyusun/menyalin pesan, sehingga
mudah berubah menjadi penulisan ganda saat overload bertambah.

**Perubahan (minimal):** teks ringkasan dipindahkan ke satu fungsi
`BatchLog.BuildSummaryMessage(settings, ok, partialCount, failed)`; `SummaryEntry`
memanggilnya satu kali. Sekarang ada **tepat satu** `entry.Message =` di file itu
(dipatok `static-checks.py`).

**Test baru (1):** `ringkasan CSV: teks status tidak dirakit ganda` — memastikan
`OK 1`, `SEBAGIAN 1`, `GAGAL 1`, dan teks pengaturan masing-masing muncul
**tepat satu kali**, dan isinya identik dengan `BuildSummaryMessage`.

Catatan pengamatan (tidak diubah karena akan mengubah format CSV yang sudah
dipakai): kolom `Halaman`/`Shape` pada baris RINGKASAN berisi jumlah file dan
jumlah OK, bukan halaman/shape. Perlu keputusan produk terpisah bila ingin
dirapikan.

---

## 5. Bukti

### 5.1 Sebelum vs sesudah (diukur di lingkungan ini, tanpa CorelDRAW)

| Pemeriksaan | Sebelum (`4f17da1`) | Sesudah (`a711c14`) |
|---|---|---|
| Regression test | **103 lulus / 0 gagal** | **138 lulus / 0 gagal** (+35) |
| Pemeriksaan statis | **200 pemeriksaan** | **237 pemeriksaan**, 0 gagal (+37) |
| Kompilasi net48 (lokal) | 0 warning / 0 error | **0 warning / 0 error** |
| Build Release x64 Windows (CI) | success (run 40444981041) | **success (run 40452213172)** |

Tidak ada test lama yang dihapus atau dilonggarkan: 103 test lama tetap ada dan
tetap lulus. Satu-satunya pemeriksaan statis lama yang berubah adalah
`penyimpanan diverifikasi (file ada & tidak kosong)` → diganti **dua**
pemeriksaan yang lebih ketat (`SaveOutputVerification.Verify` + snapshot sebelum
SaveAs), bukan dihapus.

### 5.2 Bukti CI Windows x64 — run 40452213172 (commit `a711c14`)

| Langkah | Hasil dari log |
|---|---|
| Static checks | `SEMUA PEMERIKSAAN STATIS LULUS` |
| Regression tests | `Ringkasan: 138 lulus, 0 gagal, 138 total.` |
| Build .NET Framework 4.8 (x64 Release) | `CorelPageCopier -> ...\bin\x64\Release\CorelPageCopier.exe` |
| MSBuild | `Build succeeded.` — `0 Warning(s)`, `0 Error(s)` |

### 5.3 Integration test CorelDRAW 2022 — **NOT RUN**

| Uji | Status | Alasan |
|---|---|---|
| Copy pada dokumen CDR nyata (ShapeRange sumber, selection, paste) | **NOT RUN** | CorelDRAW tidak tersedia di lingkungan ini |
| Simpan mode timpa & mode file baru pada CorelDRAW nyata | **NOT RUN** | idem |
| Verifikasi visual status Verified/Partial/Failed & OK/SEBAGIAN/GAGAL | **NOT RUN** | idem |

Langkah manualnya ada di `PANDUAN-UJI.md` §14. Status **hanya** boleh diubah
menjadi PASS setelah langkah itu benar-benar dijalankan.

---

## 6. Risiko yang tersisa

| Risiko | Dampak | Mitigasi / catatan |
|---|---|---|
| `MemberMissing` disimpulkan dari `MissingMemberException`/`DISP_E_MEMBERNOTFOUND` | Bila COM melaporkan `DISP_E_UNKNOWNNAME` untuk ShapeRange yang sah, objek bisa diukur sebagai satu Shape (union bounds) | Signature tetap mencatat 1 tipe; bila hasil paste > 1 objek → **Partial**, bukan Verified. Tidak ada jalur ke Verified kecuali ShapeRange memang berisi tepat 1 objek |
| Verifikasi simpan memakai `LastWriteTimeUtc` | Pada filesystem dengan granularitas kasar, penulisan yang sangat cepat bisa tidak terdeteksi | Hasilnya **Unchanged → GAGAL** (false negative, bukan false positive). Aman terhadap klaim palsu; risiko hanya laporan GAGAL padahal tersimpan |
| CorelDRAW menulis ke file sementara lalu mengganti path | Ukuran/waktu tetap berubah pada path akhir, sehingga terdeteksi | Diuji lewat FakeCorel; perilaku nyata **NOT RUN** |
| `SaveAs` yang "berhasil" tetapi menulis ke path lain | Terdeteksi sebagai `Missing`/`Unchanged` → GAGAL | Tidak ada fallback diam-diam |
| Baris RINGKASAN CSV: kolom `Halaman`/`Shape` berisi jumlah file/jumlah OK | Membingungkan bila dibaca sebagai halaman/shape | Dibiarkan (format lama); perlu keputusan produk |
| Overload `SummaryEntry` bertambah lagi di masa depan | Bisa memunculkan perakitan pesan ganda | Dipatok static check: hanya satu `entry.Message =` |
| PR bertumpuk (#3, #4, #5 dibuat di atas `4f17da1`) | Perbaikan ini belum ada di branch turunan; saat merge berurutan bisa timbul konflik teks | Rebase/cherry-pick setelah PR #2 di-merge, atau merge berurutan lalu selesaikan konflik di `CopyMeasurement.vb`, `WatermarkEngine.vb`, `BatchLog.vb` |

---

## 7. Yang **tidak** diubah

* Verifikasi copy (Verified/Partial/Failed) dan aturan `CopyVerification.Classify`
  — hanya jalur input yang diperbaiki agar statusnya jujur.
* Normalisasi satuan, `UndoGroupGuard`, `DocumentUnitScope`.
* Kebijakan file batch (`BatchFilePolicy`), filter output lama, benturan nama.
* `BatchStatusRules` sebagai satu sumber status akhir.
* Format CSV, teks dialog UI, tata letak dialog watermark.
