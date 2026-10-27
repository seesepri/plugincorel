# Review Follow-up Kedua — PR #2 (`fix(reliability)`)

Tanggal: 2026-10-27 · Basis: `origin/audit/reliability-copy-verification` @ `f0deb26`
(PR #2 head, CI run [40452373897](https://github.com/seesepri/plugincorel/actions/runs/40452373897) = success)

Dokumen ini adalah hasil **review lanjutan (putaran kedua)** atas PR #2, bukan pengganti
`docs/AUDIT-RELIABILITY.md` (temuan A–F) maupun `docs/REVIEW-PR2-FOLLOWUP.md` (follow-up
pertama). Semua perbaikan reliabilitas yang sudah ada **dipertahankan**; tidak ada rewrite,
tidak ada fitur yang diubah tanpa alasan terukur.

---

## 0. Ringkasan satu paragraf

Putaran pertama sudah memisahkan status **Verified/Partial/Failed** dan **OK/SEBAGIAN/GAGAL**,
tetapi masih ada empat celah yang bisa membuat status **lebih benar daripada buktinya**:

1. objek **tanpa** `Count` langsung dianggap “satu Shape” (tanpa bukti bentuk), dan
   `NotSupportedException` ikut dianggap “anggota hilang”;
2. pengukuran yang **terpotong** tetap meninggalkan sidik jari struktur yang tampak utuh;
3. status penyimpanan hanya memakai **“ada + > 0 byte”** dan perubahan ukuran/waktu —
   berkas yang **isinya bukan CDR** atau yang **ditulis ulang dengan panjang sama** tidak
   dapat dibedakan;
4. `Program.vb` masih memakai `CInt(selection.Count)` / `CInt(layerShapes.Count)`, sehingga
   kegagalan membaca `Count` menjadi **exception mentah** atau pesan yang menyesatkan.

Keempatnya diperbaiki minimal, ditambah 29 regression test baru, dan `BatchLog.vb` (butir 3)
diverifikasi bersih lalu dikunci oleh test.

---

## 1. Status validasi (dipisah per tahap, jujur)

| Tahap | Status | Bukti |
|---|---|---|
| `python3 tools/static-checks.py` | **PASS** | **279 pemeriksaan, 0 FAIL**, exit 0 (237 → 279; +42 pemeriksaan baru) |
| Regression/unit test **tanpa CorelDRAW** | **PASS** | **167 test, 0 gagal** (sebelumnya 138; **+29 test baru**), exit 0 |
| Compile seluruh 13+ file produksi (net48, `Option Strict Off`, x64) | **PASS (compile-check)** | `dotnet build tests/CorelPageCopier.CompileCheck -c Release` → `Build succeeded, 0 Warning, 0 Error` |
| **Build Release x64 via MSBuild di Windows** | **NOT RUN (lingkungan ini)** | Tidak ada Windows/MSBuild di lingkungan review. Build Windows terakhir yang PASS ada pada PR head sebelumnya (CI run 40452373897, artifact `CorelPageCopier-Windows-x64`). Commit perbaikan ini **wajib** dipush agar CI `windows-latest` membangunnya; lihat §8. |
| **Integration test CorelDRAW 2022** | **NOT RUN** | CorelDRAW tidak tersedia di lingkungan mana pun saat ini. Prosedur manual: `PANDUAN-UJI.md` §10, §14, dan **§15 (baru)**. |

> Tidak ada klaim “sudah diuji di Windows/CorelDRAW” di dokumen ini. Dua baris terakhir
> sengaja NOT RUN sampai benar-benar dijalankan.

---

## 2. Temuan baru & perbaikan

### G. Bukti penyimpanan masih terlalu lemah (`SaveOutputVerification.vb`, `WatermarkEngine.vb`)

**G1 — “ada + tidak kosong” bukan bukti sah.**
Kode putaran pertama memverifikasi `File.Exists` + `Length > 0` + perubahan ukuran/waktu.
Berkas 12 KB berisi teks acak, atau keluaran setengah jadi, akan lolos sebagai **OK**.
*Perbaikan:* `SaveOutputSnapshot` kini mengklasifikasi **bentuk isi** dari 12 byte pertama:

- `RIFF` + form type `CD…` → berkas CorelDRAW klasik (≤ X3),
- `PK\x03\x04` (ZIP) → paket CDR modern (**X4 ke atas, termasuk CorelDRAW 2022**),
- lainnya → `TIDAK DIKENALI` → status baru **`Invalid` (TIDAK VALID)** → tidak pernah OK.

Pemeriksaan hanya aktif bila diminta (`requireCorelSignature`) dan engine hanya memintanya
untuk `.cdr/.cdt/.cdx` (`BatchFilePolicy.HasCorelExtension`) — output lain (mis. PDF proof)
tidak terpengaruh.

**G2 — bukti penulisan hanya ukuran + waktu modifikasi.**
Banyak filesystem menyimpan waktu dengan granularitas 1–2 detik. Berkas yang benar-benar
ditulis ulang dengan panjang sama bisa terlihat “tidak berubah”. Ini **terbukti nyata di
lingkungan review**: pada filesystem sandbox, dua penulisan berurutan menghasilkan
`LastWriteTimeUtc` **identik** (lihat test `mode TIMPA: SaveAs menulis ulang CDR valid`).
*Perbaikan:* snapshot opsional menghitung **sidik jari SHA-256** atas panjang + 4 KiB awal +
4 KiB akhir; perbedaan sidik jari dihitung sebagai bukti perubahan (`changedByContent`).
Arah kesalahan tetap aman: sidik jari hanya bisa **menambah** bukti, tidak pernah
mengubah “tidak berubah” menjadi sukses tanpa perubahan nyata.

**G3 — mode timpa bisa merusak sumber tanpa peringatan.**
Pada mode timpa, target = berkas sumber. Bila `SaveAs` menghasilkan berkas kosong/hilang/
tidak valid, kode lama hanya menulis “tidak terverifikasi” ke daftar kegagalan.
*Perbaikan:* engine menambahkan **peringatan eksplisit** bahwa berkas sumber kemungkinan
tidak lengkap/rusak dan harus dipulihkan dari cadangan.

**G4 — file yang masih dipegang CorelDRAW.**
Membaca isi berkas tepat setelah `SaveAs` bisa gagal karena lock.
*Perbaikan:* `FileStream` memakai `FileShare.ReadWrite | FileShare.Delete` + 3 percobaan
(jeda 40 ms). Bila tetap gagal: bentuk **tidak** dipastikan, keterbatasan ditulis di alasan,
dan keputusan tetap memakai metadata/sidik jari — tidak ada klaim “valid”.

**Test baru (16):** `SaveOutputContentTests` — klasifikasi RIFF/CDR vs ZIP vs asing, output
tidak valid (mode baru & timpa), sidik jari (isi berubah walau ukuran+waktu sama), isi tak
terbaca, snapshot nyata, mode file baru (tidak valid / ZIP / target lama), mode timpa
(0 byte, isi tidak valid, tulis ulang valid).

### H. Bukti jenis input pengukuran (`CopyMeasurement.vb`, `Program.vb`)

**H1 — “tidak punya `Count`” terlalu cepat disebut satu Shape.**
`ClassifyInput` mengembalikan `SingleShape` untuk setiap objek yang `Count`-nya “member
missing”. Bila objek itu sebenarnya koleksi rusak atau objek asing yang kebetulan
mengekspos `CenterX/SizeWidth`, geometrinya dipakai sebagai bukti → **Verified palsu**.
*Perbaikan:* kandidat `SingleShape` wajib punya **bukti bentuk**, yaitu `Shape.Type` terbaca
(`LooksLikeShape`). Tanpa bukti → `Unknown` → hasil tanpa geometri + struktur tidak terbaca
→ **Partial**.

**H2 — `NotSupportedException` bukan “anggota hilang”.**
Anggota yang ada tetapi tidak didukung tidak membuktikan objek tidak punya `Count`.
*Perbaikan:* `IsMemberMissing` hanya menerima `MissingMemberException` dan
`DISP_E_MEMBERNOTFOUND`/`DISP_E_UNKNOWNNAME`. Sisanya `Failed` → jumlah objek tidak diketahui
→ **Partial**.

**H3 — pengukuran terpotong meninggalkan struktur yang tampak utuh.**
`MeasureSlice` berhenti saat objek ke-*k* tidak terbaca, tetapi sidik jari (jumlah + tipe)
tetap “readable” untuk objek yang sempat terbaca.
*Perbaikan:* kegagalan pengukuran menandai `Signature.MarkUnreadable()` sehingga struktur
tidak pernah dianggap bisa dibandingkan secara utuh.

**H4 — fallback `MeasurePasted` memakai `SafeCount`.**
`SafeCount` mengembalikan `-1` baik untuk “tidak punya Count” maupun “gagal dibaca”, sehingga
kecocokan jumlah bisa dinilai dari nilai yang tidak bermakna.
*Perbaikan:* fallback hanya dipakai bila `TryReadCount` **Ok** dan jumlahnya sama; bila
`Count` gagal dibaca, alasan “fallback selection tidak dipakai (…)” masuk laporan.

**H5 — `Program.vb` masih memakai `CInt` mentah.**
`CInt(selection.Count)` membuat kegagalan membaca `Count` menjadi exception umum
(“Operasi dihentikan oleh kesalahan”), dan kegagalan COM saat membaca seleksi dilaporkan
sebagai “tidak ada seleksi”. Sekarang:

| Kondisi | Perilaku baru |
|---|---|
| `ActiveSelectionRange` melempar | pesan error COM jelas, **dokumen tidak diubah** |
| `Count` terbaca = 0 | “Tidak ada seleksi” (seperti sebelumnya) |
| `Count` gagal dibaca | copy **dibatalkan sebelum** dokumen disentuh + alasan |
| objek tanpa `Count` (Shape tunggal) | diteruskan; pengukuran menentukan Partial/Verified |
| `Count` layer gagal **sebelum** paste | halaman **Failed**, tanpa klaim berhasil |
| `Count` layer gagal **sesudah** paste | halaman **Partial** (paste terjadi, verifikasi tidak lengkap) |

**Test baru (9):** `CopyInputEvidenceTests` — objek tanpa `Count` & tanpa tipe, `Count`
`NotSupportedException`, Shape tunggal sah (tidak regresi), koleksi yang “berbohong”
(`Count=3`, objek ke-3 hilang), fallback `MeasurePasted` (gagal baca / jumlah salah / cocok),
kompatibilitas `SafeCount`.

### I. `BatchLog.vb` — assignment `entry.Message` duplikat (butir 3)

Diverifikasi: berkas ini **sudah bersih** setelah putaran pertama — satu-satunya perakitan
teks ada di `BatchLog.BuildSummaryMessage` dan hanya ada **satu** assignment `entry.Message`.
Yang ditambahkan pada putaran ini adalah **pengunci perilaku**: `BatchLogSummaryTests` (4 test)
memastikan teks ringkasan muncul **tepat sekali** di baris CSV, kedua overload `SummaryEntry`
konsisten, dan escaping `;` tidak menggandakan pesan. Pemeriksaan statis juga dinaikkan dari
“satu assignment” menjadi “masih satu assignment setelah perubahan ini”.

Tidak ada perubahan perilaku pada `BatchLog.vb` (0 baris kode produksi diubah).

---

## 3. Daftar perubahan berkas

| Berkas | Perubahan | Sifat |
|---|---|---|
| `CorelPageCopier/CopyMeasurement.vb` | `LooksLikeShape`, syarat bukti untuk `SingleShape`, `NotSupportedException` bukan member-missing, `MeasureSlice` menandai struktur terpotong, `MeasurePasted` memakai `TryReadCount` | minimal, aditif |
| `CorelPageCopier/SaveOutputVerification.vb` | status `Invalid`, klasifikasi bentuk (RIFF/CDR vs ZIP), sidik jari SHA-256, baca isi dengan `FileShare.ReadWrite` + retry, `requireCorelSignature` | minimal, aditif |
| `CorelPageCopier/BatchFilePolicy.vb` | `HasCorelExtension` (hanya `.cdr/.cdt/.cdx` diperiksa bentuknya) | aditif |
| `CorelPageCopier/WatermarkEngine.vb` | snapshot dengan isi, `requireCorelSignature`, peringatan kerusakan sumber pada mode timpa | perubahan kecil di blok save |
| `CorelPageCopier/Program.vb` | `Count` seleksi & layer dibaca dengan status eksplisit; pesan COM vs “tidak ada seleksi” dipisah; `Partial` untuk paste yang tak terverifikasi | perubahan kecil di jalur copy |
| `CorelPageCopier/BatchLog.vb` | **tidak diubah** (sudah bersih) | — |
| `tests/.../Tests/CopyInputEvidenceTests.vb` | **baru**, 9 test | baru |
| `tests/.../Tests/SaveOutputContentTests.vb` | **baru**, 16 test | baru |
| `tests/.../Tests/BatchLogSummaryTests.vb` | **baru**, 4 test | baru |
| `tests/.../FakeCorel.vb` | `SaveAs` menulis berkas berbentuk CDR sah + token unik; opsi `SaveWritesInvalidContent`, `SaveWritesZipContainer` | fixture |
| `tests/.../Tests/SaveOutputVerificationTests.vb` | 2 test lama: isi berkas “lama” dibuat **CDR sah** (bukan teks) — lihat §7 | fixture |
| `tests/.../TestMain.vb` | mendaftarkan 3 suite baru | test |
| `tools/static-checks.py` | +42 pemeriksaan untuk temuan G/H/I; helper `code_only()` | checks |

---

## 4. Semantik status yang dijaga

### Per halaman copy

| Status | Syarat (tidak berubah dari putaran pertama, diperketat) |
|---|---|
| **Verified** | jumlah objek bertambah, geometri sumber & hasil terukur dan cocok dalam **0,5 mm**, struktur (jumlah + tipe) cocok |
| **Partial** | paste terjadi, tetapi ada fakta yang tidak dapat diverifikasi: jumlah objek tak terbaca, jenis input tidak dapat dipastikan, tipe objek tak terbaca, pengukuran terpotong, struktur berbeda, geometri tak terukur |
| **Failed** | tidak ada objek hasil, jumlah objek layer tak terbaca **sebelum** paste, atau geometri jelas berbeda di luar toleransi |

Partial **tidak pernah** masuk hitungan sukses (`CopyRunSummary`).

### Per file batch

| Status | Syarat |
|---|---|
| **OK** | semua halaman target diterapkan **dan** penyimpanan terbukti (baru / berubah ukuran-waktu-isi / bentuk berkas sah bila diperiksa) |
| **SEBAGIAN** | penyimpanan terbukti, tetapi ada halaman gagal |
| **GAGAL** | fatal error, tidak ada halaman diterapkan, save tidak dipanggil, save tidak terbukti, **atau bentuk berkas salah (TIDAK VALID)** |

---

## 5. Bukti pengujian (perintah + keluaran)

```text
$ python3 tools/static-checks.py
...
SEMUA PEMERIKSAAN STATIS LULUS        # 279 pemeriksaan PASS, 0 FAIL, exit 0

$ dotnet run --project tests/CorelPageCopier.RegressionTests
Ringkasan: 167 lulus, 0 gagal, 167 total.
Catatan: integration test CorelDRAW 2022 TIDAK dijalankan oleh test ini (NOT RUN).

$ dotnet build tests/CorelPageCopier.CompileCheck/CorelPageCopier.CompileCheck.vbproj -c Release
Build succeeded. 0 Warning(s) 0 Error(s)
```

Rincian suite baru:

| Suite | Jumlah | Inti yang dibuktikan |
|---|---|---|
| `Bukti jenis input pengukuran (review kedua)` | 9 | tanpa `Count` belum tentu Shape; `NotSupportedException` → Partial; koleksi “berbohong” → struktur tak utuh; fallback `MeasurePasted` |
| `Bukti penyimpanan: bentuk isi & sidik jari (review kedua)` | 16 | RIFF/CDR vs ZIP vs asing; output tidak valid → GAGAL; sidik jari; mode file baru & mode timpa |
| `Log CSV: teks ringkasan dirakit satu kali (review kedua)` | 4 | teks ringkasan tepat sekali di CSV; kedua overload konsisten |

138 (lama, semuanya tetap lulus) + 29 (baru) = **167**.

Beberapa bukti yang layak disorot:

- **tanpa `Count` & tipe tak terbaca** → `Unknown`, tanpa geometri, **Partial** (dulu: “satu Shape” → Verified).
- **`Count` NotSupportedException** → `Failed` → **Partial**.
- **koleksi `Count=3` dengan objek ke-3 hilang** → `Signature.IsComparable=False`, **bukan Verified**.
- **mode timpa, `SaveAs` menulis 0 byte** → `KOSONG`, status **GAGAL**, dan **peringatan** “berkas sumber kemungkinan TIDAK LENGKAP”.
- **mode file baru, `SaveAs` menulis isi non-CDR** → `TIDAK VALID`, **GAGAL**, CSV tidak memuat `;OK;`.
- **CDR modern (ZIP) diterima** → **OK/BARU** (CorelDRAW 2022 menulis paket ZIP, bukan RIFF).
- **ukuran & waktu sama tetapi isi berbeda** → **DIPERBARUI**, bukan “tidak berubah”.

---

## 6. Risiko yang masih tersisa (jujur)

1. **Pemeriksaan bentuk bersifat dangkal.** Hanya 12 byte pertama; berkas ZIP apa pun
   (walau bukan CDR) akan lolos klasifikasi bentuk. Ini **menolak** keluaran yang jelas rusak,
   bukan **menjamin** isi CDR valid. Validasi isi sejati hanya mungkin dengan membuka dokumen —
   yang sengaja tidak dilakukan (akan mengganggu instance COM pengguna).
2. **Isi berkas tidak terbaca** (masih terkunci CorelDRAW setelah 3 percobaan): bentuk tidak
   dipastikan. Keputusan tetap dibuat dari metadata/sidik jari, dan keterbatasan dicatat di
   alasan. Pada kasus ini status bisa `BARU/DIPERBARUI` tanpa pemeriksaan bentuk.
3. **Sidik jari = sebagian isi** (panjang + 4 KiB awal + 4 KiB akhir). Penulisan ulang yang
   hanya mengubah byte di tengah, dengan panjang **persis** sama dan timestamp identik, masih
   bisa terlewat. Untuk CDR praktis tidak terjadi (chunk header/ukuran total ikut berubah),
   tetapi secara teori mungkin.
4. **Arah false negative bertambah.** Berkas yang benar-benar tersimpan tetapi isinya tidak
   terbaca dan metadatanya tidak berubah akan dilaporkan `TIDAK BERUBAH`/`GAGAL`. Ini arah
   yang aman (tidak pernah mengklaim sukses), tetapi bisa membuat pengguna memeriksa manual.
5. **Overhead I/O baru**: dua kali pembacaan ±8 KiB per berkas (sebelum & sesudah `SaveAs`).
   Untuk batch ratusan berkas dampaknya kecil, tetapi tidak nol.
6. **`NotSupportedException` sekarang `Failed`.** Bila ada build CorelDRAW yang melaporkan
   Shape tunggal lewat `NotSupportedException`, copy satu objek akan menjadi **Partial**
   (bukan Verified). Tidak pernah menjadi false success — tetapi bisa mengganggu UX.
   Verifikasi: langkah manual **15.2.2** di `PANDUAN-UJI.md` §15.
7. **Glue UI belum tercakup unit test.** `CopyClickedCore` (±120 baris: dialog konfirmasi,
   urutan Activate→Paste, `DoEvents`) tetap tidak dapat diuji tanpa Windows/WinForms. Logika
   yang dipindahkan (`CopyMeasurement`, `CopyVerification`, `SaveOutputVerification`) sudah
   tercakup; sisanya harus lewat uji manual.
8. **Integration CorelDRAW 2022 = NOT RUN.** Semua perilaku COM nyata (nilai `Shape.Type`,
   bentuk berkas hasil `SaveAs` CorelDRAW 2022, perilaku lock setelah save) belum
   dibuktikan di aplikasi sungguhan. Ini risiko terbesar yang tersisa.
9. **Build Release x64 Windows = NOT RUN di lingkungan ini.** Lihat §8.

---

## 7. Perubahan pada test lama (transparansi)

Dua test lama **tidak dihapus** dan **tidak dilonggarkan**; hanya **data ujinya** dibuat
realistis, karena produksi kini memeriksa properti yang baru (bentuk berkas):

| Test | Sebelum | Sesudah | Alasan |
|---|---|---|---|
| `mode file baru: pemanggilan ok tapi file lama tetap -> TIDAK BERUBAH, GAGAL` | berkas “output lama” ditulis sebagai teks `"output lama"` | ditulis dengan `FakeDocument.FakeCdrBytes(...)` (RIFF/CDRX sah) | agar yang diuji benar-benar cabang **“file lama ada & tidak berubah”**, bukan cabang “isi bukan CDR”. Asersi tidak berubah. |
| `mode TIMPA: SaveAs tidak menulis, file lama masih ada -> GAGAL` | berkas sumber teks `"isi asli"` | berkas sumber CDR sah | sama seperti di atas; asersi tidak berubah. |

Selain itu `FakeDocument.SaveAs` kini menulis **token unik** per penyimpanan (seperti
CorelDRAW sungguhan) dan `FakeCdrBytes` menyediakan varian RIFF/CDRX, ZIP, dan “tidak valid”.
Tanpa token unik, fixture menulis isi identik sehingga benar-benar tidak ada bukti perubahan
— kasus itu sendiri sudah diuji terpisah (`Unchanged`).

---

## 8. Yang harus dijalankan sebelum merge (jangan merge otomatis)

```bash
# 1. dorong commit perbaikan ke branch PR (fast-forward, tanpa merge)
git push origin HEAD:audit/reliability-copy-verification

# 2. CI akan menjalankan: static checks -> regression tests -> msbuild Release x64 (windows-latest)
#    Lihat tab Actions pada PR #2; artifact: CorelPageCopier-Windows-x64

# 3. di mesin Windows + CorelDRAW 2022: jalankan PANDUAN-UJI.md §10, §14, §15
```

Checklist merge (semua harus ✅):

- [ ] CI PR hijau termasuk langkah `msbuild /t:Rebuild /p:Configuration=Release /p:Platform=x64`
- [ ] artifact berisi `CorelPageCopier.exe` PE32+ **x86-64**
- [ ] §15 manual: mode timpa yang gagal **tidak** merusak berkas sumber tanpa peringatan
- [ ] §15 manual: berkas `<nama>_wm.cdr` hasil CorelDRAW 2022 lolos verifikasi (bentuk ZIP)
- [ ] §15 manual: copy satu objek tetap **Verified** (bukan Partial)
- [ ] tidak ada baris `OK` pada CSV untuk berkas yang penyimpanannya gagal

**Jangan merge otomatis.** PR ini menyentuh jalur penyimpanan berkas pengguna; keputusan
merge harus setelah dua langkah di atas (CI Windows + uji manual CorelDRAW 2022).
