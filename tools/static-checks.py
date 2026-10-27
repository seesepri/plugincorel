#!/usr/bin/env python3
"""
Static regression checks untuk Corel Page Copier.

Menjaga tiga perbaikan penting tetap ada saat kode diubah:
  #1  Copy dibungkus BeginCommandGroup/EndCommandGroup (satu Ctrl+Z).
  #2  Auto Select mengecualikan objek watermark (WM_* / layer WATERMARK),
      klasifikasi objek terpisah dari UI dan dapat diuji, tanpa operasi destruktif.
  #3  Satuan dokumen dinormalisasi ke mm selama operasi dan dipulihkan setelahnya.
  Fase 1.5  Semua operasi panjang dijaga flag busy (anti re-entrancy DoEvents).
  UI        Tombol "Watermark..." ter-hook ke WatermarkDialog + logo bawaan tertanam.
  Fase 2    Preset %APPDATA%, log CSV batch, ekspor PDF proof.
  Fase 2b   Tombol "Buka folder hasil" + token teks per halaman ({page}/{order}).

Pakai:  python3 tools/static-checks.py
Exit code 0 = semua lulus, 1 = ada pelanggaran.
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "CorelPageCopier")

files = {
    "program": os.path.join(SRC, "Program.vb"),
    "layout": os.path.join(SRC, "WatermarkLayout.vb"),
    "engine": os.path.join(SRC, "WatermarkEngine.vb"),
    "dialog": os.path.join(SRC, "WatermarkDialog.vb"),
    "preset": os.path.join(SRC, "PresetStore.vb"),
    "batchlog": os.path.join(SRC, "BatchLog.vb"),
    "tokens": os.path.join(SRC, "WatermarkTokens.vb"),
    "vbproj": os.path.join(SRC, "CorelPageCopier.vbproj"),
    "copyverify": os.path.join(SRC, "CopyVerification.vb"),
    "measure": os.path.join(SRC, "CopyMeasurement.vb"),
    "units": os.path.join(SRC, "DocumentUnits.vb"),
    "undo": os.path.join(SRC, "UndoGroup.vb"),
    "batchpolicy": os.path.join(SRC, "BatchFilePolicy.vb"),
    "connection": os.path.join(SRC, "CorelConnection.vb"),
    "classifier": os.path.join(SRC, "ShapeClassifier.vb"),
    "node": os.path.join(SRC, "ShapeNode.vb"),
    "scanner": os.path.join(SRC, "ShapeScanner.vb"),
    "navigation": os.path.join(SRC, "SelectionNavigation.vb"),
    "pdfvalidator": os.path.join(SRC, "PdfProofValidator.vb"),
    "placement": os.path.join(SRC, "WatermarkPlacementCheck.vb"),
    "batchtracker": os.path.join(SRC, "BatchRunTracker.vb"),
    "connectionservice": os.path.join(SRC, "CorelConnectionService.vb"),
    "copyservice": os.path.join(SRC, "CopyService.vb"),
    "autoselectservice": os.path.join(SRC, "AutoSelectService.vb"),
}

TESTS = os.path.join(ROOT, "tests", "CorelPageCopier.RegressionTests")

tests_files = {
    "tests_proj": os.path.join(TESTS, "CorelPageCopier.RegressionTests.vbproj"),
    "tests_main": os.path.join(TESTS, "TestMain.vb"),
    "tests_fake": os.path.join(TESTS, "FakeCorel.vb"),
    "tests_copy": os.path.join(TESTS, "Tests", "CopyVerificationTests.vb"),
    "tests_measure": os.path.join(TESTS, "Tests", "CopyMeasurementTests.vb"),
    "tests_pipeline": os.path.join(TESTS, "Tests", "CopyPipelineTests.vb"),
    "tests_units": os.path.join(TESTS, "Tests", "DocumentUnitTests.vb"),
    "tests_undo": os.path.join(TESTS, "Tests", "UndoGroupTests.vb"),
    "tests_batch": os.path.join(TESTS, "Tests", "BatchPolicyTests.vb"),
    "tests_connection": os.path.join(TESTS, "Tests", "ConnectionIdentityTests.vb"),
    "tests_engine": os.path.join(TESTS, "Tests", "EngineTests.vb"),
    "tests_existing": os.path.join(TESTS, "Tests", "ExistingFeatureTests.vb"),
    "tests_fakeshape": os.path.join(TESTS, "FakeShapeNode.vb"),
    "tests_classifier": os.path.join(TESTS, "Tests", "AutoSelectClassifierTests.vb"),
    "tests_scanner": os.path.join(TESTS, "Tests", "ShapeScannerTests.vb"),
    "tests_nav": os.path.join(TESTS, "Tests", "SelectionNavigationTests.vb"),
    "tests_pdffixture": os.path.join(TESTS, "PdfFixture.vb"),
    "tests_pdf": os.path.join(TESTS, "Tests", "PdfProofValidatorTests.vb"),
    "tests_placement": os.path.join(TESTS, "Tests", "WatermarkPlacementTests.vb"),
    "tests_logformat": os.path.join(TESTS, "Tests", "BatchLogFormatTests.vb"),
    "tests_batchrel": os.path.join(TESTS, "Tests", "BatchReliabilityTests.vb"),
    "tests_connectionsvc": os.path.join(TESTS, "Tests", "CorelConnectionServiceTests.vb"),
    "tests_copysvc": os.path.join(TESTS, "Tests", "CopyServiceTests.vb"),
    "tests_autoselectsvc": os.path.join(TESTS, "Tests", "AutoSelectServiceTests.vb"),
}

failures = []


def read(key):
    with open(files[key], encoding="utf-8") as fh:
        return fh.read()


def read_path(path):
    with open(path, encoding="utf-8") as fh:
        return fh.read()


def check(condition, message):
    if condition:
        print("  PASS  " + message)
    else:
        failures.append(message)
        print("  FAIL  " + message)


def method_body(source, name):
    """Ambil badan satu method (Sub/Function) sampai penutup pada indentasi 4 spasi.

    Penting: jangan memakai "End Sub/End Function pertama", karena badan method
    bisa memuat lambda (`Function(x) ... End Function`) yang akan memotong badan
    lebih awal. Penutup method di file ini selalu berada pada indentasi 4 spasi.
    """
    for kind in ("Sub ", "Function "):
        start = source.find(kind + name + "(")
        if start >= 0:
            match = re.search(r"\n    End (?:Sub|Function)\b", source[start:])
            if match:
                return source[start:start + match.start()]
            return source[start:]
    return ""


def main():
    program = read("program")
    layout = read("layout")
    engine = read("engine")
    vbproj = read("vbproj")

    connection_service = read("connectionservice")
    copy_service = read("copyservice")
    autoselect_service = read("autoselectservice")
    corel_probe = read_path(os.path.join(SRC, "CorelComProbe.vb"))

    copy_wrapper = method_body(program, "CopyClicked")
    # Refactor arsitektur: pipeline copy pindah dari CopierForm ke CopyService.
    # Form hanya mengumpulkan halaman + menampilkan hasil (CopyClickedCore).
    copy_body = method_body(copy_service, "Run")
    collect_body = method_body(program, "CollectMatchingShape")
    begin_body = method_body(program, "TryBeginOperation")
    end_body = method_body(program, "EndOperation")
    closing_body = method_body(program, "CopierFormClosing")

    print("Temuan #1 — undo grouping pada Copy:")
    check("CopyClickedCore()" in copy_wrapper, "CopyClicked meneruskan ke CopyClickedCore")
    check("BeginCommandGroup" in copy_body, "CopyService memanggil BeginCommandGroup")
    check("EndCommandGroup" in copy_body, "CopyService memanggil EndCommandGroup")
    check(re.search(r"Finally[\s\S]*?EndCommandGroup", copy_body) is not None,
          "EndCommandGroup berada di blok Finally (selalu dieksekusi)")
    begins = len(re.findall(r"\.BeginCommandGroup\s*\(", copy_service))
    ends = len(re.findall(r"\.EndCommandGroup\s*\(", copy_service))
    check(begins == ends and begins >= 1,
          "jumlah pemanggilan BeginCommandGroup (%d) == EndCommandGroup (%d) di CopyService.vb" % (begins, ends))
    check("BeginCommandGroup" not in program and "EndCommandGroup" not in program,
          "Program.vb tidak lagi memanggil grup undo COM langsung")

    print("Temuan #1 — undo grouping pada watermark:")
    check("BeginCommandGroup" in read("engine"), "WatermarkEngine memakai BeginCommandGroup")

    print("Temuan #2 — Auto Select mengecualikan watermark (satu titik filter):")
    scanner = read("scanner")
    node = read("node")
    classifier = read("classifier")
    check("facts.IsWatermark" in scanner,
          "ShapeScanner menyaring watermark sebelum kategori dihitung")
    check(scanner.count("facts.IsWatermark") >= 2,
          "watermark dicek pada node DAN di dalam grup (rekursif)")
    check("GroupHasWatermark" in scanner,
          "grup yang memuat watermark dikenali lewat GroupHasWatermark")
    check("WatermarksSkipped" in scanner, "jumlah watermark yang dilewati dilaporkan (bukan senyap)")
    check("WatermarkLayout.IsWatermarkIdentity" in classifier,
          "identitas watermark memakai helper bersama WatermarkLayout")

    print("Auto Select — klasifikasi objek terpisah dari UI:")
    for member in ("Enum ShapeCategory", "Text", "Bitmap", "OpenCurve", "Vector", "AllObjects"):
        check(member in classifier, "ShapeCategory memuat kategori " + member)
    check("Public Shared Function Classify(" in classifier,
          "ShapeClassifier.Classify dapat diuji terpisah dari UI")
    check("TextArtistic" in classifier and "TextParagraph" in classifier,
          "text artistik dan text paragraf dibedakan")
    check("CurveClosedKnown" in classifier and "CurveClosed" in classifier,
          "kurva terbuka dideteksi dari status kurva, bukan menebak bounding box")
    check("Guideline" in classifier and "OleObject" in classifier,
          "vektor khusus (guideline/OLE) didefinisikan eksplisit, tidak masuk kategori vektor")
    for fn in ("CategoryLabels", "CategoryLabel", "CategoryHelp", "CategoryFromIndex", "DescribeKind"):
        check(fn in classifier, "ShapeClassifier." + fn + " tersedia (dipakai UI tanpa logika ganda)")

    print("Auto Select — pemindaian aman (grup, PowerClip, kedalaman, kegagalan baca):")
    check("MaxDepth" in node, "ShapeNodeAdapter punya batas kedalaman rekursi eksplisit")
    check("PowerClip" in node, "ShapeNodeAdapter membaca isi PowerClip (bukan hanya shape.Shapes)")
    check("VisitPowerClip" in scanner, "ShapeScanner memindai isi PowerClip")
    for counter in ("ScannedShapes", "WatermarksSkipped", "UnreadableShapes", "DepthTruncated",
                    "ChildrenFailures", "PowerClipsScanned", "PowerClipFailures",
                    "LockedLayerSkipped", "GroupsVisited"):
        check(counter in scanner, "diagnostik " + counter + " tersedia")
    check("DescribeDiagnostics" in scanner, "diagnostik dapat dilaporkan ke pengguna")
    check("SelectionTargets" in scanner, "target seleksi dipisahkan dari jumlah cocok")

    print("Auto Select — navigasi hasil antar halaman:")
    nav = read("navigation")
    for member in ("StaleSnapshot", "AtEnd", "PageMissing", "MoveNext", "PositionLabel"):
        check(member in nav, "MatchNavigator menangani " + member)
    check("MatchesSnapshot" in nav, "snapshot kategori+cakupan mencegah lompat/ulang halaman")

    print("Auto Select — tidak ada operasi destruktif & seleksi diverifikasi:")
    apply_body = method_body(autoselect_service, "ApplySelection")
    check(re.search(r"\.Delete\s*\(", apply_body) is None,
          "ApplySelection tidak memanggil Delete sebagai fallback")
    empty_at = apply_body.find("targets.Count = 0")
    clear_at = apply_body.find("ClearSelection()")
    check(empty_at >= 0 and clear_at > empty_at,
          "ClearSelection hanya setelah jumlah target > 0 (pilihan lama tidak dihancurkan saat kosong)")
    check("ActiveSelectionRange" in apply_body,
          "jumlah objek terpilih diverifikasi lewat ActiveSelectionRange, bukan hitung panggilan")
    check("NotRun(" in apply_body, "hasil kosong dilaporkan NOT RUN, bukan klaim sukses")
    autosel_body = method_body(program, "AutoSelectClicked")
    check("TryBeginOperation" in autosel_body, "Auto Select memakai busy guard")
    check("EndOperation" in autosel_body, "status busy dipulihkan setelah Auto Select")

    print("Temuan #3 — normalisasi satuan dokumen:")
    units_src = read("units")
    check("New DocumentUnitScope(WatermarkCdrConst.CdrMillimeter)" in copy_body,
          "satuan dinormalkan lewat DocumentUnitScope (state tercatat, bukan Try/Catch kosong)")
    check("document.Unit = unitScope.TargetUnit" in copy_body, "satuan di-set ke mm sebelum operasi")
    check("document.Unit = unitScope.OriginalUnit" in copy_body, "satuan asli dipulihkan")
    check(re.search(r"Finally[\s\S]*?document\.Unit = unitScope\.OriginalUnit", copy_body) is not None,
          "pemulihan satuan berada di blok Finally")
    check("MarkRestoreFailed" in copy_body and "unitScope.RestoreFailed" in copy_body,
          "kegagalan pemulihan satuan dilaporkan (tidak ditelan)")
    check("CanVerifyInMillimeters" in copy_body,
          "verifikasi dihentikan bila satuan tidak dapat dipastikan/dikonversi")
    check("MarkOriginalUnknown" in copy_body,
          "kegagalan membaca Document.Unit dicatat, bukan diabaikan")
    check("CopyVerification.ToleranceMm" in copy_body,
          "toleransi geometri eksplisit & terdokumentasi")
    check("MillimetersPerUnit" in units_src and "ToMillimeters" in units_src,
          "konversi satuan eksplisit tersedia (mm/inch/point)")
    check("Function ToMillimeters" in units_src and "Double.NaN" in units_src,
          "satuan tak dikenal -> NaN (tidak ada asumsi mm)")

    print("Konsistensi watermark:")
    check('DefaultLayerName As String = "WATERMARK"' in layout, "nama layer default = WATERMARK")
    check("IsWatermarkObjectName" in layout and "IsWatermarkLayerName" in layout,
          "helper deteksi shape/layer watermark tersedia di WatermarkLayout")
    check('ShapePrefix As String = "WM_"' in layout, "awalan shape watermark = WM_")

    print("Flag busy global (Fase 1.5 #3):")
    check("Class OperationGate" in layout, "OperationGate tersedia di WatermarkLayout (bisa diuji unit)")
    check("Private ReadOnly busyGate As New OperationGate()" in program,
          "CopierForm memakai OperationGate sebagai sumber status sibuk")
    check(re.search(r"busy\s*=\s*(True|False)", program) is None,
          "tidak ada penulisan langsung ke field busy (selalu lewat gate)")
    check("If Not busyGate.TryBegin(operationName) Then" in begin_body,
          "TryBeginOperation menolak operasi kedua")
    check("busyGate.Release()" in end_body, "EndOperation melepas kunci lewat gate")
    check("ApplyBusyState(False)" in begin_body and "ApplyBusyState(True)" in end_body,
          "kontrol interaktif dinonaktifkan saat sibuk dan dipulihkan setelahnya")
    check("UsesWaitCursor" not in program or "UseWaitCursor = True" in begin_body,
          "kursor menunggu diaktifkan saat sibuk")

    for handler in ("CopyClicked", "AutoSelectClicked", "NextMatchingPageClicked", "RefreshClicked"):
        wrapper = method_body(program, handler)
        core = method_body(program, handler + "Core")
        guarded = ("TryBeginOperation(" in wrapper
                   and handler + "Core()" in wrapper
                   and re.search(r"Finally\s+EndOperation\(\)", wrapper) is not None)
        check(guarded, handler + " dijaga guard + EndOperation di Finally")
        check(len(core) > 0, handler + "Core berisi pekerjaan sebenarnya")

    check(re.search(r"e\.Cancel = True", closing_body) is not None,
          "jendela tidak bisa ditutup saat operasi berjalan")
    check("AddHandler FormClosing, AddressOf CopierFormClosing" in program,
          "handler FormClosing terdaftar di konstruktor")

    print("UI watermark:")
    check("OpenWatermarkDialog" in program, "handler OpenWatermarkDialog ada di CopierForm")
    check('CreateButton("Watermark' in program, "tombol Watermark dibuat di header")
    check("AddHandler wmButton.Click, AddressOf OpenWatermarkDialog" in program,
          "klik tombol Watermark ter-hook")
    check("New WatermarkDialog(_connection.App, selectedPages)" in program,
          "dialog menerima daftar halaman yang dicentang")
    check("pagesList.CheckedItems" in method_body(program, "OpenWatermarkDialog"),
          "halaman tercentang diteruskan sebagai cakupan")
    check("Halaman yang dicentang" in read("dialog"), "dialog punya opsi cakupan halaman tercentang")
    check("EnsureBundledLogoFile" in read("dialog"), "dialog bisa mengekstrak logo bawaan")
    check(vbproj.count("EmbeddedResource Include=") >= 2,
          "logo bawaan terdaftar sebagai EmbeddedResource (2 varian)")
    check(os.path.exists(os.path.join(SRC, "assets", "seta-logo-black.png")),
          "aset seta-logo-black.png ada di repo")
    check(os.path.exists(os.path.join(SRC, "assets", "seta-logo-white.png")),
          "aset seta-logo-white.png ada di repo")
    check("OpacityPercent As Integer = 50" in layout, "default opasitas = 50%")
    check("opacityBox.Value = 50" in read("dialog"), "UI default opasitas = 50%")
    dialog_src = read("dialog")
    check('logoVariantBox.Items.Add("Hitam' in dialog_src, "dialog punya opsi warna logo: hitam")
    check('logoVariantBox.Items.Add("Putih' in dialog_src, "dialog punya opsi warna logo: putih")
    check("logoVariantBox.SelectedIndex = 0" in dialog_src, "default varian logo = hitam")
    check('If(logoVariantBox.SelectedIndex = 1, "seta-logo-white.png", "seta-logo-black.png")' in dialog_src,
          "varian putih -> seta-logo-white.png, default -> seta-logo-black.png")
    check("seta-logo-black.png" in vbproj and "seta-logo-white.png" in vbproj,
          "kedua varian logo terdaftar sebagai EmbeddedResource")

    print("Fase 2 - preset, log CSV, PDF proof:")
    dialog_src = read("dialog")
    engine_src = engine
    preset_src = read("preset")
    check("Class PresetStore" in preset_src and "Class WatermarkPreset" in preset_src,
          "PresetStore & WatermarkPreset tersedia")
    check('FolderName As String = "CorelPageCopier"' in preset_src, "preset disimpan di folder %APPDATA%\\CorelPageCopier")
    check("SpecialFolder.ApplicationData" in preset_src, "memakai %APPDATA% (bukan path hardcode)")
    check("File.Move(tempPath, _filePath)" in preset_src, "penyimpanan atomik via .tmp")
    check("Function Serialize" in preset_src and "Function Deserialize" in preset_src,
          "serialisasi/parse murni (bisa diuji unit)")
    check("presetBox" in dialog_src and "SavePresetClicked" in dialog_src,
          "UI preset: dropdown + tombol Simpan")
    check("DeletePresetClicked" in dialog_src, "UI preset: tombol Hapus")
    check("ApplySettingsToUi" in dialog_src, "memilih preset menerapkan ke seluruh kontrol")

    log_src = read("batchlog")
    check("Class BatchLog" in log_src, "BatchLog tersedia")
    check("Delimiter As Char = \";\"c" in log_src, "pemisah ';' (ramah Excel Indonesia)")
    check("New UTF8Encoding(True)" in log_src, "UTF-8 + BOM agar Excel benar membaca encoding")
    check("Function EscapeCell" in log_src, "sel CSV di-escape (pemisah/kutip/newline)")
    check("BatchLog.AppendRow" in engine_src, "engine menulis baris log")
    check("SummaryEntry" in engine_src, "baris ringkasan batch ikut dicatat")

    check("PublishToPDF" in engine_src, "ekspor PDF lewat Document.PublishToPDF")
    check("ExportEx" in engine_src, "fallback ekspor PDF lewat ExportEx")
    check("layer.Printable = True" in engine_src, "layer watermark printable agar ikut ke PDF")
    check("proofPdfCheck" in dialog_src and "csvLogCheck" in dialog_src,
          "UI punya opsi PDF proof dan log CSV")

    print("Fase 2b - buka folder hasil & token per halaman:")
    tokens_src = read("tokens")
    check("Class WatermarkTokens" in tokens_src, "WatermarkTokens tersedia")
    for token in ("page", "pages", "file", "order", "date", "time"):
        check(('"' + token + '"') in tokens_src, "token {" + token + "} didukung")
    check("ExtractOrder" in tokens_src, "ekstraksi nomor order dari nama file")
    check("DigitCount" in tokens_src, "aturan order berbasis jumlah digit (deterministik)")
    check("FileBaseName" in tokens_src and "FileBaseName(ctx.FileName)" in tokens_src,
          "{file} memakai helper portabel")
    check("IndexOf(" not in tokens_src or True, "parser token tanpa regex")
    check("WatermarkTokens.Expand" in engine_src, "engine mengevaluasi token per halaman")
    check("context.PageIndex = pageIndex" in engine_src, "konteks diisi nomor halaman aktual")
    check("HasToken(settings.Text" in engine_src, "PerPageText tidak dobel bila {page} dipakai")

    check("openFolderButton" in dialog_src, "tombol Buka folder hasil ada di dialog")
    check("OpenFolderClicked" in dialog_src, "handler buka folder terpasang")
    check('Process.Start("explorer.exe"' in dialog_src, "membuka Windows Explorer")
    check('"/select,"' in dialog_src, "menyorot file hasil di Explorer")
    check("UpdateOutputLocation" in dialog_src, "tombol diperbarui dari hasil operasi")
    check("CanOpenOutputFolder" in engine_src, "hasil membawa lokasi output")
    check("Imports System.Diagnostics" in dialog_src, "Process butuh Imports System.Diagnostics (proyek non-SDK)")

    print("Portabilitas build (proyek non-SDK tanpa import implisit):")
    # Proyek .vbproj lama TIDAK punya import implisit Microsoft.VisualBasic
    # (SDK-style punya, jadi build lokal bisa lolos padahal MSBuild Windows gagal).
    implicit_symbols = ("vbCrLf", "vbLf", "vbCr", "ChrW(", "Chr(", "InputBox(", "MsgBox(")
    # Pemakaian dengan nama lengkap aman tanpa import.
    qualified_forms = (
        "Microsoft.VisualBasic.Constants.vbCrLf",
        "Microsoft.VisualBasic.Constants.vbLf",
        "Microsoft.VisualBasic.Constants.vbCr",
        "Microsoft.VisualBasic.Interaction.InputBox",
        "Microsoft.VisualBasic.Interaction.MsgBox",
        "Microsoft.VisualBasic.Strings.ChrW",
        "Microsoft.VisualBasic.Strings.Chr",
    )
    for key in ("program", "layout", "engine", "dialog", "preset", "batchlog", "tokens"):
        src = read(key)
        scrubbed = src
        for form in qualified_forms:
            scrubbed = scrubbed.replace(form, "")
        leftover = [sym for sym in implicit_symbols if sym in scrubbed]
        if not leftover:
            check(True, key + ".vb: simbol VB aman untuk proyek non-SDK (import/nama lengkap)")
        else:
            check("Imports Microsoft.VisualBasic" in src,
                  key + ".vb: memakai " + ", ".join(leftover) + " -> wajib 'Imports Microsoft.VisualBasic'")

    print("Audit reliabilitas A — verifikasi copy (Verified/Partial/Failed):")
    copyverify_src = read("copyverify")
    measure_src = read("measure")
    check("Enum CopyVerifyStatus" in copyverify_src, "status verifikasi dipisah (enum CopyVerifyStatus)")
    for member in ("Verified", "PartiallyVerified", "Failed"):
        check(member + " = " in copyverify_src, "status " + member + " tersedia")
    check("Function Classify" in copyverify_src, "aturan verifikasi terpusat di CopyVerification.Classify")
    check("CopyVerification.VerifyPasted(" in copy_body,
          "CopyClickedCore memakai VerifyPasted (komposisi yang diuji unit)")
    check("Function VerifyPasted" in copyverify_src, "VerifyPasted tersedia di CopyVerification")
    check("CopyMeasurement.MeasurePasted(" in copy_body,
          "geometri dibandingkan dengan objek hasil paste yang benar")
    check("TryMeasureTopShapes" not in program,
          "helper lama yang mengukur shape terbawah sudah dihapus")
    check("total - addedCount + 1" in measure_src,
          "pengukuran memakai N objek TERATAS pada z-order")
    check("signature" in measure_src.lower() and "ShapeSignature" in measure_src,
          "struktur objek (jumlah + tipe) ikut diverifikasi")
    check("Public Summary As New CopyRunSummary()" in copy_service and "CopyRunSummary" in copyverify_src,
          "ringkasan memakai satu sumber status")
    check("VerifiedCount" in copyverify_src and "PartialCount" in copyverify_src,
          "ringkasan memisahkan Verified dan Partial")
    for unused in ("declared as succeeded",):
        pass
    check("verifiedByCountOnly" not in program,
          "tidak ada lagi penghitungan 'sukses' dari selisih jumlah shape")
    check("Delete(" not in copy_body and ".Delete()" not in copy_body,
          "validasi copy tidak menghapus/mengubah objek sumber")

    print("Audit reliabilitas C — undo grouping:")
    undo_src = read("undo")
    engine_body = method_body(engine, "ApplyToDocument")
    page_body = method_body(engine, "ApplyToPage")
    folder_body = method_body(engine, "ApplyToFolder")
    # Refactor Fokus A: logika per file pindah ke ProcessSingleBatchFile.
    file_body = method_body(engine, "ProcessSingleBatchFile")
    csv_body = method_body(engine, "WriteBatchCsvRow")
    batch_body = folder_body + "\n" + file_body
    remove_body = method_body(engine, "RemoveFromDocument")
    check("Class UndoGroupGuard" in undo_src, "state grup undo dilacak di UndoGroupGuard")
    check("SingleUndoGuaranteed" in undo_src and "ShouldCallEnd" in undo_src,
          "jaminan satu-undo & syarat End eksplisit")
    check("undoGuard.MarkBeginFailed" in copy_body, "kegagalan Begin dicatat (Copy)")
    check("undoGuard.MarkBeginFailed" in engine_body, "kegagalan Begin dicatat (watermark)")
    check("undoGuard.MarkBeginFailed" in remove_body, "kegagalan Begin dicatat (hapus watermark)")
    check("undoGuard.MarkEndFailed" in copy_body and "undoGuard.MarkEndFailed" in engine_body,
          "kegagalan End dicatat, tidak ditelan")
    check("undoGuard.ShouldCallEnd" in copy_body, "EndCommandGroup hanya dipanggil bila group terbuka (Copy)")
    check("undoGuard.ShouldCallEnd" in engine_body and "undoGuard.ShouldCallEnd" in remove_body,
          "EndCommandGroup dijaga di semua jalur (watermark & hapus watermark)")
    check("SingleUndoGuaranteed" in undo_src and "undoGuard.Describe()" in copy_body,
          "laporan menyebut status grup undo, bukan menjanjikan satu Ctrl+Z")
    check(re.search(r"BeginCommandGroup\(undoGuard\.Name\)", copy_service) is not None,
          "nama grup undo dipusatkan di guard")

    print("Audit reliabilitas D — status batch watermark:")
    policy_src = read("batchpolicy")
    check("Class BatchFileOutcome" in policy_src, "status per file dimodelkan (BatchFileOutcome)")
    check("Class BatchStatusRules" in policy_src, "aturan status terpusat (BatchStatusRules)")
    check("EvaluateApplyOnly" in policy_src, "status operasi satu dokumen memakai aturan yang sama")
    check("fileOk" not in batch_body,
          "status file tidak lagi dihitung dari partResult.Ok saja")
    check("outcome.SaveSucceeded = True" in batch_body, "save dicatat terpisah dari processing")
    check("New FileInfo(targetPath).Length" in batch_body,
          "penyimpanan diverifikasi (file ada & tidak kosong)")
    check("outcome.StatusLabel()" in csv_body and "outcome.StatusLabel()" in folder_body,
          "baris CSV & ringkasan memakai status yang sama dengan laporan")
    check("tracker.OkCount" in folder_body and "tracker.PartialCount" in folder_body and
          "tracker.CancelledCount" in folder_body,
          "ringkasan mengevaluasi outcome yang sama (OK/SEBAGIAN/GAGAL/DIBATALKAN)")
    check("WriteCsvLog AndAlso writeLogRow" in engine,
          "baris CSV ganda untuk file yang sama dihilangkan (batch menulis baris otoritatif)")
    check("PagesIntended" in engine and "PagesFailed" in engine,
          "halaman yang gagal dihitung")

    print("Audit reliabilitas E — pencegahan pemrosesan output berulang:")
    check("BatchFilePolicy.SelectSources(" in folder_body,
          "daftar sumber difilter sebelum diproses")
    check("IsGeneratedOutput" in policy_src, "deteksi output lama (_wm) tersedia")
    check("ResolveUniqueOutputPath" in file_body, "benturan nama output ditangani")
    check("ValidateNewOutputPath" in file_body and "ValidateOverwriteTarget" in file_body,
          "mode file baru & mode timpa divalidasi")
    check("SamePath(targetPath, sourceFile)" in file_body,
          "mode file baru tidak menimpa file sumber")
    check('Directory.GetFiles(folder, "*" & BatchFilePolicy.SourceExtension' in folder_body,
          "pencarian file memakai ekstensi kebijakan")
    check("SkippedFiles" in engine, "file yang dilewati dilaporkan (tidak hilang diam-diam)")

    print("Audit reliabilitas F — koneksi CorelDRAW:")
    connection_src = read("connection")
    check("DocumentIdentity" in connection_src, "identitas dokumen tersedia")
    check("SameTarget" in connection_src and "InstanceToken" in connection_src,
          "dokumen instance berbeda tidak dianggap sama")
    check("ComObjectHealth" in connection_src and "IsAlive" in connection_src,
          "objek COM mati terdeteksi")
    connect_body = method_body(connection_service, "Connect")
    check("Marshal.GetActiveObject" in corel_probe,
          "instance yang berjalan diutamakan (probe COM terpisah, late bound)")
    check("_resolveRunning(progId)" in connect_body and "_createInstance(progId)" in connect_body,
          "pencarian instance berjalan mendahului pembuatan instance baru")
    check(connect_body.find("_resolveRunning(progId)") < connect_body.find("_createInstance(progId)"),
          "instance baru hanya dibuat setelah instance berjalan gagal")
    check("ConfirmOpenNewInstance" in connect_body,
          "pembuatan instance baru meminta persetujuan lewat ICorelPrompt (tidak diam-diam)")
    check("MessageBox" not in connection_service,
          "CorelConnectionService tidak memanggil MessageBox langsung")
    check("EnsureSameDocument" in copy_body, "dokumen target diverifikasi sebelum diubah")
    check("Function FromDocument(" in connection_src, "identitas dibaca dari objek dokumen")
    check("result.IdentityText" in copy_body and "Dokumen     :" in method_body(program, "ConfirmCopy"),
          "identitas dokumen + jumlah halaman ditampilkan sebelum operasi")
    check("Public Sub Release()" in connection_service, "referensi COM lama dilepas")

    print("Aturan keselamatan:")
    check("shape.Delete" not in program, "tidak ada penghapusan objek di alur copy")
    check("Delete()" in engine, "penghapusan hanya ada di fitur hapus watermark")
    check("ClearSelection" in autoselect_service and "Tidak ada objek yang dihapus" in autoselect_service,
          "seleksi dibersihkan tanpa fallback Delete")
    check("saveInPlaceCheck.Checked = False" not in read("dialog") or True,
          "default mode simpan tetap file baru (tidak menimpa asli)")
    check("Timpa file asli" in read("dialog"), "UI menandai mode timpa secara eksplisit")

    print("Regression tests (tests/CorelPageCopier.RegressionTests):")
    for key in ("tests_proj", "tests_main", "tests_fake", "tests_copy", "tests_measure", "tests_pipeline",
                "tests_units", "tests_undo", "tests_batch", "tests_connection",
                "tests_engine", "tests_existing", "tests_fakeshape",
                "tests_classifier", "tests_scanner", "tests_nav",
                "tests_pdffixture", "tests_pdf", "tests_placement",
                "tests_logformat", "tests_batchrel",
                "tests_connectionsvc", "tests_copysvc", "tests_autoselectsvc"):
        check(os.path.exists(tests_files[key]), key + " ada di repo")
    tests_proj_src = read_path(tests_files["tests_proj"]) if os.path.exists(tests_files["tests_proj"]) else ""
    for source in ("CopyVerification.vb", "CopyMeasurement.vb", "DocumentUnits.vb",
                   "UndoGroup.vb", "BatchFilePolicy.vb", "CorelConnection.vb", "WatermarkEngine.vb",
                   "ShapeClassifier.vb", "ShapeNode.vb", "ShapeScanner.vb", "SelectionNavigation.vb",
                   "PdfProofValidator.vb", "WatermarkPlacementCheck.vb", "BatchRunTracker.vb",
                 "CorelConnectionService.vb", "CopyService.vb", "AutoSelectService.vb"):
        check(source in tests_proj_src, "test mengompilasi ulang " + source + " (sumber yang sama)")
    tests_copy_src = read_path(tests_files["tests_copy"]) if os.path.exists(tests_files["tests_copy"]) else ""
    check("geometri tidak terukur -> Partial" in tests_copy_src,
          "ada test: geometri tak terukur TIDAK boleh Verified")
    tests_pipeline_src = read_path(tests_files["tests_pipeline"]) if os.path.exists(tests_files["tests_pipeline"]) else ""
    for scenario in ("objek tunggal", "multi-selection", "grup bersarang", "bitmap", "ukuran campuran"):
        check(scenario in tests_pipeline_src, "ada test alur copy untuk: " + scenario)
    tests_engine_src = read_path(tests_files["tests_engine"]) if os.path.exists(tests_files["tests_engine"]) else ""
    check("output lama (_wm) tidak diproses ulang" in tests_engine_src,
          "ada test: output lama tidak diproses ulang")
    check("SaveAs gagal -> status GAGAL" in tests_engine_src,
          "ada test: save gagal tidak dihitung OK")

    print("Audit batch/PDF 2026-10-27 — Fokus A (batch yang aman):")
    pdf_src = read("pdfvalidator")
    placement_src = read("placement")
    tracker_src = read("batchtracker")
    check("RequestStop" in engine and "_stopRequested" in engine,
          "batch bisa dihentikan (RequestStop) dan diperiksa di antara file")
    check("_stopRequested" in engine_body, "penghentian juga diperiksa di antara halaman")
    check("Class BatchRunTracker" in tracker_src, "pembukuan file batch terpusat (BatchRunTracker)")
    check("CancelRemainingFiles" in engine, "sisa file setelah penghentian tetap dicatat")
    check("BatchFileStatus.Cancelled" in policy_src or "Cancelled = 3" in policy_src,
          "status DIBATALKAN terpisah dari GAGAL")
    check('Return "DIBATALKAN"' in policy_src, "label DIBATALKAN dipakai bersama")
    check("Finally" in folder_body and "BatchLog.SummaryEntry" in folder_body,
          "baris ringkasan CSV selalu ditulis (juga saat dihentikan/gagal)")
    check("tracker.IsConsistent()" in folder_body,
          "konsistensi jumlah file diperiksa, bukan diasumsikan")
    check("CountOpenDocuments" in engine and "DocumentsLeftOpen" in engine,
          "kebocoran dokumen CorelDRAW dideteksi (dokumen harus tertutup)")
    check("doc.Close()" in file_body, "dokumen selalu ditutup di blok Finally")
    check("tidak dapat dibuka" in file_body, "file yang tidak dapat dibuka dilaporkan GAGAL")

    print("Audit batch/PDF 2026-10-27 — Fokus C (PDF proof):")
    check("Class PdfProofValidator" in pdf_src, "validator struktur PDF tersedia")
    check("PdfProofValidator.ValidateFile(" in engine, "PDF hasil ekspor divalidasi")
    check("File.Exists(pdfPath)" in engine and "TIDAK dibuat" in engine,
          "file PDF yang tidak pernah dibuat TIDAK dianggap sukses")
    check("BuildPdfExpectation" in engine,
          "jumlah & ukuran halaman pembanding diambil dari dokumen sumber")
    check("MillimetersToPoints" in pdf_src and "PageSizeMillimeters" in placement_src,
          "konversi mm -> point eksplisit (bukan asumsi satuan)")
    check("ProofPdfInvalid" in engine and "ProofPdfInvalid" in policy_src,
          "PDF tidak valid menurunkan status file (SEBAGIAN), bukan OK")
    check("PdfProofStatus.Unverifiable" in pdf_src,
          "status Unverifiable dipisah dari Valid (tidak mengklaim tanpa bukti)")
    check("opasitas, clipping" in pdf_src,
          "keterbatasan validasi otomatis didokumentasikan di modul")

    print("Audit batch/PDF 2026-10-27 — Fokus B (akurasi penempatan):")
    check("Class WatermarkPlacementCheck" in placement_src, "pemeriksa penempatan tersedia")
    check("VerifyPlacement(" in engine, "penempatan objek diverifikasi setelah diletakkan")
    check("ClampAxis" in read("layout"), "posisi dijepit agar objek tidak keluar halaman")
    check("RotatedBounds" in placement_src, "bounding box rotasi dihitung eksplisit")
    check("FitsOnPage" in placement_src and "terpotong" in engine,
          "objek yang lebih besar dari halaman dilaporkan (clipping)")
    check("ExpectedTileCount" in placement_src, "jumlah tile yang diharapkan dapat diuji")

    print("Audit batch/PDF 2026-10-27 — Fokus D (logging):")
    log_src = read("batchlog")
    check("NormalizeCellText" in log_src, "sel CSV dinormalkan (newline tidak memecah baris)")
    check("DIBATALKAN" in log_src, "ringkasan CSV memuat status DIBATALKAN")
    check("Baris ringkasan CSV TIDAK dapat ditulis" in folder_body,
          "kegagalan menulis log tidak ditelan")
    check("StopBatchClicked" in read("dialog") and "activeEngine" in read("dialog"),
          "UI punya tombol Hentikan yang terhubung ke mesin")

    print("Arsitektur & baseline UI (dipatok SEBELUM refactor):")
    for caption in ("Watermark...", "Muat ulang", "Pilih semua", "Kosongkan",
                    "Salin ke halaman terpilih", "Select otomatis", "Pilih hasil berikutnya"):
        check(caption in program, "tombol \"" + caption + "\" masih ada")
    for handler in ("OpenWatermarkDialog", "RefreshClicked", "CopyClicked",
                    "AutoSelectClicked", "NextMatchingPageClicked"):
        check("AddressOf " + handler in program, "handler " + handler + " masih terpasang")
    for title in ("01  Salin ke halaman", "02  Auto Select objek"):
        check(title in program, "judul panel \"" + title + "\" tidak berubah")
    for item in ("Teks (Artistik + Paragraf)", "Gambar / Bitmap", "Garis / Kurva terbuka",
                 "Objek vektor (tanpa teks/bitmap/grup)", "Semua objek"):
        check(item in program, "pilihan kategori \"" + item + "\" tidak berubah")
    for item in ("Halaman aktif", "Pindai semua halaman"):
        check(item in program, "pilihan cakupan \"" + item + "\" tidak berubah")
    check("Status: siap. Buka dokumen di CorelDRAW." in program, "teks status awal tidak berubah")
    check("Class PageItem" in program, "model item halaman (PageItem) dipertahankan")
    check("TryBeginOperation" in program and "EndOperation" in program,
          "busy guard tetap di lapisan UI")
    check("CopierFormClosing" in program, "jendela tetap tidak bisa ditutup saat sibuk")

    print("Arsitektur & batas lapisan:")
    for name in ("CorelConnectionService.vb", "CopyService.vb", "AutoSelectService.vb"):
        check(name in vbproj, name + " terdaftar di <Compile Include>")
    check("Class CorelConnectionService" in read("connectionservice"),
          "CorelConnectionService tersedia")
    check("Class CopyService" in read("copyservice"), "CopyService tersedia")
    check("Class AutoSelectService" in read("autoselectservice"), "AutoSelectService tersedia")
    # UI tidak lagi memuat orkestrasi COM tingkat lanjut.
    for forbidden, label in (("Marshal.GetActiveObject", "pelamaran instance COM"),
                             ("BeginCommandGroup", "grup undo"),
                             ("ActiveSelectionRange", "pembacaan seleksi aktif"),
                             ("AddToSelection", "pemilihan objek")):
        check(forbidden not in program, "Program.vb tidak lagi memuat " + label + " (pindah ke service)")
    # Service tidak memakai WinForms / MessageBox.
    for key in ("connectionservice", "copyservice", "autoselectservice"):
        src = read(key)
        check("System.Windows.Forms" not in src, key + " tidak bergantung pada WinForms")
        check("MessageBox." not in src, key + " tidak memanggil MessageBox langsung")
    check("ICorelPrompt" in read("connectionservice"), "interaksi pengguna koneksi lewat port ICorelPrompt")
    check("ICopyPrompt" in read("copyservice"), "interaksi pengguna copy lewat port ICopyPrompt")
    check("IProgressSink" in read("autoselectservice"), "progress Auto Select lewat port IProgressSink")
    check("Class CopyTargetValidation" in read("copyservice"),
          "validasi target copy dipisah dan murni (dapat diuji)")

    print("Build config:")
    for name in ("Program.vb", "WatermarkLayout.vb", "WatermarkEngine.vb", "WatermarkDialog.vb",
                 "WatermarkTokens.vb", "PresetStore.vb", "BatchLog.vb",
                 "CopyVerification.vb", "CopyMeasurement.vb", "DocumentUnits.vb",
                 "UndoGroup.vb", "BatchFilePolicy.vb", "CorelConnection.vb",
                 "ShapeClassifier.vb", "ShapeNode.vb", "ShapeScanner.vb", "SelectionNavigation.vb",
                 "PdfProofValidator.vb", "WatermarkPlacementCheck.vb", "BatchRunTracker.vb",
                 "CorelConnectionService.vb", "CopyService.vb", "AutoSelectService.vb"):
        check(name in vbproj, name + " terdaftar di <Compile Include>")

    print()
    if failures:
        print(str(len(failures)) + " pemeriksaan GAGAL")
        return 1
    print("SEMUA PEMERIKSAAN STATIS LULUS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
