#!/usr/bin/env python3
"""
Static regression checks untuk Corel Page Copier.

Menjaga tiga perbaikan penting tetap ada saat kode diubah:
  #1  Copy dibungkus BeginCommandGroup/EndCommandGroup (satu Ctrl+Z).
  #2  Auto Select mengecualikan objek watermark (WM_* / layer WATERMARK).
  #3  Satuan dokumen dinormalisasi ke mm selama operasi dan dipulihkan setelahnya.
  Fase 1.5  Semua operasi panjang dijaga flag busy (anti re-entrancy DoEvents).
  UI        Tombol "Watermark..." ter-hook ke WatermarkDialog + logo bawaan tertanam.
  Fase 2    Preset %APPDATA%, log CSV batch, ekspor PDF proof.

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
    "vbproj": os.path.join(SRC, "CorelPageCopier.vbproj"),
}

failures = []


def read(key):
    with open(files[key], encoding="utf-8") as fh:
        return fh.read()


def check(condition, message):
    if condition:
        print("  PASS  " + message)
    else:
        failures.append(message)
        print("  FAIL  " + message)


def method_body(source, name):
    """Ambil badan satu method: Sub maupun Function, sampai End Sub/End Function pertama."""
    for kind, terminator in (("Sub ", "End Sub"), ("Function ", "End Function")):
        start = source.find(kind + name + "(")
        if start >= 0:
            end = source.find(terminator, start)
            return source[start:end] if end > 0 else source[start:]
    return ""


def main():
    program = read("program")
    layout = read("layout")
    engine = read("engine")
    vbproj = read("vbproj")

    copy_wrapper = method_body(program, "CopyClicked")
    copy_body = method_body(program, "CopyClickedCore")   # isi sebenarnya (wrapper busy ada di atasnya)
    collect_body = method_body(program, "CollectMatchingShape")
    begin_body = method_body(program, "TryBeginOperation")
    end_body = method_body(program, "EndOperation")
    closing_body = method_body(program, "CopierFormClosing")

    print("Temuan #1 — undo grouping pada Copy:")
    check("CopyClickedCore()" in copy_wrapper, "CopyClicked meneruskan ke CopyClickedCore")
    check("BeginCommandGroup" in copy_body, "CopyClicked memanggil BeginCommandGroup")
    check("EndCommandGroup" in copy_body, "CopyClicked memanggil EndCommandGroup")
    check(re.search(r"Finally[\s\S]*?EndCommandGroup", copy_body) is not None,
          "EndCommandGroup berada di blok Finally (selalu dieksekusi)")
    begins = len(re.findall(r"\.BeginCommandGroup\s*\(", program))
    ends = len(re.findall(r"\.EndCommandGroup\s*\(", program))
    check(begins == ends and begins >= 1,
          "jumlah pemanggilan BeginCommandGroup (%d) == EndCommandGroup (%d) di Program.vb" % (begins, ends))

    print("Temuan #1 — undo grouping pada watermark:")
    check("BeginCommandGroup" in read("engine"), "WatermarkEngine memakai BeginCommandGroup")

    print("Temuan #2 — Auto Select mengecualikan watermark:")
    check("IsWatermarkShape(shape, depth)" in collect_body,
          "CollectMatchingShape menyaring watermark sebelum menambahkan objek")
    check(len(re.findall(r"IsWatermarkShape\(shape, depth\)", collect_body)) >= 1,
          "filter diterapkan tepat sekali per shape (bukan per kategori)")
    check("IsWatermarkIdentity" in program, "CopierForm memakai helper identitas bersama")

    print("Temuan #3 — normalisasi satuan dokumen:")
    check("CDR_UNIT_MM" in program, "konstanta CDR_UNIT_MM (=3) didefinisikan")
    check("corelDoc.Unit = CDR_UNIT_MM" in copy_body, "satuan di-set ke mm sebelum operasi")
    check("corelDoc.Unit = previousUnit" in copy_body, "satuan asli dipulihkan")
    check(re.search(r"Finally[\s\S]*?corelDoc\.Unit = previousUnit", copy_body) is not None,
          "pemulihan satuan berada di blok Finally")
    check("GEOMETRY_TOLERANCE_MM" in program, "toleransi geometri eksplisit")

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
    check("New WatermarkDialog(corelApp, selectedPages)" in program,
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

    print("Build config:")
    for name in ("Program.vb", "WatermarkLayout.vb", "WatermarkEngine.vb", "WatermarkDialog.vb",
                 "PresetStore.vb", "BatchLog.vb"):
        check(name in vbproj, name + " terdaftar di <Compile Include>")

    print()
    if failures:
        print(str(len(failures)) + " pemeriksaan GAGAL")
        return 1
    print("SEMUA PEMERIKSAAN STATIS LULUS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
