#!/usr/bin/env python3
"""
Static regression checks untuk Corel Page Copier.

Menjaga tiga perbaikan penting tetap ada saat kode diubah:
  #1  Copy dibungkus BeginCommandGroup/EndCommandGroup (satu Ctrl+Z).
  #2  Auto Select mengecualikan objek watermark (WM_* / layer WATERMARK).
  #3  Satuan dokumen dinormalisasi ke mm selama operasi dan dipulihkan setelahnya.

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
    """Ambil badan satu method (dari 'Sub name(' sampai 'End Sub' pertama)."""
    start = source.find("Sub " + name + "(")
    if start < 0:
        return ""
    end = source.find("End Sub", start)
    return source[start:end]


def main():
    program = read("program")
    layout = read("layout")
    engine = read("engine")
    vbproj = read("vbproj")

    copy_body = method_body(program, "CopyClicked")
    collect_body = method_body(program, "CollectMatchingShape")

    print("Temuan #1 — undo grouping pada Copy:")
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

    print("Build config:")
    for name in ("Program.vb", "WatermarkLayout.vb", "WatermarkEngine.vb", "WatermarkDialog.vb"):
        check(name in vbproj, name + " terdaftar di <Compile Include>")

    print()
    if failures:
        print(str(len(failures)) + " pemeriksaan GAGAL")
        return 1
    print("SEMUA PEMERIKSAAN STATIS LULUS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
