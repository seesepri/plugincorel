Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  WatermarkPlacementTests.vb
'  Fokus B — akurasi penempatan: jangkar 3×3 + margin, rotasi (bounding box),
'  tile, ukuran halaman berbeda, dan konsistensi satuan mm/inch/point.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Drawing

Public Module WatermarkPlacementTests

    Private Function ExpectedLeftBottom(pageWidth As Double, pageHeight As Double,
                                        objWidth As Double, objHeight As Double,
                                        anchorPosition As WatermarkAnchor, marginMm As Double) As PointF
        Return WatermarkPlacementCheck.ExpectedLeftBottom(pageWidth, pageHeight, objWidth, objHeight,
                                                          anchorPosition, marginMm)
    End Function

    Public Sub Register(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("Watermark — akurasi penempatan")

        ' --- jangkar 3×3 + margin ---------------------------------------------
        suite.Add("jangkar 3×3 dengan margin 10 mm (A4)", Sub()
            Const W As Double = 210
            Const H As Double = 297
            Const m As Double = 10
            Const ow As Double = 50
            Const oh As Double = 20

            Dim topLeft As PointF = ExpectedLeftBottom(W, H, ow, oh, WatermarkAnchor.TopLeft, m)
            Check.AreEqual(10.0, Math.Round(topLeft.X, 3), "kiri atas: x = margin")
            Check.AreEqual(267.0, Math.Round(topLeft.Y, 3), "kiri atas: y = tinggi - margin - tinggi objek")

            Dim topCenter As PointF = ExpectedLeftBottom(W, H, ow, oh, WatermarkAnchor.TopCenter, m)
            Check.AreEqual(80.0, Math.Round(topCenter.X, 3), "tengah atas: x dipusatkan")

            Dim topRight As PointF = ExpectedLeftBottom(W, H, ow, oh, WatermarkAnchor.TopRight, m)
            Check.AreEqual(150.0, Math.Round(topRight.X, 3), "kanan atas: x = lebar - margin - lebar objek")

            Dim center As PointF = ExpectedLeftBottom(W, H, ow, oh, WatermarkAnchor.Center, m)
            Check.AreEqual(80.0, Math.Round(center.X, 3), "tengah: x")
            Check.AreEqual(138.5, Math.Round(center.Y, 3), "tengah: y")

            Dim bottomRight As PointF = ExpectedLeftBottom(W, H, ow, oh, WatermarkAnchor.BottomRight, m)
            Check.AreEqual(150.0, Math.Round(bottomRight.X, 3), "kanan bawah: x")
            Check.AreEqual(10.0, Math.Round(bottomRight.Y, 3), "kanan bawah: y = margin")
        End Sub)

        suite.Add("margin negatif diperlakukan 0 (tidak keluar halaman)", Sub()
            Dim pt As PointF = ExpectedLeftBottom(210, 297, 50, 20, WatermarkAnchor.BottomLeft, -25)
            Check.AreEqual(0.0, Math.Round(pt.X, 3), "x")
            Check.AreEqual(0.0, Math.Round(pt.Y, 3), "y")
        End Sub)

        suite.Add("objek lebih besar dari halaman dipusatkan, bukan keluar halaman", Sub()
            ' Teks 300 mm pada halaman A4 210 mm: dulu x menjadi negatif (terpotong).
            Dim pt As PointF = ExpectedLeftBottom(210, 297, 300, 20, WatermarkAnchor.TopLeft, 8)
            Check.AreEqual(-45.0, Math.Round(pt.X, 3), "objek > halaman: dipusatkan")
            Dim clamp As Double = WatermarkLayout.ClampAxis(-45, 210, 300)
            Check.AreEqual(-45.0, Math.Round(clamp, 3), "clamp memusatkan objek yang tidak muat")
            Check.IsFalse(WatermarkPlacementCheck.FitsOnPage(210, 297, 300, 20), "terdeteksi tidak muat")
        End Sub)

        suite.Add("margin besar pada halaman kecil tidak menghasilkan koordinat negatif", Sub()
            ' Halaman A5 148×210, margin 20, objek 120×10.
            Dim pt As PointF = ExpectedLeftBottom(148, 210, 120, 10, WatermarkAnchor.TopRight, 20)
            Check.IsTrue(pt.X >= 0, "x tidak negatif (" & pt.X & ")")
            Check.AreEqual(8.0, Math.Round(pt.X, 3), "x dijepit ke lebar - lebar objek")
        End Sub)

        ' --- rotasi ------------------------------------------------------------
        suite.Add("rotasi 90° menukar bounding box", Sub()
            Dim bounds As SizeInfo = WatermarkPlacementCheck.RotatedBounds(40, 10, 90)
            Check.Near(10.0, bounds.Width, 0.001, "lebar menjadi tinggi")
            Check.Near(40.0, bounds.Height, 0.001, "tinggi menjadi lebar")
        End Sub)

        suite.Add("rotasi 45° memperbesar bounding box sesuai rumus", Sub()
            Dim bounds As SizeInfo = WatermarkPlacementCheck.RotatedBounds(40, 10, 45)
            Dim expected As Double = (40 + 10) * Math.Sqrt(2) / 2.0
            Check.Near(expected, bounds.Width, 0.001, "lebar")
            Check.Near(expected, bounds.Height, 0.001, "tinggi")
        End Sub)

        suite.Add("rotasi 0° dan 360° tidak mengubah ukuran", Sub()
            Check.Near(40.0, WatermarkPlacementCheck.RotatedBounds(40, 10, 0).Width, 0.001, "0°")
            Check.Near(40.0, WatermarkPlacementCheck.RotatedBounds(40, 10, 360).Width, 0.001, "360°")
        End Sub)

        ' --- tile --------------------------------------------------------------
        suite.Add("tile: jarak lebih kecil dari ukuran objek tidak menumpuk", Sub()
            Dim count As Integer = WatermarkPlacementCheck.ExpectedTileCount(210, 297, 50, 20, 10, 5)
            Check.AreEqual(WatermarkLayout.TilePositions(210, 297, 50, 20, 10, 5).Count, count, "jumlah = fungsi resmi")
            Check.IsTrue(count >= 4, "minimal 4 posisi pada A4")
        End Sub)

        suite.Add("tile: posisi berada di dalam halaman", Sub()
            Dim positions As List(Of PointF) = WatermarkLayout.TilePositions(210, 297, 40, 20, 60, 45)
            Check.IsTrue(positions.Count >= 4, "ada posisi")
            For Each pt As PointF In positions
                Check.IsTrue(pt.X >= -0.001 AndAlso pt.X + 40 <= 210.001, "x dalam halaman (" & pt.X & ")")
                Check.IsTrue(pt.Y >= -0.001 AndAlso pt.Y + 20 <= 297.001, "y dalam halaman (" & pt.Y & ")")
            Next
        End Sub)

        suite.Add("tile: objek lebih besar dari halaman tetap satu posisi di tengah", Sub()
            Dim positions As List(Of PointF) = WatermarkLayout.TilePositions(100, 100, 300, 300, 60, 45)
            Check.AreEqual(1, positions.Count, "satu posisi")
            Check.AreEqual(-100.0, Math.Round(positions(0).X, 3), "dipusatkan")
        End Sub)

        ' --- verifikasi hasil --------------------------------------------------
        suite.Add("verifikasi: objek tepat -> WithinTolerance", Sub()
            Dim expected As PointF = WatermarkPlacementCheck.ExpectedCenter(210, 297, 50, 20, WatermarkAnchor.Center, 8)
            Dim verification As PlacementVerification = WatermarkPlacementCheck.Verify(
                expected, 50, 20, expected.X, expected.Y, 50, 20)
            Check.IsTrue(verification.WithinTolerance, "cocok")
            Check.Contains(verification.Describe(), "Δpusat", "deskripsi")
        End Sub)

        suite.Add("verifikasi: objek bergeser 2 mm -> gagal toleransi 0,5 mm", Sub()
            Dim expected As PointF = WatermarkPlacementCheck.ExpectedCenter(210, 297, 50, 20, WatermarkAnchor.Center, 8)
            Dim verification As PlacementVerification = WatermarkPlacementCheck.Verify(
                expected, 50, 20, expected.X + 2.0, expected.Y, 50, 20)
            Check.IsFalse(verification.WithinTolerance, "di luar toleransi")
            Check.Near(2.0, verification.DX, 0.001, "Δx terukur")
        End Sub)

        suite.Add("verifikasi: ukuran objek tidak terbaca -> tidak dianggap cocok", Sub()
            Dim expected As PointF = WatermarkPlacementCheck.ExpectedCenter(210, 297, 50, 20, WatermarkAnchor.Center, 8)
            Dim verification As PlacementVerification = WatermarkPlacementCheck.Verify(
                expected, 50, 20, expected.X, expected.Y, 0, 0)
            Check.IsFalse(verification.WithinTolerance, "tidak diklaim cocok")
            Check.IsFalse(verification.Measured, "ditandai tidak terukur")
        End Sub)

        ' --- satuan ------------------------------------------------------------
        suite.Add("ukuran halaman dikonversi dari inch ke mm", Sub()
            ' 8,27 × 11,69 inch = A4
            Dim size As SizeInfo = WatermarkPlacementCheck.PageSizeMillimeters(8.27, 11.69, DocumentUnits.Inch)
            Check.Near(210.06, size.Width, 0.1, "lebar mm")
            Check.Near(296.93, size.Height, 0.1, "tinggi mm")
        End Sub)

        suite.Add("ukuran halaman dikonversi dari point ke mm", Sub()
            Dim size As SizeInfo = WatermarkPlacementCheck.PageSizeMillimeters(595.28, 841.89, DocumentUnits.Point)
            Check.Near(210.0, size.Width, 0.1, "lebar mm")
            Check.Near(297.0, size.Height, 0.1, "tinggi mm")
        End Sub)

        suite.Add("satuan mm tidak berubah", Sub()
            Dim size As SizeInfo = WatermarkPlacementCheck.PageSizeMillimeters(210, 297, DocumentUnits.Millimeter)
            Check.Near(210.0, size.Width, 0.0001, "lebar")
            Check.Near(297.0, size.Height, 0.0001, "tinggi")
        End Sub)

        suite.Add("satuan tak dikenal -> tidak dipakai (NaN), bukan asumsi mm", Sub()
            Dim size As SizeInfo = WatermarkPlacementCheck.PageSizeMillimeters(210, 297, DocumentUnits.Unknown)
            Check.IsFalse(WatermarkPlacementCheck.PageSizeUsable(size), "tidak usable")
        End Sub)

        suite.Add("penempatan A4 dihitung sama setelah konversi dari inch", Sub()
            Dim sizeMm As SizeInfo = WatermarkPlacementCheck.PageSizeMillimeters(11.6929, 8.2677, DocumentUnits.Inch)
            Dim fromInch As PointF = ExpectedLeftBottom(sizeMm.Width, sizeMm.Height, 50, 20, WatermarkAnchor.TopLeft, 10)
            Dim fromMm As PointF = ExpectedLeftBottom(297, 210, 50, 20, WatermarkAnchor.TopLeft, 10)
            Check.Near(fromMm.X, fromInch.X, 0.1, "x konsisten")
            Check.Near(fromMm.Y, fromInch.Y, 0.1, "y konsisten")
        End Sub)

        suite.Add("label orientasi halaman", Sub()
            Check.AreEqual("portrait", WatermarkPlacementCheck.OrientationLabel(210, 297), "portrait")
            Check.AreEqual("landscape", WatermarkPlacementCheck.OrientationLabel(297, 210), "landscape")
            Check.AreEqual("persegi", WatermarkPlacementCheck.OrientationLabel(200, 200), "persegi")
        End Sub)
    End Sub
End Module
