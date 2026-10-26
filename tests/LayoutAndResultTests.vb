Option Strict Off
Imports System
Imports System.Collections.Generic
Imports System.IO

''' <summary>Model hasil, geometri layout, dan perilaku laporan.</summary>
Module LayoutAndResultTests
    Public Sub Run()
        TestMain.Section("WatermarkResult (laporan & tombol folder hasil)")
        Dim res As New WatermarkResult()
        TestMain.Check(Not res.CanOpenOutputFolder, "tanpa folder -> tombol nonaktif")
        res.OutputFolder = Path.GetTempPath()
        TestMain.Check(res.CanOpenOutputFolder, "folder yang ada -> tombol aktif")
        res.OutputFolder = Path.Combine(Path.GetTempPath(), "tidak-ada-xyz-123")
        TestMain.Check(Not res.CanOpenOutputFolder, "folder tidak ada -> tombol nonaktif")

        Dim a As New WatermarkResult()
        a.OutputFolder = "D:\hasil"
        a.OutputFile = "x_wm.cdr"
        a.Warnings.Add("Halaman 1: logo gagal")
        Dim b As New WatermarkResult()
        Dim c As New WatermarkResult()
        c.Merge(a)
        TestMain.Check(c.OutputFolder = "D:\hasil", "Merge mempertahankan folder hasil")
        TestMain.Check(c.OutputFile = "x_wm.cdr", "Merge mempertahankan file hasil")
        TestMain.Check(c.Warnings.Count = 1, "Merge ikut membawa peringatan")

        Dim gagal As New WatermarkResult()
        gagal.Failures.Add("Halaman 1: error")
        TestMain.Check(Not gagal.Ok, "ada kegagalan -> Ok = False")
        Dim hanyaPeringatan As New WatermarkResult()
        hanyaPeringatan.Warnings.Add("Halaman 1: logo gagal")
        TestMain.Check(hanyaPeringatan.Ok, "hanya peringatan -> Ok tetap True")

        TestMain.Check(New WatermarkResult().ToReport().Contains("Halaman diproses"), "laporan memuat jumlah halaman")
        Dim rep As String = hanyaPeringatan.ToReport()
        TestMain.Check(rep.Contains("Peringatan"), "laporan menampilkan bagian peringatan")
        TestMain.Check(rep.Contains("logo gagal"), "laporan memuat isi peringatan")

        TestMain.Section("Geometri layout")
        Dim d As WatermarkLayout.GeometryDelta = WatermarkLayout.CompareGeometry(10, 20, 30, 40, 10.2, 20.2, 30.2, 40.2)
        TestMain.Check(d.Matches(0.5), "selisih 0.2 mm dianggap cocok (toleransi 0.5)")
        Dim d2 As WatermarkLayout.GeometryDelta = WatermarkLayout.CompareGeometry(10, 20, 30, 40, 15, 20, 30, 40)
        TestMain.Check(Not d2.Matches(0.5), "selisih 5 mm dianggap tidak cocok")
        TestMain.Check(WatermarkLayout.IsWatermarkIdentity("WM_TEXT", "Layer 1"), "nama WM_* terdeteksi watermark")
        TestMain.Check(WatermarkLayout.IsWatermarkIdentity("Kotak", "WATERMARK"), "layer WATERMARK terdeteksi")
        TestMain.Check(Not WatermarkLayout.IsWatermarkIdentity("Kotak", "Layer 1"), "objek biasa bukan watermark")

        TestMain.Section("Token terpakai di template default")
        TestMain.Check(String.Join(",", WatermarkTokens.SupportedTokens()).Contains("order"), "daftar token memuat order")
        TestMain.Check(WatermarkTokens.HelpText().Contains("{page}"), "teks bantuan menyebut {page}")
    End Sub
End Module
