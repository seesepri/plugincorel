Option Strict Off
Imports System
Imports System.Collections.Generic

''' <summary>Token teks per halaman + ekstraksi nomor order dari nama file.</summary>
Module TokenTests
    Public Sub Run()
        TestMain.Section("Token teks per halaman ({page}/{order}/...)")

        Dim ctx As New WatermarkTokenContext()
        ctx.FileName = "INV-2026-0912 klien budi.cdr"
        ctx.PageIndex = 3
        ctx.PageCount = 12
        ctx.Now = New DateTime(2026, 10, 26, 14, 5, 0)

        TestMain.Check(WatermarkTokens.Expand("Hal {page}/{pages}", ctx) = "Hal 3/12", "token page + pages")
        TestMain.Check(WatermarkTokens.Expand("{file}", ctx) = "INV-2026-0912 klien budi", "token file tanpa ekstensi")
        TestMain.Check(WatermarkTokens.Expand("{filename}", ctx) = "INV-2026-0912 klien budi.cdr", "token filename lengkap")
        TestMain.Check(WatermarkTokens.Expand("{order}", ctx) = "INV-2026-0912", "token order dari nama file")
        TestMain.Check(WatermarkTokens.Expand("{date}", ctx) = "2026-10-26", "token date yyyy-MM-dd")
        TestMain.Check(WatermarkTokens.Expand("{time}", ctx) = "14:05", "token time HH:mm")
        TestMain.Check(WatermarkTokens.Expand("ORDER {ORDER}", ctx) = "ORDER INV-2026-0912", "token case-insensitive")
        TestMain.Check(WatermarkTokens.Expand("{tidak_ada}", ctx) = "{tidak_ada}", "token asing dibiarkan")
        TestMain.Check(WatermarkTokens.Expand("{}", ctx) = "{}", "{} tidak crash")
        TestMain.Check(WatermarkTokens.Expand("teks biasa", ctx) = "teks biasa", "teks tanpa token tidak berubah")
        TestMain.Check(WatermarkTokens.Expand("", ctx) = "", "template kosong")
        TestMain.Check(WatermarkTokens.Expand("{page}", Nothing) = "1", "context Nothing -> halaman 1")

        TestMain.Check(WatermarkTokens.HasToken("{page}", "page"), "HasToken mendeteksi {page}")
        TestMain.Check(WatermarkTokens.HasToken("{PAGE}", "page"), "HasToken case-insensitive")
        TestMain.Check(Not WatermarkTokens.HasToken("Hal {pages}", "page"), "HasToken tidak salah cocok {page} vs {pages}")
        TestMain.Check(Not WatermarkTokens.HasToken("CONTOH", "page"), "HasToken false untuk teks biasa")

        TestMain.Section("Ekstraksi nomor order dari nama file")
        TestMain.Check(WatermarkTokens.ExtractOrder("INV-2026-0912 klien budi.cdr") = "INV-2026-0912", "tanda hubung dipertahankan")
        TestMain.Check(WatermarkTokens.ExtractOrder("0123_order mei.cdr") = "0123_order", "garis bawah dipertahankan, spasi pemisah")
        TestMain.Check(WatermarkTokens.ExtractOrder("PO_2026_0042.cdr") = "PO_2026_0042", "tanpa spasi -> seluruh nama")
        TestMain.Check(WatermarkTokens.ExtractOrder("desain final.cdr") = "desain final", "tanpa angka -> nama utuh")
        TestMain.Check(WatermarkTokens.ExtractOrder("ab12 cd.cdr") = "ab12 cd", "angka < 4 digit bukan order")
        TestMain.Check(WatermarkTokens.ExtractOrder("C:\klien\SO-9911.cdr") = "SO-9911", "path Windows dibuang")
        TestMain.Check(WatermarkTokens.ExtractOrder("/home/x/PO_1234.cdr") = "PO_1234", "path gaya Unix dibuang")
        TestMain.Check(WatermarkTokens.ExtractOrder("") = "", "nama kosong aman")
        TestMain.Check(WatermarkTokens.ExtractOrder(Nothing) = "", "Nothing aman")
        TestMain.Check(WatermarkTokens.ExtractOrder("12345.cdr") = "12345", "nama angka saja")
        TestMain.Check(WatermarkTokens.ExtractOrder("-INV-2026-0912-.cdr") = "INV-2026-0912", "tanda hubung tepi dibersihkan")
    End Sub
End Module
