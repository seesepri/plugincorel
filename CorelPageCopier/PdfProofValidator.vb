Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  PdfProofValidator.vb
'  Validasi struktur file PDF proof hasil ekspor CorelDRAW.
'
'  Kenapa file ini ada:
'    Kode lama menganggap "PublishToPDF tidak melempar exception" = ekspor
'    sukses. Bahkan file yang TIDAK PERNAH dibuat dilaporkan sebagai
'    "PDF proof: <path>". Itu klaim tanpa bukti.
'
'  Yang divalidasi di sini (murni, tanpa CorelDRAW maupun library PDF luar):
'    1. Header %PDF-1.x ada di awal file.
'    2. Trailer %%EOF ada di akhir file.
'    3. startxref ada (tanda struktur xref lengkap).
'    4. Jumlah objek halaman (/Type /Page) — inilah jumlah halaman PDF.
'    5. Ukuran MediaBox per halaman (lebar×tinggi dalam point) dan /Rotate.
'
'  Yang TIDAK dapat divalidasi dari struktur PDF (dan karena itu TIDAK
'  diklaim): apakah watermark benar-benar terlihat, opasitas, clipping, dan
'  posisi visual. Bagian itu wajib diperiksa manual di CorelDRAW / pembaca PDF
'  (lihat PANDUAN-UJI.md §12). Kelas ini sengaja mengembalikan status
'  "Unverifiable" untuk hal yang tidak bisa dipastikan, bukan "Valid".
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

''' <summary>Status akhir validasi PDF proof.</summary>
Public Enum PdfProofStatus
    ''' <summary>Struktur PDF terbaca dan cocok dengan dokumen sumber.</summary>
    Valid = 0
    ''' <summary>Struktur PDF rusak atau tidak cocok dengan dokumen sumber.</summary>
    Invalid = 1
    ''' <summary>Struktur tidak dapat dipastikan otomatis (mis. object stream) — perlu cek manual.</summary>
    Unverifiable = 2
End Enum

''' <summary>Ukuran satu halaman PDF (dalam point, 1 pt = 1/72 inch).</summary>
Public Class PdfPageInfo
    Public WidthPoints As Double = 0
    Public HeightPoints As Double = 0
    Public Rotate As Integer = 0
    ''' <summary>False bila MediaBox halaman tidak ditemukan (mis. diwarisi dari /Pages).</summary>
    Public SizeKnown As Boolean = False

    Public ReadOnly Property IsLandscape As Boolean
        Get
            Return SizeKnown AndAlso WidthPoints > HeightPoints
        End Get
    End Property

    ''' <summary>Ukuran setelah memperhitungkan /Rotate 90 atau 270 (mis. tampilan landscape).</summary>
    Public Function DisplaySize() As SizeInfo
        Dim info As New SizeInfo()
        info.Width = WidthPoints
        info.Height = HeightPoints
        Dim normalized As Integer = ((Rotate Mod 360) + 360) Mod 360
        If normalized = 90 OrElse normalized = 270 Then
            info.Width = HeightPoints
            info.Height = WidthPoints
        End If
        Return info
    End Function
End Class

''' <summary>Pasangan lebar/tinggi sederhana (menghindari ketergantungan WinForms di test).</summary>
Public Structure SizeInfo
    Public Width As Double
    Public Height As Double
End Structure

''' <summary>Ukuran halaman yang DIHARAPKAN, dibangun dari dokumen sumber.</summary>
Public Class PdfProofExpectation
    Public ExpectedPageCount As Integer = 0
    Public PageWidthsPoints As New List(Of Double)()
    Public PageHeightsPoints As New List(Of Double)()
    ''' <summary>Toleransi ukuran (point). Pembulatan PublishToPDF ~0,5 pt.</summary>
    Public TolerancePoints As Double = 1.0

    ''' <summary>Bangun harapan dari ukuran halaman dalam milimeter (urutan halaman 1..N).</summary>
    Public Shared Function FromMillimeters(widthsMm As IList(Of Double),
                                           heightsMm As IList(Of Double)) As PdfProofExpectation
        Dim expectation As New PdfProofExpectation()
        If widthsMm Is Nothing OrElse heightsMm Is Nothing Then Return expectation
        Dim count As Integer = Math.Min(widthsMm.Count, heightsMm.Count)
        expectation.ExpectedPageCount = count
        For i As Integer = 0 To count - 1
            expectation.PageWidthsPoints.Add(PdfProofValidator.MillimetersToPoints(widthsMm(i)))
            expectation.PageHeightsPoints.Add(PdfProofValidator.MillimetersToPoints(heightsMm(i)))
        Next
        Return expectation
    End Function

    Public ReadOnly Property HasSizes As Boolean
        Get
            Return PageWidthsPoints.Count > 0
        End Get
    End Property
End Class

''' <summary>Hasil validasi satu file PDF.</summary>
Public Class PdfProofValidation
    Public FilePath As String = ""
    Public Status As PdfProofStatus = PdfProofStatus.Invalid
    Public FileSizeBytes As Long = 0
    Public HeaderOk As Boolean = False
    Public EofOk As Boolean = False
    Public StartXRefOk As Boolean = False
    ''' <summary>Jumlah halaman yang terbaca dari struktur (0 = tidak terbaca).</summary>
    Public PageCount As Integer = 0
    Public Pages As New List(Of PdfPageInfo)()
    ''' <summary>Masalah yang membuat PDF TIDAK layak dianggap sukses.</summary>
    Public Problems As New List(Of String)()
    ''' <summary>Catatan non-fatal (mis. ukuran tidak dapat dibandingkan).</summary>
    Public Notes As New List(Of String)()

    Public ReadOnly Property IsValid As Boolean
        Get
            Return Status = PdfProofStatus.Valid
        End Get
    End Property

    Public Function SummaryLine() As String
        Select Case Status
            Case PdfProofStatus.Valid
                Return "PDF VALID (" & PageCount.ToString(CultureInfo.InvariantCulture) & " halaman, struktur & ukuran cocok)"
            Case PdfProofStatus.Unverifiable
                Return "PDF TIDAK DAPAT DIVERIFIKASI OTOMATIS (" &
                       If(Problems.Count > 0, Problems(0), "struktur tidak terbaca") & ")"
            Case Else
                Return "PDF TIDAK VALID (" & If(Problems.Count > 0, Problems(0), "struktur rusak") & ")"
        End Select
    End Function

    Public Function Describe() As String
        Dim sb As New StringBuilder()
        sb.AppendLine("Validasi PDF : " & SummaryLine())
        sb.AppendLine("  File       : " & If(String.IsNullOrEmpty(FilePath), "(tidak ada)", FilePath))
        sb.AppendLine("  Ukuran     : " & FileSizeBytes.ToString("N0", CultureInfo.InvariantCulture) & " byte")
        sb.AppendLine("  Struktur   : header=" & YesNo(HeaderOk) & ", %%EOF=" & YesNo(EofOk) &
                      ", startxref=" & YesNo(StartXRefOk))
        sb.AppendLine("  Halaman    : " & PageCount.ToString(CultureInfo.InvariantCulture))
        For i As Integer = 0 To Pages.Count - 1
            Dim page As PdfPageInfo = Pages(i)
            If page.SizeKnown Then
                Dim display As SizeInfo = page.DisplaySize()
                sb.AppendLine("    " & (i + 1).ToString(CultureInfo.InvariantCulture) & ". " &
                              display.Width.ToString("0.0", CultureInfo.InvariantCulture) & " × " &
                              display.Height.ToString("0.0", CultureInfo.InvariantCulture) & " pt" &
                              If(page.Rotate <> 0, " (Rotate " & page.Rotate & "°)", ""))
            Else
                sb.AppendLine("    " & (i + 1).ToString(CultureInfo.InvariantCulture) & ". (ukuran tidak terbaca)")
            End If
        Next
        For Each problem As String In Problems
            sb.AppendLine("  • " & problem)
        Next
        For Each note As String In Notes
            sb.AppendLine("  catatan: " & note)
        Next
        Return sb.ToString()
    End Function

    Private Shared Function YesNo(value As Boolean) As String
        Return If(value, "ya", "TIDAK")
    End Function
End Class

''' <summary>
''' Pemeriksa struktur PDF. Semua metode statis dan murni sehingga bisa diuji
''' dengan byte PDF buatan (tanpa CorelDRAW).
''' </summary>
Public Class PdfProofValidator

    ''' <summary>Konversi mm ke point PDF (1 pt = 1/72 inch).</summary>
    Public Shared Function MillimetersToPoints(mm As Double) As Double
        Return mm / 25.4 * 72.0
    End Function

    ''' <summary>Toleransi bawaan perbandingan ukuran halaman (point).</summary>
    Public Const DefaultTolerancePoints As Double = 1.0

    ''' <summary>Batas ukuran file yang dibaca (PDF proof normal < 50 MB).</summary>
    Public Const MaxReadBytes As Long = 64L * 1024L * 1024L

    ''' <summary>Validasi file di disk. File tidak ada / tidak terbaca = Invalid (bukan sukses).</summary>
    Public Shared Function ValidateFile(filePath As String,
                                        Optional expected As PdfProofExpectation = Nothing) As PdfProofValidation
        Dim validation As New PdfProofValidation()
        validation.FilePath = If(filePath, "")

        If String.IsNullOrEmpty(filePath) Then
            validation.Problems.Add("path PDF kosong.")
            validation.Status = PdfProofStatus.Invalid
            Return validation
        End If

        Dim info As FileInfo
        Try
            info = New FileInfo(filePath)
        Catch ex As Exception
            validation.Problems.Add("path PDF tidak valid (" & ex.GetType().Name & ").")
            validation.Status = PdfProofStatus.Invalid
            Return validation
        End Try

        If Not info.Exists Then
            validation.Problems.Add("file PDF tidak ditemukan — ekspor tidak menghasilkan file.")
            validation.Status = PdfProofStatus.Invalid
            Return validation
        End If

        validation.FileSizeBytes = info.Length
        If info.Length = 0 Then
            validation.Problems.Add("file PDF berukuran 0 byte.")
            validation.Status = PdfProofStatus.Invalid
            Return validation
        End If
        If info.Length > MaxReadBytes Then
            validation.Problems.Add("file PDF lebih besar dari batas baca " & MaxReadBytes & " byte.")
            validation.Status = PdfProofStatus.Unverifiable
            Return validation
        End If

        Dim bytes As Byte()
        Try
            bytes = File.ReadAllBytes(filePath)
        Catch ex As Exception
            validation.Problems.Add("file PDF tidak dapat dibaca (" & ex.GetType().Name & ": " & ex.Message & ").")
            validation.Status = PdfProofStatus.Invalid
            Return validation
        End Try

        Return ValidateBytes(bytes, expected, filePath)
    End Function

    ''' <summary>Validasi isi PDF (byte) — dipakai test dan oleh ValidateFile.</summary>
    Public Shared Function ValidateBytes(bytes As Byte(), Optional expected As PdfProofExpectation = Nothing,
                                         Optional filePath As String = "") As PdfProofValidation
        Dim validation As New PdfProofValidation()
        validation.FilePath = If(filePath, "")
        If bytes Is Nothing OrElse bytes.Length = 0 Then
            validation.Problems.Add("isi PDF kosong.")
            validation.Status = PdfProofStatus.Invalid
            Return validation
        End If
        validation.FileSizeBytes = bytes.LongLength

        ' Latin-1: satu byte = satu karakter, sehingga pencarian teks aman
        ' untuk PDF yang memuat data biner terkompresi.
        Dim text As String = Encoding.GetEncoding(28591).GetString(bytes)

        validation.HeaderOk = HasHeader(text)
        validation.EofOk = HasEof(text)
        validation.StartXRefOk = HasStartXRef(text)

        If Not validation.HeaderOk Then
            validation.Problems.Add("header %PDF tidak ditemukan (bukan file PDF).")
        End If
        If Not validation.EofOk Then
            validation.Problems.Add("trailer %%EOF tidak ditemukan (file terpotong / ekspor gagal di tengah).")
        End If
        If Not validation.StartXRefOk Then
            validation.Problems.Add("startxref tidak ditemukan (tabel silang PDF tidak lengkap).")
        End If

        ' --- Halaman: hitung objek /Type /Page (bukan /Pages) -----------------
        Dim pages As List(Of PdfPageInfo) = ReadPages(text)
        validation.PageCount = pages.Count
        validation.Pages = pages

        Dim usesObjectStreams As Boolean = text.IndexOf("/ObjStm", StringComparison.Ordinal) >= 0

        If pages.Count = 0 Then
            If usesObjectStreams Then
                validation.Problems.Add(
                    "objek halaman dikemas dalam object stream dan tidak dapat dihitung otomatis — " &
                    "jumlah halaman WAJIB dicek manual di pembaca PDF.")
            Else
                validation.Problems.Add("tidak ada objek halaman (/Type /Page) yang terbaca — file bukan PDF normal.")
            End If
        End If

        ' --- Bandingkan dengan dokumen sumber --------------------------------
        Dim expectationAnswer As String = CompareWithExpectation(validation, expected)
        If expectationAnswer = "unverifiable" Then
            ' Jangan pernah menaikkan status menjadi Valid bila pembandingan gagal.
        End If

        If validation.Problems.Count > 0 Then
            ' Struktur rusak selalu Invalid; ketidakmampuan membaca (object stream)
            ' tidak boleh disebut Invalid karena bisa jadi PDF-nya benar.
            If usesObjectStreams AndAlso validation.HeaderOk AndAlso validation.EofOk AndAlso
               validation.StartXRefOk AndAlso validation.Problems.Count = 1 Then
                validation.Status = PdfProofStatus.Unverifiable
            Else
                validation.Status = PdfProofStatus.Invalid
            End If
        Else
            validation.Status = PdfProofStatus.Valid
        End If

        Return validation
    End Function

    Private Shared Function CompareWithExpectation(validation As PdfProofValidation,
                                                   expected As PdfProofExpectation) As String
        If expected Is Nothing OrElse expected.ExpectedPageCount <= 0 Then
            validation.Notes.Add("jumlah halaman yang diharapkan tidak diberikan — hanya struktur yang divalidasi.")
            Return "skipped"
        End If

        If validation.PageCount > 0 AndAlso validation.PageCount <> expected.ExpectedPageCount Then
            validation.Problems.Add("jumlah halaman PDF " & validation.PageCount &
                                    " berbeda dari dokumen sumber " & expected.ExpectedPageCount & ".")
        End If

        If Not expected.HasSizes OrElse validation.Pages.Count = 0 Then
            validation.Notes.Add("ukuran halaman tidak dapat dibandingkan (MediaBox tidak terbaca).")
            Return "unverifiable"
        End If

        Dim tolerance As Double = If(expected.TolerancePoints > 0, expected.TolerancePoints, DefaultTolerancePoints)
        Dim compared As Integer = 0
        For i As Integer = 0 To Math.Min(validation.Pages.Count, expected.ExpectedPageCount) - 1
            Dim page As PdfPageInfo = validation.Pages(i)
            If Not page.SizeKnown Then Continue For

            Dim display As SizeInfo = page.DisplaySize()
            Dim expectedWidth As Double = expected.PageWidthsPoints(i)
            Dim expectedHeight As Double = expected.PageHeightsPoints(i)

            Dim directMatch As Boolean =
                Math.Abs(display.Width - expectedWidth) <= tolerance AndAlso
                Math.Abs(display.Height - expectedHeight) <= tolerance
            Dim swappedMatch As Boolean =
                Math.Abs(display.Width - expectedHeight) <= tolerance AndAlso
                Math.Abs(display.Height - expectedWidth) <= tolerance

            compared += 1
            If directMatch Then Continue For

            If swappedMatch Then
                validation.Problems.Add(
                    "halaman " & (i + 1) & ": orientasi/ukuran tertukar (" &
                    display.Width.ToString("0.0", CultureInfo.InvariantCulture) & " × " &
                    display.Height.ToString("0.0", CultureInfo.InvariantCulture) & " pt, " &
                    "diharapkan " & expectedWidth.ToString("0.0", CultureInfo.InvariantCulture) & " × " &
                    expectedHeight.ToString("0.0", CultureInfo.InvariantCulture) & " pt).")
            Else
                validation.Problems.Add(
                    "halaman " & (i + 1) & ": ukuran PDF " &
                    display.Width.ToString("0.0", CultureInfo.InvariantCulture) & " × " &
                    display.Height.ToString("0.0", CultureInfo.InvariantCulture) & " pt, diharapkan " &
                    expectedWidth.ToString("0.0", CultureInfo.InvariantCulture) & " × " &
                    expectedHeight.ToString("0.0", CultureInfo.InvariantCulture) & " pt.")
            End If
        Next

        If compared = 0 Then
            validation.Notes.Add("tidak ada halaman yang ukurannya bisa dibandingkan.")
            Return "unverifiable"
        End If
        Return "compared"
    End Function

    ' ========================================================================
    '  PARSER SEDERHANA (cukup untuk pemeriksaan struktur, bukan renderer)
    ' ========================================================================

    Private Shared Function HasHeader(text As String) As Boolean
        Dim limit As Integer = Math.Min(1024, text.Length)
        Return text.Substring(0, limit).IndexOf("%PDF-", StringComparison.Ordinal) >= 0
    End Function

    Private Shared Function HasEof(text As String) As Boolean
        Dim start As Integer = Math.Max(0, text.Length - 2048)
        Return text.Substring(start).IndexOf("%%EOF", StringComparison.Ordinal) >= 0
    End Function

    Private Shared Function HasStartXRef(text As String) As Boolean
        Dim start As Integer = Math.Max(0, text.Length - 8192)
        Return text.Substring(start).IndexOf("startxref", StringComparison.Ordinal) >= 0
    End Function

    Private Shared ReadOnly PageTypePattern As New Regex("/Type\s*/Page(?![A-Za-z])",
                                                          RegexOptions.Compiled Or RegexOptions.CultureInvariant)
    Private Shared ReadOnly MediaBoxPattern As New Regex(
        "/MediaBox\s*\[\s*(-?[\d.]+)\s+(-?[\d.]+)\s+(-?[\d.]+)\s+(-?[\d.]+)\s*\]",
        RegexOptions.Compiled Or RegexOptions.CultureInvariant)
    Private Shared ReadOnly RotatePattern As New Regex("/Rotate\s*(-?\d+)",
        RegexOptions.Compiled Or RegexOptions.CultureInvariant)

    ''' <summary>
    ''' Baca objek halaman. MediaBox dicari di sekitar objek halaman; bila
    ''' diwarisi dari /Pages, ukuran ditandai tidak diketahui (bukan ditebak).
    ''' </summary>
    Private Shared Function ReadPages(text As String) As List(Of PdfPageInfo)
        Dim pages As New List(Of PdfPageInfo)()
        For Each match As Match In PageTypePattern.Matches(text)
            Dim page As New PdfPageInfo()
            Dim windowStart As Integer = match.Index
            Dim windowLength As Integer = Math.Min(4000, text.Length - windowStart)
            Dim window As String = text.Substring(windowStart, windowLength)

            Dim box As Match = MediaBoxPattern.Match(window)
            If box.Success Then
                Dim x1 As Double = ParseDouble(box.Groups(1).Value)
                Dim y1 As Double = ParseDouble(box.Groups(2).Value)
                Dim x2 As Double = ParseDouble(box.Groups(3).Value)
                Dim y2 As Double = ParseDouble(box.Groups(4).Value)
                page.WidthPoints = Math.Abs(x2 - x1)
                page.HeightPoints = Math.Abs(y2 - y1)
                page.SizeKnown = page.WidthPoints > 0 AndAlso page.HeightPoints > 0
            End If

            Dim rotate As Match = RotatePattern.Match(window)
            If rotate.Success Then
                Dim value As Integer
                If Integer.TryParse(rotate.Groups(1).Value, NumberStyles.Integer,
                                    CultureInfo.InvariantCulture, value) Then
                    page.Rotate = value
                End If
            End If

            pages.Add(page)
        Next
        Return pages
    End Function

    Private Shared Function ParseDouble(value As String) As Double
        Dim parsed As Double
        If Double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, parsed) Then Return parsed
        Return 0
    End Function
End Class
