Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  PdfFixture.vb
'  Pembuat byte PDF sintetis untuk menguji PdfProofValidator TANPA CorelDRAW.
'  Struktur yang dihasilkan menyerupai PDF hasil PublishToPDF (xref klasik,
'  satu objek /Type /Page per halaman dengan /MediaBox).
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text

Public Module PdfFixture

    ''' <summary>PDF valid dengan ukuran halaman tertentu (point).</summary>
    Public Function BuildPdf(widthsPoints As IList(Of Double), heightsPoints As IList(Of Double),
                             Optional rotatePerPage As IList(Of Integer) = Nothing,
                             Optional includeEof As Boolean = True,
                             Optional includeStartXRef As Boolean = True,
                             Optional includeMediaBox As Boolean = True,
                             Optional header As String = "%PDF-1.4") As Byte()
        Dim count As Integer = widthsPoints.Count
        Dim sb As New StringBuilder()
        sb.AppendLine(header)
        sb.AppendLine("1 0 obj")
        sb.AppendLine("<< /Type /Catalog /Pages 2 0 R >>")
        sb.AppendLine("endobj")
        sb.AppendLine("2 0 obj")
        sb.AppendLine("<< /Type /Pages /Count " & count.ToString(CultureInfo.InvariantCulture) & " >>")
        sb.AppendLine("endobj")

        For i As Integer = 0 To count - 1
            Dim objectNumber As Integer = 3 + i
            sb.AppendLine(objectNumber.ToString(CultureInfo.InvariantCulture) & " 0 obj")
            Dim body As String = "<< /Type /Page /Parent 2 0 R "
            If includeMediaBox Then
                body = body & "/MediaBox [0 0 " &
                       widthsPoints(i).ToString("0.####", CultureInfo.InvariantCulture) & " " &
                       heightsPoints(i).ToString("0.####", CultureInfo.InvariantCulture) & "] "
            End If
            If rotatePerPage IsNot Nothing AndAlso i < rotatePerPage.Count AndAlso rotatePerPage(i) <> 0 Then
                body = body & "/Rotate " & rotatePerPage(i).ToString(CultureInfo.InvariantCulture) & " "
            End If
            sb.AppendLine(body & ">>")
            sb.AppendLine("endobj")
        Next

        sb.AppendLine("xref")
        sb.AppendLine("0 " & (count + 3).ToString(CultureInfo.InvariantCulture))
        sb.AppendLine("0000000000 65535 f ")
        sb.AppendLine("trailer")
        sb.AppendLine("<< /Size " & (count + 3).ToString(CultureInfo.InvariantCulture) & " /Root 1 0 R >>")
        If includeStartXRef Then sb.AppendLine("startxref")
        sb.AppendLine("9")
        If includeEof Then sb.Append("%%EOF")
        Return Encoding.GetEncoding(28591).GetBytes(sb.ToString())
    End Function

    ''' <summary>PDF 3 halaman A4 (ukuran halaman dalam point).</summary>
    Public Function A4ThreePages() As Byte()
        Return BuildPdf(New Double() {595.28, 595.28, 595.28},
                        New Double() {841.89, 841.89, 841.89})
    End Function

    ''' <summary>PDF yang memakai object stream (halaman tidak terbaca sebagai objek biasa).</summary>
    Public Function ObjectStreamPdf() As Byte()
        Dim sb As New StringBuilder()
        sb.AppendLine("%PDF-1.7")
        sb.AppendLine("1 0 obj")
        sb.AppendLine("<< /Type /ObjStm /N 2 /First 8 >>")
        sb.AppendLine("stream")
        sb.AppendLine("endstream")
        sb.AppendLine("endobj")
        sb.AppendLine("startxref")
        sb.AppendLine("9")
        sb.Append("%%EOF")
        Return Encoding.GetEncoding(28591).GetBytes(sb.ToString())
    End Function

    ''' <summary>Harapan ukuran halaman dari ukuran milimeter.</summary>
    Public Function ExpectationMm(widthsMm As IList(Of Double), heightsMm As IList(Of Double)) As PdfProofExpectation
        Return PdfProofExpectation.FromMillimeters(widthsMm, heightsMm)
    End Function
End Module
