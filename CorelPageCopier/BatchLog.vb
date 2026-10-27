Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  BatchLog.vb
'  Log CSV untuk operasi watermark (single & batch).
'
'  Kenapa ';' sebagai pemisah: Excel dengan regional setting Indonesia memakai
'  ';' (koma dipakai sebagai pemisah desimal), jadi CSV ini langsung terbuka
'  rapi tanpa fitur import. File ditulis UTF-8 + BOM agar huruf beraksen benar.
'  Semua pembentukan baris bersifat murni -> bisa diuji unit.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports Microsoft.VisualBasic
Imports System.Globalization
Imports System.IO
Imports System.Text

''' <summary>Satu baris log (satu file/proses).</summary>
Public Class BatchLogEntry
    Public Timestamp As DateTime = DateTime.Now
    Public FileName As String = ""
    Public Pages As Integer = 0
    Public Shapes As Integer = 0
    Public Status As String = ""
    Public Message As String = ""
    Public DurationSeconds As Double = 0
End Class

Public Class BatchLog

    Public Const Delimiter As Char = ";"c
    Public Const DefaultFileName As String = "watermark-log.csv"

    ''' <summary>Header kolom (dibuat sekali saat file pertama kali ditulis).</summary>
    Public Shared Function Header() As String
        Return String.Join(Delimiter.ToString(), New String() {
            "Waktu", "File", "Halaman", "Shape", "Status", "Pesan", "Durasi_detik"
        })
    End Function

    ''' <summary>Escape satu sel CSV: dibungkus kutip bila memuat pemisah/kutip/newline.</summary>
    Public Shared Function EscapeCell(value As String) As String
        Dim text As String = If(value, "")
        Dim needsQuote As Boolean = text.IndexOf(Delimiter) >= 0 OrElse
                                    text.IndexOf(""""c) >= 0 OrElse
                                    text.IndexOf(vbLf) >= 0 OrElse
                                    text.IndexOf(vbCr) >= 0 OrElse
                                    text.StartsWith(" ") OrElse text.EndsWith(" ")
        If Not needsQuote Then Return text
        Return """" & text.Replace("""", """""") & """"
    End Function

    ''' <summary>Bentuk satu baris CSV (tanpa newline).</summary>
    Public Shared Function BuildLine(entry As BatchLogEntry) As String
        Dim cells As New List(Of String)()
        cells.Add(entry.Timestamp.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture))
        cells.Add(EscapeCell(entry.FileName))
        cells.Add(entry.Pages.ToString(CultureInfo.InvariantCulture))
        cells.Add(entry.Shapes.ToString(CultureInfo.InvariantCulture))
        cells.Add(EscapeCell(entry.Status))
        cells.Add(EscapeCell(entry.Message))
        cells.Add(entry.DurationSeconds.ToString("0.0", CultureInfo.InvariantCulture))
        Return String.Join(Delimiter.ToString(), cells.ToArray())
    End Function

    ''' <summary>Path log default di dalam folder batch.</summary>
    Public Shared Function CsvPathFor(folder As String) As String
        If String.IsNullOrEmpty(folder) Then Return DefaultFileName
        Return Path.Combine(folder, DefaultFileName)
    End Function

    ''' <summary>
    ''' Tambahkan satu baris ke file CSV. Membuat folder + header bila belum ada.
    ''' Mengembalikan True bila berhasil (kegagalan log tidak boleh menghentikan batch).
    ''' </summary>
    Public Shared Function AppendRow(csvPath As String, entry As BatchLogEntry) As Boolean
        If String.IsNullOrEmpty(csvPath) OrElse entry Is Nothing Then Return False
        Try
            Dim folder As String = Path.GetDirectoryName(csvPath)
            If Not String.IsNullOrEmpty(folder) Then Directory.CreateDirectory(folder)

            Dim isNew As Boolean = Not File.Exists(csvPath) OrElse New FileInfo(csvPath).Length = 0
            Dim sb As New StringBuilder()
            If isNew Then sb.AppendLine(Header())
            sb.AppendLine(BuildLine(entry))

            ' UTF-8 + BOM supaya Excel membaca encoding dengan benar.
            File.AppendAllText(csvPath, sb.ToString(), New UTF8Encoding(True))
            Return True
        Catch
            Return False
        End Try
    End Function

    ''' <summary>Baris ringkasan untuk log batch (status = RINGKASAN).</summary>
    Public Shared Function SummaryEntry(settings As String, files As Integer, ok As Integer,
                                        failed As Integer, seconds As Double) As BatchLogEntry
        Return SummaryEntry(settings, files, ok, 0, failed, seconds)
    End Function

    ''' <summary>
    ''' Baris ringkasan batch dengan status partial terpisah (Temuan D), supaya
    ''' file yang hanya berhasil sebagian tidak dihitung sebagai berhasil.
    ''' </summary>
    Public Shared Function SummaryEntry(settings As String, files As Integer, ok As Integer,
                                        partialCount As Integer, failed As Integer,
                                        seconds As Double) As BatchLogEntry
        Dim entry As New BatchLogEntry()
        entry.FileName = "(ringkasan batch)"
        entry.Pages = files
        entry.Shapes = ok
        entry.Status = "RINGKASAN"
        ' Satu sumber teks ringkasan: tidak ada perakitan pesan di dua tempat.
        entry.Message = BuildSummaryMessage(settings, ok, partialCount, failed)
        entry.DurationSeconds = seconds
        Return entry
    End Function

    ''' <summary>
    ''' Satu-satunya tempat teks ringkasan dirakit, sehingga tidak ada penulisan
    ''' ganda ke entry.Message (review lanjutan PR #2 butir 3).
    ''' Format: "&lt;pengaturan&gt; | OK n, SEBAGIAN n, GAGAL n".
    ''' </summary>
    Public Shared Function BuildSummaryMessage(settings As String, ok As Integer,
                                               partialCount As Integer, failed As Integer) As String
        Return If(settings, "") & " | OK " & ok & ", SEBAGIAN " & partialCount & ", GAGAL " & failed
    End Function
End Class
