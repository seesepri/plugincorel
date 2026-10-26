Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  WatermarkTokens.vb
'  Substitusi token pada teks watermark, dievaluasi PER HALAMAN.
'
'  Token yang didukung (case-insensitive, ditulis dengan kurung kurawal):
'    {page}       nomor halaman yang sedang diproses
'    {pages}      jumlah halaman dokumen
'    {file}       nama file dokumen tanpa ekstensi
'    {filename}   nama file dokumen lengkap dengan ekstensi
'    {order}      nomor order yang diekstrak dari nama file (lihat ExtractOrder)
'    {date}       tanggal hari ini (yyyy-MM-dd)
'    {time}       jam sekarang (HH:mm)
'
'  Token yang tidak dikenal dibiarkan apa adanya, sehingga teks biasa yang
'  kebetulan memakai kurung kurawal tidak rusak.
'  seluruh logika di file ini murni -> bisa diuji unit tanpa CorelDRAW.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Text

''' <summary>Data yang dipakai saat mengevaluasi token untuk satu halaman.</summary>
Public Class WatermarkTokenContext
    Public PageIndex As Integer = 1
    Public PageCount As Integer = 1
    Public FileName As String = ""
    Public Now As DateTime = DateTime.Now

    ''' <summary>Nomor order hasil ekstraksi dari FileName (kosong = tidak ada).</summary>
    Public Function Order() As String
        Return WatermarkTokens.ExtractOrder(FileName)
    End Function
End Class

Public Class WatermarkTokens

    ''' <summary>Daftar token untuk ditampilkan di UI / dokumentasi.</summary>
    Public Shared Function SupportedTokens() As String()
        Return New String() {"{page}", "{pages}", "{file}", "{filename}", "{order}", "{date}", "{time}"}
    End Function

    ''' <summary>Teks bantuan singkat untuk dialog.</summary>
    Public Shared Function HelpText() As String
        Return "Token: {page} {pages} {file} {order} {date} {time} — contoh: ORDER {order} · Hal. {page}/{pages}"
    End Function

    ''' <summary>True bila template memakai token tertentu (mis. "page").</summary>
    Public Shared Function HasToken(template As String, token As String) As Boolean
        If String.IsNullOrEmpty(template) OrElse String.IsNullOrEmpty(token) Then Return False
        Dim parser As New TokenScanner(template)
        While parser.MoveNext()
            If String.Equals(parser.Token, token, StringComparison.OrdinalIgnoreCase) Then Return True
        End While
        Return False
    End Function

    ''' <summary>Ganti semua token yang dikenal dengan nilainya. Token asing dibiarkan.</summary>
    Public Shared Function Expand(template As String, context As WatermarkTokenContext) As String
        If String.IsNullOrEmpty(template) Then Return ""
        Dim ctx As WatermarkTokenContext = If(context, New WatermarkTokenContext())

        Dim sb As New StringBuilder()
        Dim scanner As New TokenScanner(template)
        Dim lastIndex As Integer = 0
        While scanner.MoveNext()
            sb.Append(template.Substring(lastIndex, scanner.StartIndex - lastIndex))
            Dim replacement As String = Resolve(scanner.Token, ctx)
            If replacement Is Nothing Then
                ' Token tidak dikenal: pertahankan tulisan aslinya.
                sb.Append(template.Substring(scanner.StartIndex, scanner.EndIndex - scanner.StartIndex))
            Else
                sb.Append(replacement)
            End If
            lastIndex = scanner.EndIndex
        End While
        If lastIndex < template.Length Then sb.Append(template.Substring(lastIndex))
        Return sb.ToString()
    End Function

    Private Shared Function Resolve(token As String, ctx As WatermarkTokenContext) As String
        Select Case token.ToLowerInvariant()
            Case "page" : Return ctx.PageIndex.ToString(CultureInfo.InvariantCulture)
            Case "pages" : Return ctx.PageCount.ToString(CultureInfo.InvariantCulture)
            Case "file" : Return FileBaseName(ctx.FileName)
            Case "filename" : Return FileNameOnly(ctx.FileName)
            Case "order" : Return ctx.Order()
            Case "date" : Return ctx.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            Case "time" : Return ctx.Now.ToString("HH:mm", CultureInfo.InvariantCulture)
            Case Else : Return Nothing
        End Select
    End Function

    ''' <summary>
    ''' Nama file tanpa folder & tanpa ekstensi. Ditangani manual (bukan
    ''' Path.GetFileNameWithoutExtension) supaya pemisah "\" maupun "/" dikenali
    ''' di OS mana pun -> perilaku aplikasi dan unit test konsisten.
    ''' </summary>
    Public Shared Function FileBaseName(filePath As String) As String
        Return StripExtension(FileNameOnly(filePath))
    End Function

    ''' <summary>Nama file tanpa folder (masih dengan ekstensi).</summary>
    Public Shared Function FileNameOnly(filePath As String) As String
        If String.IsNullOrEmpty(filePath) Then Return ""
        Dim normalized As String = filePath.Replace(Convert.ToChar(92), "/"c)
        Dim slash As Integer = normalized.LastIndexOf("/"c)
        If slash >= 0 Then normalized = normalized.Substring(slash + 1)
        Return normalized.Trim()
    End Function

    Private Shared Function StripExtension(fileName As String) As String
        If String.IsNullOrEmpty(fileName) Then Return ""
        Dim dot As Integer = fileName.LastIndexOf("."c)
        If dot > 0 Then Return fileName.Substring(0, dot)
        Return fileName
    End Function

    ''' <summary>
    ''' Ambil "nomor order" dari nama file.
    '''
    ''' Aturan (deterministik, tanpa regex):
    '''   1. Buang folder & ekstensi; hanya SPASI yang dianggap pemisah
    '''      (tanda hubung & garis bawah dipertahankan, karena nomor order
    '''      biasanya berbentuk INV-2026-0912 atau PO_2026_0042).
    '''   2. Bila nama tidak memuat spasi: seluruh nama dipakai asalkan memuat
    '''      minimal 4 digit. "PO_2026_0042.cdr" -> "PO_2026_0042".
    '''   3. Bila ada spasi: ambil potongan pertama yang memuat minimal 4 digit.
    '''      "INV-2026-0912 klien budi.cdr" -> "INV-2026-0912".
    '''   4. Bila tidak ada angka yang meyakinkan: pakai nama file utuh.
    '''      "desain final.cdr" -> "desain final".
    ''' </summary>
    Public Shared Function ExtractOrder(fileName As String) As String
        If String.IsNullOrEmpty(fileName) Then Return ""
        Dim baseName As String = FileBaseName(fileName)
        If baseName.Length = 0 Then Return ""

        Dim cleanedBase As String = baseName.Trim("-"c, "."c)
        Dim candidates As String() = baseName.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)

        If candidates.Length <= 1 Then
            If DigitCount(cleanedBase) >= 4 Then Return cleanedBase
            Return baseName
        End If

        For Each candidate As String In candidates
            Dim cleaned As String = candidate.Trim("-"c, "."c)
            If cleaned.Length > 0 AndAlso DigitCount(cleaned) >= 4 Then Return cleaned
        Next

        Return baseName
    End Function

    Private Shared Function DigitCount(value As String) As Integer
        Dim count As Integer = 0
        For Each ch As Char In value
            If Char.IsDigit(ch) Then count += 1
        Next
        Return count
    End Function

    ''' <summary>Pemindai token sederhana: menemukan pasangan { ... } berurutan.</summary>
    Private Class TokenScanner
        Private ReadOnly _source As String
        Private _position As Integer = 0

        Public StartIndex As Integer = -1
        Public EndIndex As Integer = -1
        Public Token As String = ""

        Public Sub New(source As String)
            _source = If(source, "")
        End Sub

        Public Function MoveNext() As Boolean
            StartIndex = -1 : EndIndex = -1 : Token = ""
            If _source Is Nothing Then Return False
            While _position < _source.Length
                Dim open As Integer = _source.IndexOf("{"c, _position)
                If open < 0 Then Return False
                Dim close As Integer = _source.IndexOf("}"c, open + 1)
                If close < 0 Then Return False
                Dim inner As String = _source.Substring(open + 1, close - open - 1).Trim()
                _position = close + 1
                If inner.Length = 0 Then Continue While      ' "{}" dilewati
                StartIndex = open
                EndIndex = close + 1
                Token = inner
                Return True
            End While
            Return False
        End Function
    End Class
End Class
