Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  AutoSelectService.vb
'  Orkestrasi Auto Select dipindahkan dari CopierForm (Program.vb) TANPA
'  mengubah perilaku:
'    - pemindaian halaman lewat ShapeNodeAdapter + ShapeScanner (murni),
'    - snapshot kategori/cakupan lewat MatchNavigator (murni),
'    - penerapan seleksi + verifikasi jumlah lewat ActiveSelectionRange.
'
'  Batas kelas:
'    - Tidak memakai WinForms; progress lewat IProgressSink, sehingga alur
'      "pindai semua halaman" bisa diuji dengan objek COM palsu.
'    - Tidak ada fallback Delete; seleksi lama TIDAK dibersihkan bila tidak ada
'      target (aturan yang sama seperti sebelum refactor).
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics

''' <summary>Hasil pemindaian seluruh halaman (untuk laporan UI).</summary>
Public Class AutoSelectScanReport
    Public Category As ShapeCategory = ShapeCategory.AllObjects
    Public PagesScanned As Integer = 0
    Public TotalMatches As Integer = 0
    Public TotalDiagnostics As Integer = 0
    Public DurationSeconds As Double = 0
    Public ReadOnly Property PagesWithMatches As New List(Of Integer)()
    Public ReadOnly Property PerPage As New Dictionary(Of Integer, ShapeScanResult)()

    Public ReadOnly Property HasMatches As Boolean
        Get
            Return PagesWithMatches.Count > 0
        End Get
    End Property

    ''' <summary>
    ''' Baris laporan per halaman, format sama seperti versi lama:
    '''   "Halaman n: x objek (y target seleksi)" bila ada hasil,
    '''   "Halaman n: 0 objek" bila tidak ada.
    ''' </summary>
    Public Function DescribePages() As List(Of String)
        Dim lines As New List(Of String)()
        For Each pageIndex As Integer In PerPage.Keys
            Dim scan As ShapeScanResult = PerPage(pageIndex)
            If scan.MatchCount > 0 Then
                lines.Add("Halaman " & pageIndex.ToString() & ": " & scan.MatchCount.ToString() &
                          " objek (" & scan.DistinctTargetCount.ToString() & " target seleksi)")
            Else
                lines.Add("Halaman " & pageIndex.ToString() & ": 0 objek")
            End If
        Next
        lines.Sort()
        Return lines
    End Function

    Public Function DescribeDiagnostics() As String
        Dim sb As New System.Text.StringBuilder()
        For Each pageIndex As Integer In PerPage.Keys
            Dim text As String = PerPage(pageIndex).DescribeDiagnostics()
            If Not String.IsNullOrEmpty(text) Then
                sb.AppendLine("  Halaman " & pageIndex.ToString() & ":")
                sb.Append(text)
            End If
        Next
        Return sb.ToString()
    End Function
End Class

Public Class AutoSelectService

    Private ReadOnly _connection As CorelConnectionService
    Private ReadOnly _sink As IProgressSink
    Private ReadOnly _navigator As New MatchNavigator()
    Private ReadOnly _scans As New Dictionary(Of Integer, ShapeScanResult)()

    Public Sub New(connection As CorelConnectionService, Optional sink As IProgressSink = Nothing)
        If connection Is Nothing Then Throw New ArgumentNullException(NameOf(connection))
        _connection = connection
        _sink = If(sink, New NullProgressSink())
    End Sub

    ''' <summary>Snapshot navigasi hasil pemindaian terakhir.</summary>
    Public ReadOnly Property Navigator As MatchNavigator
        Get
            Return _navigator
        End Get
    End Property

    ''' <summary>Hasil pemindaian per halaman (dipakai ulang saat navigasi).</summary>
    Public ReadOnly Property Scans As Dictionary(Of Integer, ShapeScanResult)
        Get
            Return _scans
        End Get
    End Property

    ''' <summary>Batalkan sesi sebelumnya (dipakai sebelum pemindaian & saat gagal).</summary>
    Public Sub Reset()
        _navigator.Reset()
        _scans.Clear()
    End Sub

    ''' <summary>
    ''' Jumlah halaman dokumen aktif; melempar bila dokumen tidak terbaca.
    ''' (Dinamai DocumentPageCount karena 'pageCount' dipakai variabel lokal —
    ''' VB tidak membedakan huruf besar/kecil.)
    ''' </summary>
    Public Function DocumentPageCount() As Integer
        Return CInt(_connection.Document.Pages.Count)
    End Function

    ''' <summary>
    ''' Pindai satu halaman. Mengembalikan Nothing bila objek halaman tidak dapat
    ''' diakses (alasannya dilaporkan lewat sink).
    ''' </summary>
    Public Function ScanPage(pageIndex As Integer, category As ShapeCategory) As ShapeScanResult
        Dim page As Object
        Try
            page = _connection.Document.Pages(pageIndex)
        Catch ex As Exception
            _sink.Report("Halaman " & pageIndex.ToString() & " tidak dapat diakses: " & ex.Message)
            Return Nothing
        End Try

        Dim rootNodes As List(Of IShapeNode)
        Try
            rootNodes = ShapeNodeAdapter.WrapCollection(page.Shapes)
        Catch ex As Exception
            Dim failed As New ShapeScanResult()
            failed.PageIndex = pageIndex
            failed.Category = category
            failed.Diagnostics.Add("Objek halaman tidak dapat dibaca: " & ex.GetType().Name & ": " & ex.Message)
            Return failed
        End Try

        Dim options As ShapeScanOptions = ShapeScanOptions.ForCategory(category)
        options.MaxDepth = ShapeNodeAdapter.MaxDepth
        Return ShapeScanner.Scan(rootNodes, pageIndex, options)
    End Function

    ''' <summary>
    ''' Pindai SEMUA halaman untuk satu kategori, laporkan progres per halaman,
    ''' lalu simpan snapshot navigasi bila ada hasil.
    ''' </summary>
    Public Function ScanAllPages(category As ShapeCategory) As AutoSelectScanReport
        Dim timer As Stopwatch = Stopwatch.StartNew()
        Dim report As New AutoSelectScanReport()
        report.Category = category

        Dim pageCount As Integer = DocumentPageCount()
        Dim totalMatches As Integer = 0
        Dim totalDiagnostics As Integer = 0

        For pageIndex As Integer = 1 To pageCount
            _sink.Report("Memindai progres: halaman " & pageIndex.ToString() & "/" & pageCount.ToString() &
                         " | total objek cocok sementara: " & totalMatches.ToString() & ".")
            _sink.Pump()

            Dim scan As ShapeScanResult = ScanPage(pageIndex, category)
            If scan Is Nothing Then Continue For

            _scans(pageIndex) = scan
            report.PerPage(pageIndex) = scan
            report.PagesScanned += 1
            totalMatches += scan.MatchCount
            totalDiagnostics += scan.Diagnostics.Count
            If scan.MatchCount > 0 Then report.PagesWithMatches.Add(pageIndex)
        Next

        report.TotalMatches = totalMatches
        report.TotalDiagnostics = totalDiagnostics
        timer.Stop()
        report.DurationSeconds = timer.Elapsed.TotalSeconds

        If report.HasMatches Then
            ' Snapshot kategori + cakupan ikut disimpan untuk navigasi.
            _navigator.Start(category, True, report.PagesWithMatches)
        End If
        Return report
    End Function

    ''' <summary>Pindai satu halaman (cakupan "Halaman aktif") lalu siapkan navigasi.</summary>
    Public Function ScanSinglePage(pageIndex As Integer, category As ShapeCategory) As ShapeScanResult
        Dim scan As ShapeScanResult = ScanPage(pageIndex, category)
        If scan Is Nothing Then Return Nothing
        _scans(pageIndex) = scan
        If scan.MatchCount > 0 Then
            _navigator.Start(category, False, New Integer() {pageIndex})
        End If
        Return scan
    End Function

    ''' <summary>
    ''' Pilih objek pada halaman hasil yang sedang aktif di navigator.
    ''' Selalu memakai kategori SNAPSHOT, bukan pilihan ComboBox saat ini.
    ''' </summary>
    Public Function SelectPageAtCurrentMatch(scanCategory As ShapeCategory) As SelectionOutcome
        If Not _navigator.HasResults Then Return SelectionOutcome.NotRun("Belum ada hasil pemindaian.")

        Dim pageIndex As Integer = _navigator.CurrentPage
        If pageIndex <= 0 Then Return SelectionOutcome.NotRun("Posisi hasil tidak valid.")

        Dim scan As ShapeScanResult = Nothing
        Dim hasScan As Boolean = _scans.TryGetValue(pageIndex, scan)
        If Not hasScan Then
            ' Hasil lama tidak tersedia (mis. dokumen berubah): pindai ulang
            ' dengan kategori snapshot agar jumlahnya konsisten.
            scan = ScanPage(pageIndex, scanCategory)
            If scan Is Nothing Then Return SelectionOutcome.NotRun("Halaman " & pageIndex.ToString() & " tidak dapat diakses.")
            _scans(pageIndex) = scan
        End If

        Try
            _connection.Document.Pages(pageIndex).Activate()
        Catch ex As Exception
            Return SelectionOutcome.NotRun("Halaman " & pageIndex.ToString() & " tidak dapat diaktifkan: " & ex.Message)
        End Try

        Return ApplySelection(scan, pageIndex)
    End Function

    ''' <summary>
    ''' Berpindah ke halaman hasil berikutnya. Mengembalikan langkah navigasi;
    ''' UI yang memutuskan pesan untuk StaleSnapshot / AtEnd / PageMissing.
    ''' </summary>
    Public Function MoveNext(categoryNow As ShapeCategory, allPagesNow As Boolean) As NavigationStep
        Dim pageCount As Integer = DocumentPageCount()
        Dim pageExists As Func(Of Integer, Boolean) =
            Function(pageIdx As Integer) pageIdx >= 1 AndAlso pageIdx <= pageCount
        Dim navStep As NavigationStep = _navigator.MoveNext(categoryNow, allPagesNow, pageExists)
        If navStep <> NavigationStep.Moved Then Return navStep

        Dim outcome As SelectionOutcome = SelectPageAtCurrentMatch(_navigator.Category)
        LastOutcome = outcome
        Return navStep
    End Function

    ''' <summary>Hasil seleksi dari MoveNext terakhir (dibaca UI untuk peringatan).</summary>
    Public Property LastOutcome As SelectionOutcome = Nothing

    ''' <summary>
    ''' Terapkan seleksi untuk hasil pemindaian satu halaman.
    ''' Aturan keselamatan:
    '''   - Tidak ada target -> seleksi lama TIDAK disentuh.
    '''   - Tidak ada fallback Delete; kegagalan ClearSelection membatalkan operasi.
    '''   - Jumlah hasil dibaca ulang dari CorelDRAW (bukan hitungan panggilan).
    ''' </summary>
    Public Function ApplySelection(scan As ShapeScanResult, expectedPageIndex As Integer) As SelectionOutcome
        If scan Is Nothing Then Return SelectionOutcome.NotRun("Hasil pemindaian tidak tersedia.")

        Dim targets As List(Of Object) = scan.SelectionTargets()
        If targets.Count = 0 Then
            ' Penting: jangan ClearSelection saat kosong (tidak ada operasi destruktif).
            Return SelectionOutcome.NotRun("Tidak ada target seleksi pada halaman " &
                                           expectedPageIndex.ToString() & ".")
        End If

        Dim document As Object = _connection.Document
        Dim app As Object = _connection.App

        ' Tidak ada fallback Delete: kegagalan ClearSelection menghentikan
        ' operasi demi keselamatan.
        Try
            document.ClearSelection()
        Catch ex As Exception
            Throw New InvalidOperationException(
                "Seleksi lama tidak dapat dibersihkan. Tidak ada objek yang dihapus; proses dibatalkan.", ex)
        End Try

        Dim outcome As New SelectionOutcome()
        outcome.ExpectedCount = targets.Count
        outcome.MatchCount = scan.MatchCount

        For Each target As Object In targets
            Try
                target.AddToSelection()
                outcome.AddedCalls += 1
            Catch ex As Exception
                outcome.Failures.Add("AddToSelection gagal: " & ex.GetType().Name & ": " & ex.Message)
            End Try
        Next

        ' Validasi: hitung objek yang BENAR-BENAR terpilih di CorelDRAW.
        Try
            Dim selectionRange As Object = app.ActiveSelectionRange
            If selectionRange Is Nothing Then
                outcome.ActualCount = 0
            Else
                outcome.ActualCount = CInt(selectionRange.Count)
            End If
            outcome.ActualCountKnown = True
        Catch ex As Exception
            outcome.ActualCountKnown = False
            outcome.Failures.Add("Jumlah seleksi tidak dapat diverifikasi: " & ex.GetType().Name & ": " & ex.Message)
        End Try

        If outcome.ActualCountKnown AndAlso outcome.ActualCount <> outcome.ExpectedCount Then
            outcome.Failures.Add("CorelDRAW melaporkan " & outcome.ActualCount & " objek terpilih, diharapkan " &
                                 outcome.ExpectedCount & " (objek pada layer terkunci/grup tidak dapat dipilih).")
        End If

        outcome.CategoryLabel = ShapeClassifier.CategoryLabel(scan.Category)
        Return outcome
    End Function
End Class
