Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  WatermarkEngine.vb
'  Inti fitur watermark: membuat / mengganti / menghapus watermark pada
'  dokumen CorelDRAW melalui COM (late binding, sama seperti Program.vb).
'
'  Semua panggilan COM dibungkus Try/Catch per-langkah supaya satu API yang
'  tidak tersedia di versi CorelDRAW tertentu tidak menggagalkan seluruh proses.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Text

''' <summary>Konstanta CorelDRAW yang dipakai (nilainya stabil lintas versi).</summary>
Public Module WatermarkCdrConst
    Public Const CdrMillimeter As Integer = 3             ' cdrUnit.cdrMillimeter
    Public Const CdrApplyToFillAndOutline As Integer = 2  ' cdrTransparencyAppliedTo
    Public Const CdrMergeNormal As Integer = 0            ' cdrMergeMode.cdrMergeNormal
    Public Const CdrGroupShape As Integer = 7             ' cdrShapeType.cdrGroupShape
    Public Const CdrTextShape As Integer = 6              ' cdrShapeType.cdrTextShape
End Module

''' <summary>Hasil satu operasi watermark (untuk dialog laporan + log CSV).</summary>
Public Class WatermarkResult
    Public PagesPassed As Integer = 0
    Public ShapesCreated As Integer = 0
    Public DocumentsTouched As Integer = 0
    ''' <summary>Jumlah halaman target yang seharusnya diproses (Temuan D).</summary>
    Public PagesIntended As Integer = 0
    ''' <summary>Jumlah halaman target yang GAGAL diproses (Temuan D).</summary>
    Public PagesFailed As Integer = 0
    Public Failures As New List(Of String)()
    Public Errors As New List(Of String)()
    ''' <summary>
    ''' Masalah yang tidak menggagalkan output tetapi wajib dilaporkan
    ''' (mis. grup undo tidak dapat dibuka). Temuan C.
    ''' </summary>
    Public Warnings As New List(Of String)()
    Public Log As New List(Of String)()
    Public ProofPdfFiles As New List(Of String)()
    Public LogRowsWritten As Integer = 0
    Public DurationSeconds As Double = 0
    ''' <summary>Hasil per file pada mode batch — SATU sumber status untuk CSV/ringkasan/UI.</summary>
    Public FileOutcomes As New List(Of BatchFileOutcome)()
    ''' <summary>File yang sengaja dilewati saat pengumpulan sumber batch (Temuan E).</summary>
    Public SkippedFiles As New List(Of SkippedBatchFile)()
    ''' <summary>Folder hasil terakhir (untuk tombol "Buka folder hasil").</summary>
    Public OutputFolder As String = ""
    ''' <summary>File hasil terakhir (dipakai untuk explorer /select).</summary>
    Public OutputFile As String = ""

    Public ReadOnly Property FilesOk As Integer
        Get
            Return BatchStatusRules.CountStatus(FileOutcomes, BatchFileStatus.Ok)
        End Get
    End Property

    Public ReadOnly Property FilesPartial As Integer
        Get
            Return BatchStatusRules.CountStatus(FileOutcomes, BatchFileStatus.PartialSuccess)
        End Get
    End Property

    Public ReadOnly Property FilesFailed As Integer
        Get
            Return BatchStatusRules.CountStatus(FileOutcomes, BatchFileStatus.Failed)
        End Get
    End Property

    Public ReadOnly Property CanOpenOutputFolder As Boolean
        Get
            If String.IsNullOrEmpty(OutputFolder) Then Return False
            Try
                Return Directory.Exists(OutputFolder)
            Catch
                Return False
            End Try
        End Get
    End Property

    Public ReadOnly Property Ok As Boolean
        Get
            Return Errors.Count = 0 AndAlso Failures.Count = 0 AndAlso FilesFailed = 0
        End Get
    End Property

    Public Sub Merge(other As WatermarkResult)
        If other Is Nothing Then Return
        PagesPassed += other.PagesPassed
        ShapesCreated += other.ShapesCreated
        DocumentsTouched += other.DocumentsTouched
        PagesIntended += other.PagesIntended
        PagesFailed += other.PagesFailed
        Failures.AddRange(other.Failures)
        Errors.AddRange(other.Errors)
        Warnings.AddRange(other.Warnings)
        Log.AddRange(other.Log)
        ProofPdfFiles.AddRange(other.ProofPdfFiles)
        FileOutcomes.AddRange(other.FileOutcomes)
        SkippedFiles.AddRange(other.SkippedFiles)
        LogRowsWritten += other.LogRowsWritten
        If String.IsNullOrEmpty(OutputFolder) AndAlso Not String.IsNullOrEmpty(other.OutputFolder) Then
            OutputFolder = other.OutputFolder
        End If
        If String.IsNullOrEmpty(OutputFile) AndAlso Not String.IsNullOrEmpty(other.OutputFile) Then
            OutputFile = other.OutputFile
        End If
    End Sub

    Public Function ToReport() As String
        Dim sb As New StringBuilder()
        If Errors.Count > 0 Then
            sb.AppendLine("GAGAL — kesalahan:")
            For Each e In Errors
                sb.AppendLine("  • " & e)
            Next
            sb.AppendLine()
        End If
        sb.AppendLine("Halaman diproses : " & PagesPassed &
                      If(PagesIntended > 0, " / " & PagesIntended & " target", ""))
        If PagesFailed > 0 Then sb.AppendLine("Halaman gagal    : " & PagesFailed)
        sb.AppendLine("Shape dibuat     : " & ShapesCreated)
        If DocumentsTouched > 0 Then sb.AppendLine("Dokumen tersimpan: " & DocumentsTouched)
        If ProofPdfFiles.Count > 0 Then
            sb.AppendLine("PDF proof        : " & ProofPdfFiles.Count & " file")
            For Each pdf As String In ProofPdfFiles
                sb.AppendLine("  - " & pdf)
            Next
        End If
        If LogRowsWritten > 0 Then sb.AppendLine("Baris log CSV    : " & LogRowsWritten)
        If OutputFolder IsNot Nothing AndAlso OutputFolder.Length > 0 Then
            sb.AppendLine("Folder hasil     : " & OutputFolder)
        End If

        ' --- Hasil per file (mode batch): hanya ada satu sumber status -------
        If FileOutcomes.Count > 0 Then
            sb.AppendLine()
            sb.AppendLine("Hasil per file   : " & FilesOk & " OK, " & FilesPartial & " SEBAGIAN, " & FilesFailed & " GAGAL")
            For Each outcome As BatchFileOutcome In FileOutcomes
                sb.AppendLine("  " & outcome.ToReportLine())
            Next
        End If

        If SkippedFiles.Count > 0 Then
            sb.AppendLine()
            sb.AppendLine("File dilewati    : " & SkippedFiles.Count)
            For Each skipped As SkippedBatchFile In SkippedFiles
                sb.AppendLine("  - " & Path.GetFileName(skipped.Path) & ": " & skipped.Reason)
            Next
        End If

        sb.AppendLine("Kegagalan        : " & Failures.Count)
        If Failures.Count > 0 Then
            sb.AppendLine()
            For Each f In Failures
                sb.AppendLine("  • " & f)
            Next
        End If
        If Warnings.Count > 0 Then
            sb.AppendLine()
            sb.AppendLine("Peringatan:")
            For Each w In Warnings
                sb.AppendLine("  • " & w)
            Next
        End If
        If Log.Count > 0 Then
            sb.AppendLine()
            sb.AppendLine("Log:")
            For Each l In Log
                sb.AppendLine("  " & l)
            Next
        End If
        Return sb.ToString()
    End Function
End Class

''' <summary>Mesin watermark. Satu instance bisa dipakai ulang untuk banyak dokumen.</summary>
Public Class WatermarkEngine

    ''' <summary>Dipanggil untuk setiap langkah penting (untuk status bar).</summary>
    Public Event Progress(message As String)

    Private ReadOnly _result As New WatermarkResult()

    Public ReadOnly Property Result As WatermarkResult
        Get
            Return _result
        End Get
    End Property

    Private Sub Report(message As String)
        _result.Log.Add(message)
        RaiseEvent Progress(message)
    End Sub

    ' ========================================================================
    '  OPERASI UTAMA: SATU DOKUMEN
    ' ========================================================================

    ''' <summary>
    ''' Terapkan watermark ke dokumen aktif CorelDRAW.
    ''' doc = objek Document (late bound). Tidak menutup/menyimpan dokumen.
    '''
    ''' Audit reliabilitas:
    '''   B. Satuan: Document.Unit dibaca dulu. Bila tidak terbaca, atau tidak bisa
    '''      diubah ke mm, operasi DIBATALKAN sebelum dokumen diubah (perhitungan
    '''      tata letak memakai mm, jadi asumsi mm tanpa normalisasi = posisi salah).
    '''      Pemulihan satuan dilakukan di Finally dan kegagalannya dilaporkan.
    '''   C. Undo: kegagalan BeginCommandGroup/EndCommandGroup dicatat ke
    '''      Warnings/Errors, tidak lagi ditelan diam-diam.
    '''   D. PagesIntended/PagesFailed diisi supaya batch bisa menilai status.
    ''' </summary>
    Public Function ApplyToDocument(doc As Object, settings As WatermarkSettings,
                                    Optional writeLogRow As Boolean = True) As WatermarkResult
        Dim localResult As New WatermarkResult()
        Dim documentTimer As Stopwatch = Stopwatch.StartNew()

        If settings Is Nothing Then
            localResult.Errors.Add("Pengaturan watermark tidak tersedia.")
            _result.Merge(localResult)
            Return localResult
        End If

        Dim validation As List(Of String) = settings.Validate()
        If validation.Count > 0 Then
            localResult.Errors.AddRange(validation)
            _result.Merge(localResult)
            Return localResult
        End If

        If doc Is Nothing Then
            localResult.Errors.Add("Dokumen CorelDRAW tidak tersedia.")
            _result.Merge(localResult)
            Return localResult
        End If

        Dim unitScope As New DocumentUnitScope(WatermarkCdrConst.CdrMillimeter)
        Dim undoGuard As New UndoGroupGuard("Watermark - Corel Page Copier")
        Dim oldPageIndex As Integer = -1

        Try
            ' --- Halaman aktif (dipulihkan di Finally) -------------------------
            Try
                oldPageIndex = CInt(doc.ActivePage.Index)
            Catch ex As Exception
                localResult.Warnings.Add("Halaman aktif tidak dapat dibaca: " & ex.Message)
            End Try

            ' --- Temuan B: satuan dokumen ------------------------------------
            Try
                unitScope.MarkOriginalKnown(CInt(doc.Unit))
            Catch ex As Exception
                unitScope.MarkOriginalUnknown(ex.Message)
            End Try
            If Not unitScope.OriginalUnitKnown Then
                localResult.Errors.Add(unitScope.BlockingReason)
                Return localResult
            End If

            Try
                doc.Unit = unitScope.TargetUnit
                unitScope.MarkChangeApplied()
            Catch ex As Exception
                unitScope.MarkChangeFailed(ex.Message)
            End Try
            If Not unitScope.ChangeApplied Then
                ' Tata letak watermark dihitung dalam mm; tanpa normalisasi hasil
                ' akan salah tetapi terlihat "sukses". Batalkan sebelum menyentuh dokumen.
                localResult.Errors.Add(
                    "Satuan dokumen tidak dapat diubah ke mm (" & unitScope.ChangeError &
                    "). Watermark dibatalkan agar ukuran/posisi tidak salah.")
                Return localResult
            End If

            ' --- Tentukan halaman target --------------------------------------
            Dim pages As Object = doc.Pages
            Dim pageCount As Integer = CInt(pages.Count)
            Dim targets As New List(Of Integer)()

            If settings.AllPages Then
                For i As Integer = 1 To pageCount
                    targets.Add(i)
                Next
            Else
                For Each p As Integer In settings.TargetPages
                    If p >= 1 AndAlso p <= pageCount AndAlso Not targets.Contains(p) Then targets.Add(p)
                Next
            End If

            If targets.Count = 0 Then
                localResult.Errors.Add("Tidak ada halaman target yang valid pada dokumen ini.")
                Return localResult
            End If
            localResult.PagesIntended = targets.Count

            ' Konteks token: nama file & jumlah halaman diketahui sekali per dokumen.
            Dim tokenContext As New WatermarkTokenContext()
            tokenContext.FileName = SafeDocumentName(doc)
            tokenContext.PageCount = pageCount
            tokenContext.Now = DateTime.Now

            Report("Mulai watermark: " & settings.Describe())

            ' --- Temuan C: satu grup perintah (tidak lagi ditelan) ------------
            undoGuard.MarkBeginAttempted()
            Try
                doc.BeginCommandGroup(undoGuard.Name)
                undoGuard.MarkBeginSucceeded()
            Catch ex As Exception
                undoGuard.MarkBeginFailed(ex.Message)
                localResult.Warnings.Add("Grup undo tidak dapat dibuka: " & ex.Message &
                                         ". Ctrl+Z harus diulang per halaman.")
            End Try

            ' --- Proses tiap halaman ------------------------------------------
            For pos As Integer = 0 To targets.Count - 1
                Dim pageIndex As Integer = targets(pos)
                RaiseEvent Progress("Watermark " & (pos + 1) & "/" & targets.Count & " — halaman " & pageIndex & "…")
                Try
                    Dim created As Integer = ApplyToPage(doc, pages(pageIndex), pageIndex, settings, tokenContext)
                    localResult.ShapesCreated += created
                    localResult.PagesPassed += 1
                    Report("Halaman " & pageIndex & ": " & created & " shape")
                Catch ex As Exception
                    localResult.PagesFailed += 1
                    localResult.Failures.Add("Halaman " & pageIndex & ": " & ex.GetType().Name & ": " & ex.Message)
                End Try
            Next

            If localResult.PagesPassed > 0 Then localResult.DocumentsTouched = 1

        Catch ex As Exception
            localResult.Errors.Add(ex.GetType().Name & ": " & ex.Message)
        Finally
            ' --- Temuan C: EndCommandGroup hanya bila group benar-benar terbuka ---
            If undoGuard.ShouldCallEnd Then
                Try
                    doc.EndCommandGroup()
                    undoGuard.MarkEndSucceeded()
                Catch ex As Exception
                    undoGuard.MarkEndFailed(ex.Message)
                    localResult.Errors.Add("EndCommandGroup gagal: " & ex.Message &
                                           ". Grup undo mungkin masih terbuka — periksa daftar Undo CorelDRAW.")
                End Try
            End If

            ' --- Temuan B: pulihkan satuan, laporkan bila gagal --------------
            If unitScope.RestoreRequired Then
                Try
                    doc.Unit = unitScope.OriginalUnit
                    unitScope.MarkRestored()
                Catch ex As Exception
                    unitScope.MarkRestoreFailed(ex.Message)
                    localResult.Errors.Add("Pemulihan satuan dokumen ke " &
                                           DocumentUnits.UnitName(unitScope.OriginalUnit) &
                                           " GAGAL: " & ex.Message & ". Satuan dokumen kemungkinan tetap mm.")
                End Try
            End If

            If oldPageIndex > 0 Then
                Try
                    doc.Pages(oldPageIndex).Activate()
                Catch ex As Exception
                    localResult.Warnings.Add("Halaman " & oldPageIndex & " tidak dapat diaktifkan kembali: " & ex.Message)
                End Try
            End If
            Try
                doc.Application.Refresh()
            Catch
                ' Refresh hanya kosmetik; kegagalan tidak memengaruhi hasil.
            End Try
        End Try

        documentTimer.Stop()
        localResult.DurationSeconds = documentTimer.Elapsed.TotalSeconds

        Try
            Dim docPath As String = CStr(doc.FilePath)
            If Not String.IsNullOrEmpty(docPath) Then
                localResult.OutputFolder = Path.GetDirectoryName(docPath)
                localResult.OutputFile = docPath
            End If
        Catch
        End Try

        ' --- Output tambahan (Fase 2) -------------------------------------
        If localResult.PagesPassed > 0 Then
            If settings.ExportProofPdf Then
                Dim pdfPath As String = TryPublishProofPdf(doc, settings, localResult)
                If pdfPath IsNot Nothing Then
                    localResult.ProofPdfFiles.Add(pdfPath)
                    localResult.OutputFolder = Path.GetDirectoryName(pdfPath)
                    localResult.OutputFile = pdfPath
                End If
            End If
            If settings.WriteCsvLog AndAlso writeLogRow Then
                Dim entry As New BatchLogEntry()
                entry.FileName = SafeDocumentName(doc)
                entry.Pages = localResult.PagesPassed
                entry.Shapes = localResult.ShapesCreated
                ' Status CSV memakai aturan yang sama dengan batch (Temuan D).
                entry.Status = BatchStatusRules.Label(
                    BatchStatusRules.EvaluateApplyOnly(True, localResult.PagesIntended,
                                                       localResult.PagesPassed, localResult.PagesFailed,
                                                       localResult.Errors.Count > 0))
                entry.Message = settings.Describe()
                entry.DurationSeconds = localResult.DurationSeconds
                If BatchLog.AppendRow(ResolveCsvPath(settings, Nothing), entry) Then
                    localResult.LogRowsWritten += 1
                End If
            End If
        End If

        _result.Merge(localResult)
        Return localResult
    End Function

    ' ========================================================================
    '  OUTPUT: PDF PROOF
    ' ========================================================================

    ''' <summary>
    ''' Ekspor PDF proof ber-watermark ke samping file dokumen (<nama>_proof.pdf).
    ''' Mengembalikan path hasil, atau Nothing bila gagal (dicatat sebagai kegagalan).
    ''' </summary>
    Private Function TryPublishProofPdf(doc As Object, settings As WatermarkSettings,
                                        localResult As WatermarkResult) As String
        Dim sourcePath As String = ""
        Try
            sourcePath = CStr(doc.FilePath)
        Catch
        End Try

        Dim folder As String
        Dim baseName As String
        If String.IsNullOrEmpty(sourcePath) Then
            ' Dokumen belum pernah disimpan: taruh di folder temp agar tetap ada hasilnya.
            folder = Path.Combine(Path.GetTempPath(), "CorelPageCopier")
            baseName = "dokumen"
        Else
            folder = Path.GetDirectoryName(sourcePath)
            baseName = Path.GetFileNameWithoutExtension(sourcePath)
        End If

        Dim suffix As String = If(String.IsNullOrEmpty(settings.ProofPdfSuffix), "_proof", settings.ProofPdfSuffix)
        Dim pdfPath As String = Path.Combine(folder, baseName & suffix & ".pdf")

        Try
            Directory.CreateDirectory(folder)
            doc.PublishToPDF(pdfPath)
            Report("PDF proof: " & pdfPath)
            Return pdfPath
        Catch ex As Exception
            ' Fallback: ExportEx dengan filter PDF (beberapa versi butuh ini).
            Try
                doc.ExportEx(pdfPath, 2, 0)   ' 2 = cdrFilterPDF
                Report("PDF proof (ExportEx): " & pdfPath)
                Return pdfPath
            Catch
                localResult.Failures.Add("PDF proof gagal dibuat (" & Path.GetFileName(pdfPath) & "): " & ex.Message)
                Return Nothing
            End Try
        End Try
    End Function

    Private Function SafeDocumentName(doc As Object) As String
        Try
            Dim p As String = CStr(doc.FilePath)
            If Not String.IsNullOrEmpty(p) Then Return Path.GetFileName(p)
        Catch
        End Try
        Return "(dokumen aktif)"
    End Function

    ''' <summary>Path CSV: dari setting, atau watermark-log.csv di folder file/dokumen.</summary>
    Private Function ResolveCsvPath(settings As WatermarkSettings, sourceFile As String) As String
        If settings IsNot Nothing AndAlso Not String.IsNullOrEmpty(settings.CsvLogPath) Then
            Return settings.CsvLogPath
        End If
        Dim folder As String = ""
        If Not String.IsNullOrEmpty(sourceFile) Then
            folder = Path.GetDirectoryName(sourceFile)
        Else
            folder = Path.Combine(Path.GetTempPath(), "CorelPageCopier")
        End If
        Return BatchLog.CsvPathFor(folder)
    End Function

    ' ========================================================================
    '  OPERASI: SATU HALAMAN
    ' ========================================================================

    Private Function ApplyToPage(doc As Object, page As Object, pageIndex As Integer,
                                 settings As WatermarkSettings, context As WatermarkTokenContext) As Integer
        Dim layer As Object = EnsureLayer(page, settings.LayerName)

        If settings.ReplaceExisting Then RemoveWatermarkShapes(page)

        Dim pageWidth As Double = CDbl(page.SizeWidth)
        Dim pageHeight As Double = CDbl(page.SizeHeight)
        Dim created As Integer = 0

        ' --- 1. Objek teks ---------------------------------------------------
        If settings.Kind = WatermarkKind.TextOnly OrElse settings.Kind = WatermarkKind.TextAndLogo Then
            ' Token dievaluasi per halaman ({page}, {order}, {file}, ...).
            If context IsNot Nothing Then context.PageIndex = pageIndex
            Dim body As String = WatermarkTokens.Expand(settings.Text, context)
            ' Perilaku lama tetap dijaga: tambahkan "Hal. n" bila diminta dan
            ' template belum memakai token {page}.
            If settings.PerPageText AndAlso Not WatermarkTokens.HasToken(settings.Text, "page") Then
                body = body & " · Hal. " & pageIndex
            End If

            Dim textShape As Object = layer.CreateArtisticText(
                0, 0, body, 0, -1, settings.FontName, settings.FontSize)

            Try
                textShape.Text.Story.Font = settings.FontName
                textShape.Text.Story.Size = settings.FontSize
            Catch
            End Try
            Try
                textShape.Outline.SetNoOutline()
            Catch
            End Try
            Try
                textShape.Fill.UniformColor.RGBAssign(settings.Color.R, settings.Color.G, settings.Color.B)
            Catch
            End Try

            textShape.Name = WatermarkLayout.ShapePrefix & "TEXT"
            StyleObject(textShape, settings)
            created += PlaceObject(textShape, pageWidth, pageHeight, 0, 0, settings, "TEXT")
        End If

        ' --- 2. Objek logo ---------------------------------------------------
        If settings.Kind = WatermarkKind.LogoOnly OrElse settings.Kind = WatermarkKind.TextAndLogo Then
            Dim logoShape As Object = ImportLogo(layer, settings.LogoPath)
            logoShape.Name = WatermarkLayout.ShapePrefix & "LOGO"
            StyleObject(logoShape, settings)
            created += PlaceObject(logoShape, pageWidth, pageHeight,
                                   settings.LogoWidthMm, 0, settings, "LOGO")
        End If

        If settings.LockLayer Then
            Try
                layer.Editable = False
            Catch
            End Try
        End If

        If created = 0 Then
            Throw New InvalidOperationException("Tidak ada objek watermark yang berhasil dibuat.")
        End If

        Return created
    End Function

    ' ========================================================================
    '  OPERASI: HAPUS WATERMARK
    ' ========================================================================

    ''' <summary>
    ''' Hapus semua shape watermark (nama berawalan WM_) pada semua halaman.
    ''' Temuan C: kegagalan grup undo dicatat, EndCommandGroup hanya dipanggil
    ''' bila group benar-benar terbuka, dan kegagalan per halaman dihitung.
    ''' </summary>
    Public Function RemoveFromDocument(doc As Object) As WatermarkResult
        Dim localResult As New WatermarkResult()
        If doc Is Nothing Then
            localResult.Errors.Add("Dokumen CorelDRAW tidak tersedia.")
            _result.Merge(localResult)
            Return localResult
        End If

        Dim undoGuard As New UndoGroupGuard("Hapus Watermark - Corel Page Copier")
        Try
            Dim pageCount As Integer = CInt(doc.Pages.Count)
            localResult.PagesIntended = pageCount

            undoGuard.MarkBeginAttempted()
            Try
                doc.BeginCommandGroup(undoGuard.Name)
                undoGuard.MarkBeginSucceeded()
            Catch ex As Exception
                undoGuard.MarkBeginFailed(ex.Message)
                localResult.Warnings.Add("Grup undo tidak dapat dibuka: " & ex.Message &
                                         ". Penghapusan tidak dapat dibatalkan dengan satu Ctrl+Z.")
            End Try

            For i As Integer = 1 To pageCount
                Try
                    Dim removed As Integer = RemoveWatermarkShapes(doc.Pages(i))
                    If removed > 0 Then
                        localResult.ShapesCreated += removed
                        localResult.PagesPassed += 1
                        Report("Halaman " & i & ": " & removed & " watermark dihapus")
                    End If
                Catch ex As Exception
                    localResult.PagesFailed += 1
                    localResult.Failures.Add("Halaman " & i & ": " & ex.GetType().Name & ": " & ex.Message)
                End Try
            Next

            If localResult.PagesPassed > 0 Then localResult.DocumentsTouched = 1
        Catch ex As Exception
            localResult.Errors.Add(ex.GetType().Name & ": " & ex.Message)
        Finally
            If undoGuard.ShouldCallEnd Then
                Try
                    doc.EndCommandGroup()
                    undoGuard.MarkEndSucceeded()
                Catch ex As Exception
                    undoGuard.MarkEndFailed(ex.Message)
                    localResult.Errors.Add("EndCommandGroup gagal: " & ex.Message &
                                           ". Grup undo mungkin masih terbuka — periksa daftar Undo CorelDRAW.")
                End Try
            End If
        End Try

        _result.Merge(localResult)
        Return localResult
    End Function

    ' ========================================================================
    '  OPERASI: BATCH FILE .CDR DALAM FOLDER
    ' ========================================================================

    ''' <summary>
    ''' Terapkan watermark ke semua file .cdr dalam sebuah folder.
    ''' saveInPlace = False (default) => hasil disimpan sebagai "<nama>_wm.cdr"
    ''' sehingga file asli tidak pernah tertimpa.
    '''
    ''' Audit reliabilitas:
    '''   D. Status per file dipisah: processing (halaman), save, dan final result.
    '''      Status OK hanya bila watermark diterapkan ke SEMUA halaman target DAN
    '''      output tersimpan & terverifikasi. CSV, ringkasan, dan UI memakai
    '''      BatchFileOutcome.Status yang sama.
    '''   E. Output batch lama (*_wm.cdr) tidak lagi dijadikan sumber, benturan
    '''      nama diperiksa, dan mode file baru tidak pernah menimpa file sumber.
    '''      Mode timpa divalidasi ketat per file.
    ''' </summary>
    Public Function ApplyToFolder(app As Object, folder As String, settings As WatermarkSettings,
                                  includeSubfolders As Boolean, saveInPlace As Boolean) As WatermarkResult
        Dim localResult As New WatermarkResult()

        If app Is Nothing Then
            localResult.Errors.Add("Aplikasi CorelDRAW tidak tersedia.")
            _result.Merge(localResult)
            Return localResult
        End If
        If settings Is Nothing Then
            localResult.Errors.Add("Pengaturan watermark tidak tersedia.")
            _result.Merge(localResult)
            Return localResult
        End If
        If String.IsNullOrEmpty(folder) OrElse Not Directory.Exists(folder) Then
            localResult.Errors.Add("Folder tidak ditemukan: " & If(folder, ""))
            _result.Merge(localResult)
            Return localResult
        End If

        Dim suffix As String = BatchFilePolicy.NormalizeOutputSuffix(settings.OutputSuffix)

        Dim candidates As New List(Of String)()
        Try
            Dim option_ As SearchOption = If(includeSubfolders, SearchOption.AllDirectories, SearchOption.TopDirectoryOnly)
            candidates.AddRange(Directory.GetFiles(folder, "*" & BatchFilePolicy.SourceExtension, option_))
        Catch ex As Exception
            localResult.Errors.Add("Gagal membaca folder: " & ex.Message)
            _result.Merge(localResult)
            Return localResult
        End Try

        If candidates.Count = 0 Then
            localResult.Errors.Add("Tidak ada file " & BatchFilePolicy.SourceExtension & " di folder tersebut.")
            _result.Merge(localResult)
            Return localResult
        End If

        ' --- Temuan E: keluarkan output batch lama dari daftar sumber --------
        Dim selection As BatchSourceSelection = BatchFilePolicy.SelectSources(candidates, suffix)
        localResult.SkippedFiles.AddRange(selection.Skipped)
        For Each skipped As SkippedBatchFile In selection.Skipped
            Report("Dilewati: " & Path.GetFileName(skipped.Path) & " — " & skipped.Reason)
        Next

        If selection.Sources.Count = 0 Then
            localResult.Errors.Add(
                "Tidak ada file yang bisa diproses: " & candidates.Count & " file ditemukan, " &
                selection.Skipped.Count & " dilewati (lihat daftar file dilewati).")
            _result.Merge(localResult)
            Return localResult
        End If

        Report("Batch: " & selection.Sources.Count & " file " & BatchFilePolicy.SourceExtension &
               " diproses dari " & folder &
               If(selection.Skipped.Count > 0, " (" & selection.Skipped.Count & " dilewati)", ""))

        Dim batchTimer As Stopwatch = Stopwatch.StartNew()
        Dim csvPath As String = ResolveCsvPath(settings, Path.Combine(folder, "x" & BatchFilePolicy.SourceExtension))

        ' Path semua sumber: dipakai agar mode "file baru" tidak menimpa sumber lain.
        Dim sourcePaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each sourceFile As String In selection.Sources
            Try
                sourcePaths.Add(Path.GetFullPath(sourceFile))
            Catch
                sourcePaths.Add(sourceFile)
            End Try
        Next
        Dim reservedOutputs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        Dim index As Integer = 0
        For Each sourceFile As String In selection.Sources
            index += 1
            RaiseEvent Progress("Batch " & index & "/" & selection.Sources.Count & " — " & Path.GetFileName(sourceFile) & "…")

            Dim outcome As New BatchFileOutcome(sourceFile)
            Dim doc As Object = Nothing
            Dim fileTimer As Stopwatch = Stopwatch.StartNew()

            ' --- Temuan E: tentukan & validasi target output ---------------
            Dim pathError As String = ""
            Dim targetPath As String = sourceFile
            If saveInPlace Then
                ' Mode timpa: validasi ketat (folder, ekstensi, read-only).
                pathError = BatchFilePolicy.ValidateOverwriteTarget(sourceFile, folder, includeSubfolders)
            Else
                targetPath = BatchFilePolicy.BuildOutputPath(sourceFile, suffix)
                pathError = BatchFilePolicy.ValidateNewOutputPath(sourceFile, targetPath)
                If pathError = "" Then
                    ' Benturan nama: jangan memakai path yang sama dengan sumber lain
                    ' atau output yang sudah dipesan file sebelumnya pada batch ini.
                    targetPath = BatchFilePolicy.ResolveUniqueOutputPath(
                        targetPath,
                        Function(candidate As String)
                            Dim full As String = candidate
                            Try
                                full = Path.GetFullPath(candidate)
                            Catch
                            End Try
                            Return sourcePaths.Contains(full) OrElse reservedOutputs.Contains(full)
                        End Function)
                    If BatchFilePolicy.SamePath(targetPath, sourceFile) Then
                        pathError = "output akan menimpa file sumber; dibatalkan."
                    End If
                End If
                If pathError = "" AndAlso File.Exists(targetPath) Then
                    outcome.Messages.Add("output sebelumnya ditimpa (regenerasi)")
                End If
            End If

            If pathError <> "" Then
                outcome.FatalError = pathError
                localResult.Failures.Add(Path.GetFileName(sourceFile) & ": " & pathError)
            Else
                Try
                    reservedOutputs.Add(Path.GetFullPath(targetPath))
                Catch
                    reservedOutputs.Add(targetPath)
                End Try
                outcome.OutputPath = targetPath

                Try
                    ' --- buka dokumen ---------------------------------------
                    doc = app.OpenDocument(sourceFile)
                    If doc Is Nothing Then
                        Throw New InvalidOperationException("Dokumen gagal dibuka (mungkin sedang dipakai atau file rusak).")
                    End If

                    ' --- processing -----------------------------------------
                    outcome.ApplyAttempted = True
                    ' writeLogRow:=False — baris CSV untuk file ini ditulis oleh batch
                    ' di bawah, supaya tidak ada dua status berbeda untuk file yang sama.
                    Dim partResult As WatermarkResult = ApplyToDocument(doc, settings, False)
                    outcome.PagesIntended = partResult.PagesIntended
                    outcome.PagesApplied = partResult.PagesPassed
                    outcome.PagesFailed = partResult.PagesFailed
                    outcome.ShapesCreated = partResult.ShapesCreated
                    For Each message As String In partResult.Errors
                        outcome.Messages.Add(message)
                    Next
                    For Each message As String In partResult.Failures
                        outcome.Messages.Add(message)
                    Next

                    localResult.PagesPassed += partResult.PagesPassed
                    localResult.PagesIntended += partResult.PagesIntended
                    localResult.PagesFailed += partResult.PagesFailed
                    localResult.ShapesCreated += partResult.ShapesCreated
                    localResult.Failures.AddRange(partResult.Failures)
                    localResult.Errors.AddRange(partResult.Errors)
                    localResult.Warnings.AddRange(partResult.Warnings)
                    localResult.ProofPdfFiles.AddRange(partResult.ProofPdfFiles)
                    localResult.LogRowsWritten += partResult.LogRowsWritten

                    ' --- save: pemanggilan SaveAs & validasi output DIPISAH ------
                    ' Review lanjutan PR #2 butir 2:
                    '   - SaveCallSucceeded mencatat berhasil/gagalnya PEMANGGILAN,
                    '   - SaveSucceeded hanya True bila penulisan TERBUKTI (berkas
                    '     baru muncul, atau berkas lama berubah ukuran/waktu),
                    '   - file lama yang masih ada TIDAK dianggap bukti sukses,
                    '   - verifikasi memakai metadata file; dokumen TIDAK dibuka ulang
                    '     sehingga dokumen aktif/instance COM pengguna tidak terganggu.
                    outcome.SaveAttempted = True
                    Dim saveMode As SaveOutputMode =
                        If(saveInPlace, SaveOutputMode.OverwriteInPlace, SaveOutputMode.NewFile)
                    Dim beforeSave As SaveOutputSnapshot = SaveOutputSnapshot.Capture(targetPath)

                    Dim saveCallError As String = ""
                    Try
                        doc.SaveAs(targetPath)
                        outcome.SaveCallSucceeded = True
                    Catch saveEx As Exception
                        saveCallError = saveEx.GetType().Name & ": " & saveEx.Message
                        outcome.Messages.Add("gagal menyimpan: " & saveCallError)
                        localResult.Failures.Add(Path.GetFileName(sourceFile) & ": gagal menyimpan — " & saveCallError)
                    End Try

                    If outcome.SaveCallSucceeded Then
                        Dim afterSave As SaveOutputSnapshot = SaveOutputSnapshot.Capture(targetPath)
                        Dim saveCheck As SaveOutputVerification =
                            SaveOutputVerification.Verify(beforeSave, afterSave, saveMode)
                        outcome.SaveVerificationLabel = saveCheck.StatusLabel
                        outcome.SaveVerificationReason = saveCheck.Reason

                        If saveCheck.IsVerified Then
                            outcome.SaveSucceeded = True
                            localResult.DocumentsTouched += 1
                            localResult.OutputFolder = Path.GetDirectoryName(targetPath)
                            localResult.OutputFile = targetPath
                            Report("Disimpan (" & saveCheck.StatusLabel & "): " & targetPath)
                        Else
                            outcome.Messages.Add("penyimpanan tidak terverifikasi (" &
                                                 saveCheck.StatusLabel & "): " & saveCheck.Reason)
                            localResult.Failures.Add(Path.GetFileName(sourceFile) &
                                                     ": penyimpanan tidak terverifikasi (" &
                                                     saveCheck.StatusLabel & ") — " & saveCheck.Reason)
                        End If
                    End If

                Catch ex As Exception
                    outcome.FatalError = ex.GetType().Name & ": " & ex.Message
                    localResult.Failures.Add(Path.GetFileName(sourceFile) & ": " & outcome.FatalError)
                Finally
                    Try
                        If doc IsNot Nothing Then doc.Close()
                    Catch closeEx As Exception
                        ' Kegagalan menutup dokumen tidak boleh disembunyikan, tetapi
                        ' bukan alasan menyatakan output tidak valid.
                        outcome.Messages.Add("dokumen tidak dapat ditutup: " & closeEx.Message)
                        localResult.Warnings.Add(Path.GetFileName(sourceFile) & ": dokumen tidak dapat ditutup: " & closeEx.Message)
                    End Try
                End Try
            End If

            fileTimer.Stop()
            outcome.DurationSeconds = fileTimer.Elapsed.TotalSeconds
            localResult.FileOutcomes.Add(outcome)
            Report("Status " & Path.GetFileName(sourceFile) & ": " & outcome.StatusLabel() & " — " & outcome.MessageText())

            ' --- CSV: memakai status yang SAMA dengan ringkasan & UI ---------
            If settings.WriteCsvLog Then
                Dim entry As New BatchLogEntry()
                entry.FileName = Path.GetFileName(sourceFile)
                entry.Pages = outcome.PagesApplied
                entry.Shapes = outcome.ShapesCreated
                entry.Status = outcome.StatusLabel()
                entry.Message = outcome.MessageText()
                entry.DurationSeconds = outcome.DurationSeconds
                If BatchLog.AppendRow(csvPath, entry) Then localResult.LogRowsWritten += 1
            End If
        Next

        batchTimer.Stop()

        Dim okFiles As Integer = localResult.FilesOk
        Dim partialFiles As Integer = localResult.FilesPartial
        Dim failedFiles As Integer = localResult.FilesFailed

        If settings.WriteCsvLog Then
            If BatchLog.AppendRow(csvPath,
                                  BatchLog.SummaryEntry(settings.Describe(), localResult.FileOutcomes.Count,
                                                        okFiles, partialFiles, failedFiles,
                                                        batchTimer.Elapsed.TotalSeconds)) Then
                localResult.LogRowsWritten += 1
            End If
            Report("Log CSV: " & csvPath)
        End If

        localResult.DurationSeconds = batchTimer.Elapsed.TotalSeconds
        Report("Batch selesai: " & okFiles & " OK, " & partialFiles & " SEBAGIAN, " &
               failedFiles & " GAGAL, " & localResult.SkippedFiles.Count & " dilewati.")

        _result.Merge(localResult)
        Return localResult
    End Function

    ' ========================================================================
    '  HELPER: LAYER
    ' ========================================================================

    Private Function EnsureLayer(page As Object, layerName As String) As Object
        Dim layer As Object = Nothing

        ' Cari layer bernama sama (dipakai ulang supaya tidak menumpuk layer baru).
        Try
            Dim layers As Object = page.Layers
            For i As Integer = 1 To CInt(layers.Count)
                Dim candidate As Object = layers(i)
                Try
                    If CStr(candidate.Name) = layerName Then
                        layer = candidate
                        Exit For
                    End If
                Catch
                End Try
            Next
        Catch
        End Try

        If layer Is Nothing Then
            Try
                layer = page.CreateLayer(layerName)
            Catch
            End Try
        End If

        If layer Is Nothing Then layer = page.ActiveLayer   ' fallback terakhir

        Try
            layer.Editable = True
            layer.Visible = True
            layer.Printable = True      ' wajib agar ikut terekspor ke PDF
        Catch
        End Try
        ' Pastikan layer watermark ada di atas objek desain.
        Try
            layer.MoveToTop()
        Catch
        End Try

        Return layer
    End Function

    ' ========================================================================
    '  HELPER: CARI & HAPUS SHAPE WATERMARK
    ' ========================================================================

    Private Function RemoveWatermarkShapes(page As Object) As Integer
        Dim removed As Integer = 0
        Try
            Dim layers As Object = page.Layers
            For i As Integer = 1 To CInt(layers.Count)
                Dim layer As Object = layers(i)
                Dim toDelete As New List(Of Object)()

                Try
                    Dim shapes As Object = layer.Shapes
                    For j As Integer = 1 To CInt(shapes.Count)
                        Dim shape As Object = shapes(j)
                        Try
                            If IsWatermarkShape(shape) Then toDelete.Add(shape)
                        Catch
                        End Try
                    Next
                Catch
                End Try

                For Each shape As Object In toDelete
                    Try
                        shape.Delete()
                        removed += 1
                    Catch
                    End Try
                Next
            Next
        Catch
        End Try
        Return removed
    End Function

    Private Function IsWatermarkShape(shape As Object) As Boolean
        Try
            If WatermarkLayout.IsWatermarkObjectName(CStr(shape.Name)) Then Return True
        Catch
        End Try
        Try
            If WatermarkLayout.IsWatermarkLayerName(CStr(shape.Layer.Name)) Then Return True
        Catch
        End Try
        ' Shape grup: cek turunannya.
        Try
            If CInt(shape.Type) = WatermarkCdrConst.CdrGroupShape Then
                Dim children As Object = shape.Shapes
                For i As Integer = 1 To CInt(children.Count)
                    If IsWatermarkShape(children(i)) Then Return True
                Next
            End If
        Catch
        End Try
        Return False
    End Function

    ' ========================================================================
    '  HELPER: GAYA, UKURAN, PENEMPATAN
    ' ========================================================================

    ''' <summary>Opasitas + rotasi. Opasitas 25% => transparency 75%.</summary>
    Private Sub StyleObject(shape As Object, settings As WatermarkSettings)
        Dim t As Integer = WatermarkLayout.TransparencyFromOpacity(settings.OpacityPercent)
        Try
            shape.Transparency.ApplyUniformTransparency(t)
            ' Workaround yang didokumentasikan Corel: nilai kadang tidak "nempel".
            Try
                shape.Transparency.Uniform = t
            Catch
            End Try
            Try
                shape.Transparency.AppliedTo = WatermarkCdrConst.CdrApplyToFillAndOutline
            Catch
            End Try
            Try
                shape.Transparency.MergeMode = WatermarkCdrConst.CdrMergeNormal
            Catch
            End Try
        Catch
        End Try

        If Math.Abs(settings.RotationDeg) > 0.01 Then
            Try
                shape.Rotate(settings.RotationDeg)
            Catch
            End Try
        End If
    End Sub

    ''' <summary>Skala objek (jaga rasio aspek bila tinggi tidak diberikan).</summary>
    Private Sub ResizeObject(shape As Object, targetWidth As Double, targetHeight As Double)
        If targetWidth <= 0 Then Return
        Try
            Dim currentWidth As Double = CDbl(shape.SizeWidth)
            Dim currentHeight As Double = CDbl(shape.SizeHeight)
            Dim newHeight As Double = targetHeight
            If newHeight <= 0 AndAlso currentWidth > 0 Then
                newHeight = currentHeight * (targetWidth / currentWidth)
            End If
            shape.SizeWidth = targetWidth
            If newHeight > 0 Then shape.SizeHeight = newHeight
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Atur ukuran lalu posisikan objek (tunggal atau tile). Mengembalikan jumlah objek jadi.
    ''' </summary>
    Private Function PlaceObject(shape As Object, pageWidth As Double, pageHeight As Double,
                                 targetWidth As Double, targetHeight As Double,
                                 settings As WatermarkSettings, tag As String) As Integer
        ResizeObject(shape, targetWidth, targetHeight)

        Dim w As Double = CDbl(shape.SizeWidth)
        Dim h As Double = CDbl(shape.SizeHeight)

        If settings.Tile Then
            Dim positions As List(Of PointF) = WatermarkLayout.TilePositions(
                pageWidth, pageHeight, w, h, settings.TileSpacingXMm, settings.TileSpacingYMm)

            Dim count As Integer = 0
            For i As Integer = 0 To positions.Count - 1
                Dim pt As PointF = positions(i)
                Dim target As Object = shape
                If i > 0 Then
                    Try
                        target = shape.Duplicate(0, 0)
                    Catch
                        Continue For
                    End Try
                End If
                Try
                    target.CenterX = pt.X + w / 2.0
                    target.CenterY = pt.Y + h / 2.0
                    target.Name = WatermarkLayout.ShapePrefix & tag & "_" & (i + 1)
                    count += 1
                Catch
                End Try
            Next
            Return count
        Else
            Dim pt As PointF = WatermarkLayout.AnchorLeftBottom(
                pageWidth, pageHeight, w, h, settings.Anchor, settings.MarginMm)
            Try
                shape.CenterX = pt.X + w / 2.0
                shape.CenterY = pt.Y + h / 2.0
                Return 1
            Catch
                Return 0
            End Try
        End If
    End Function

    ''' <summary>Impor file logo/gambar ke layer, dengan beberapa strategi fallback.</summary>
    Private Function ImportLogo(layer As Object, filePath As String) As Object
        Dim shape As Object = Nothing
        Dim lastMessage As String = ""

        Try
            shape = layer.Import(filePath)
        Catch ex As Exception
            lastMessage = ex.Message
        End Try

        If shape Is Nothing Then
            Try
                shape = layer.ImportEx(filePath, 0, Nothing)
            Catch ex As Exception
                lastMessage = ex.Message
            End Try
        End If

        If shape Is Nothing Then
            Throw New InvalidOperationException(
                "Logo tidak dapat diimpor (" & Path.GetFileName(filePath) & "). " &
                "Pastikan format didukung (PNG/JPG/TIFF) dan file tidak terkunci. Detail: " & lastMessage)
        End If

        Return shape
    End Function
End Class
