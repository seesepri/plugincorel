Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  CopyService.vb
'  Pipeline "Salin objek ke halaman" dipindahkan dari CopierForm.CopyClickedCore
'  TANPA mengubah urutan, keputusan, atau pesan.
'
'  Batas kelas:
'    - Tidak memakai WinForms. Konfirmasi & pemberitahuan lewat ICopyPrompt,
'      progress lewat IProgressSink.
'    - Semua jalur batal mengembalikan CopyAbortReason; UI memutuskan dialog
'      mana yang ditampilkan (teks dialog tetap milik UI).
'    - Validasi daftar halaman tujuan dipisah ke CopyTargetValidation (murni)
'      supaya bisa diuji tanpa COM.
'
'  Jaminan yang dipertahankan dari audit reliabilitas:
'    A. Verifikasi per halaman (Verified / Partial / Failed) memakai objek hasil
'       paste, bukan shape terbawah.
'    B. Satuan: wajib terbaca, dinormalkan ke mm, dipulihkan di Finally.
'    C. Grup undo: kegagalan pembukaan dilaporkan & butuh persetujuan pengguna;
'       EndCommandGroup hanya bila grup benar-benar terbuka.
'    F. Identitas dokumen + halaman ditampilkan sebelum dokumen diubah.
'    Objek sumber tidak pernah dihapus/diubah; tidak ada fallback Delete.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics

''' <summary>Alasan operasi copy tidak dijalankan (bukan kegagalan teknis).</summary>
Public Enum CopyAbortReason
    ''' <summary>Operasi berjalan (lihat CopyServiceResult.Executed).</summary>
    None = 0
    ''' <summary>Tidak ada halaman tujuan yang dicentang.</summary>
    NoTargetPages = 1
    ''' <summary>Semua halaman tujuan tidak valid (di luar rentang / halaman sumber).</summary>
    NoValidTargets = 2
    ''' <summary>Dokumen/instance aktif berbeda dari saat daftar halaman dimuat.</summary>
    DocumentChanged = 3
    ''' <summary>Halaman aktif bukan lagi halaman sumber.</summary>
    SourcePageChanged = 4
    ''' <summary>Tidak ada objek terpilih di halaman sumber.</summary>
    NoSelection = 5
    ''' <summary>Satuan dokumen tidak dapat dibaca.</summary>
    UnitUnreadable = 6
    ''' <summary>Pengguna membatalkan konfirmasi.</summary>
    UserCancelled = 7
    ''' <summary>Satuan dokumen tidak dapat dinormalkan ke mm.</summary>
    UnitNotNormalizable = 8
    ''' <summary>Grup undo tidak tersedia dan pengguna memilih berhenti.</summary>
    UndoGroupDeclined = 9
End Enum

''' <summary>Rincian yang ditampilkan UI sebelum dokumen diubah.</summary>
Public Class CopyPlan
    Public DocumentIdentity As String = ""
    Public SourcePage As Integer = 0
    Public Targets As New List(Of Integer)()
    Public UnitName As String = ""
    Public Warnings As New List(Of String)()

    Public ReadOnly Property TargetCount As Integer
        Get
            Return Targets.Count
        End Get
    End Property
End Class

''' <summary>
''' Interaksi pengguna selama operasi copy. Implementasi UI memakai MessageBox
''' dengan teks yang sama seperti sebelum refactor.
''' </summary>
Public Interface ICopyPrompt

    ''' <summary>Konfirmasi sebelum dokumen diubah. False = batalkan.</summary>
    Function ConfirmCopy(plan As CopyPlan) As Boolean

    ''' <summary>Konfirmasi lanjut tanpa grup undo. False = batalkan.</summary>
    Function ConfirmContinueWithoutUndoGroup(reason As String) As Boolean

    ''' <summary>Beri tahu bahwa satuan dokumen gagal dipulihkan.</summary>
    Sub NotifyUnitRestoreFailed(targetUnitName As String, errorMessage As String)
End Interface

''' <summary>Hasil satu operasi copy (data untuk UI; tidak menampilkan dialog).</summary>
Public Class CopyServiceResult
    ''' <summary>True bila paste benar-benar dijalankan.</summary>
    Public Executed As Boolean = False
    Public AbortReason As CopyAbortReason = CopyAbortReason.None
    ''' <summary>Detail untuk dialog batal (mis. alasan blokir satuan). Boleh kosong.</summary>
    Public AbortDetail As String = ""
    ''' <summary>Peringatan yang WAJIB ikut ditampilkan.</summary>
    Public Warnings As New List(Of String)()
    Public Summary As New CopyRunSummary()
    Public IdentityText As String = ""
    Public UnitNote As String = ""
    Public UndoNote As String = ""
    Public DurationSeconds As Double = 0
    Public RestoreFailed As Boolean = False
    Public FatalError As Exception = Nothing

    Public Function BuildReport() As String
        Return Summary.BuildReport(IdentityText, UnitNote, UndoNote, DurationSeconds)
    End Function

    Public ReadOnly Property AllVerified As Boolean
        Get
            Return Summary.AllVerified
        End Get
    End Property

    Public ReadOnly Property Headline As String
        Get
            Return Summary.Headline(DurationSeconds)
        End Get
    End Property
End Class

''' <summary>
''' Validasi daftar halaman tujuan — murni, tanpa COM, sehingga bisa diuji
''' terpisah. Aturan sama dengan versi lama:
'''   - indeks di luar 1..pageCount dilewati (dicatat),
'''   - halaman sumber tidak boleh menjadi tujuan (dicatat),
'''   - duplikat dibuang, urutan dipertahankan.
''' </summary>
Public Class CopyTargetValidation

    Public ReadOnly Property ValidTargets As New List(Of Integer)()
    Public ReadOnly Property Warnings As New List(Of String)()

    Public Shared Function Validate(targets As IEnumerable(Of Integer), sourcePage As Integer,
                                    pageCount As Integer) As CopyTargetValidation
        Dim result As New CopyTargetValidation()
        If targets Is Nothing Then Return result

        For Each pageIndex As Integer In targets
            If pageIndex < 1 OrElse pageIndex > pageCount Then
                result.Warnings.Add("Halaman " & pageIndex & ": indeks halaman tidak valid — dilewati.")
            ElseIf pageIndex = sourcePage Then
                result.Warnings.Add("Halaman " & pageIndex & ": halaman sumber tidak boleh menjadi tujuan — dilewati.")
            ElseIf Not result.ValidTargets.Contains(pageIndex) Then
                result.ValidTargets.Add(pageIndex)
            End If
        Next
        Return result
    End Function
End Class

Public Class CopyService

    Private ReadOnly _connection As CorelConnectionService
    Private ReadOnly _prompt As ICopyPrompt
    Private ReadOnly _sink As IProgressSink

    Public Sub New(connection As CorelConnectionService, prompt As ICopyPrompt, sink As IProgressSink)
        If connection Is Nothing Then Throw New ArgumentNullException(NameOf(connection))
        If prompt Is Nothing Then Throw New ArgumentNullException(NameOf(prompt))
        _connection = connection
        _prompt = prompt
        _sink = If(sink, New NullProgressSink())
    End Sub

    ''' <summary>
    ''' Salin objek terpilih dari halaman sumber ke halaman tujuan.
    ''' Tidak pernah menampilkan dialog sendiri; hasilnya dikembalikan ke UI.
    ''' </summary>
    Public Function Run(sourcePageIndex As Integer, targets As IEnumerable(Of Integer)) As CopyServiceResult
        Dim stopwatch As Stopwatch = Stopwatch.StartNew()
        Dim result As New CopyServiceResult()

        If targets Is Nothing OrElse Not HasAny(targets) Then
            result.AbortReason = CopyAbortReason.NoTargetPages
            Return result
        End If

        ' Status dokumen yang wajib dipulihkan (Temuan B & C).
        Dim unitScope As New DocumentUnitScope(WatermarkCdrConst.CdrMillimeter)
        Dim undoGuard As New UndoGroupGuard("Salin objek ke halaman")
        Dim sourcePageToRestore As Integer = -1
        Dim fatalError As Exception = Nothing

        Try
            _connection.Connect()
            If Not _connection.EnsureSameDocument() Then
                result.AbortReason = CopyAbortReason.DocumentChanged
                Return result
            End If

            Dim document As Object = _connection.Document
            Dim app As Object = _connection.App
            Dim loadedIdentity As DocumentIdentity = _connection.LoadedIdentity
            result.IdentityText = If(loadedIdentity Is Nothing, "", loadedIdentity.Describe())

            Dim actualSourcePage As Integer = CInt(document.ActivePage.Index)
            If actualSourcePage <> sourcePageIndex Then
                result.AbortReason = CopyAbortReason.SourcePageChanged
                Return result
            End If
            sourcePageToRestore = actualSourcePage

            Dim documentPageCount As Integer = CInt(document.Pages.Count)
            Dim validation As CopyTargetValidation = CopyTargetValidation.Validate(
                targets, actualSourcePage, documentPageCount)
            result.Warnings.AddRange(validation.Warnings)

            If validation.ValidTargets.Count = 0 Then
                result.AbortReason = CopyAbortReason.NoValidTargets
                result.AbortDetail = String.Join(Environment.NewLine, validation.Warnings.ToArray())
                Return result
            End If

            Dim selection As Object = Nothing
            Try
                selection = app.ActiveSelectionRange
            Catch
                selection = Nothing
            End Try
            If selection Is Nothing OrElse CInt(selection.Count) = 0 Then
                result.AbortReason = CopyAbortReason.NoSelection
                Return result
            End If

            ' --- Temuan B: baca satuan asli SEBELUM menyentuh dokumen -----------
            Try
                unitScope.MarkOriginalKnown(CInt(document.Unit))
            Catch ex As Exception
                unitScope.MarkOriginalUnknown(ex.Message)
            End Try
            If Not unitScope.OriginalUnitKnown Then
                result.AbortReason = CopyAbortReason.UnitUnreadable
                result.AbortDetail = unitScope.BlockingReason
                Return result
            End If

            ' --- Temuan F: konfirmasi (identitas + halaman + satuan) ------------
            Dim plan As New CopyPlan()
            plan.DocumentIdentity = result.IdentityText
            plan.SourcePage = actualSourcePage
            plan.Targets.AddRange(validation.ValidTargets)
            plan.UnitName = DocumentUnits.UnitName(unitScope.OriginalUnit)
            plan.Warnings.AddRange(validation.Warnings)
            If Not _prompt.ConfirmCopy(plan) Then
                result.AbortReason = CopyAbortReason.UserCancelled
                Return result
            End If

            ' --- Temuan B: normalisasi satuan -----------------------------------
            Try
                document.Unit = unitScope.TargetUnit
                unitScope.MarkChangeApplied()
            Catch ex As Exception
                unitScope.MarkChangeFailed(ex.Message)
            End Try
            If Not unitScope.CanVerifyInMillimeters Then
                result.AbortReason = CopyAbortReason.UnitNotNormalizable
                result.AbortDetail = unitScope.BlockingReason
                Return result
            End If
            Dim unitFactor As Double = unitScope.MeasurementFactorMm

            ' Ukur objek SUMBER sebagai acuan verifikasi.
            Dim sourceSet As MeasuredShapeSet = CopyMeasurement.MeasureSelection(selection)

            ' --- Temuan C: grup undo --------------------------------------------
            undoGuard.MarkBeginAttempted()
            Try
                document.BeginCommandGroup(undoGuard.Name)
                undoGuard.MarkBeginSucceeded()
            Catch ex As Exception
                undoGuard.MarkBeginFailed(ex.Message)
            End Try
            If undoGuard.BeginFailed Then
                If Not _prompt.ConfirmContinueWithoutUndoGroup(undoGuard.BeginError) Then
                    result.AbortReason = CopyAbortReason.UndoGroupDeclined
                    result.Warnings.Add(undoGuard.Describe())
                    Return result
                End If
                result.Warnings.Add(undoGuard.Describe())
            End If

            ' Objek sumber TIDAK PERNAH dihapus sebagai bagian dari validasi.
            selection.Copy()
            result.Executed = True

            For targetPosition As Integer = 0 To validation.ValidTargets.Count - 1
                Dim pageIndex As Integer = validation.ValidTargets(targetPosition)
                _sink.Report("Copy progres: " & (targetPosition + 1) & "/" & validation.ValidTargets.Count &
                             " — halaman tujuan " & pageIndex & ".")
                _sink.Pump()
                Try
                    Dim destination As Object = document.Pages(pageIndex)
                    destination.Activate()
                    Dim layer As Object = destination.ActiveLayer
                    Dim layerShapes As Object = layer.Shapes
                    Dim beforeCount As Integer = CInt(layerShapes.Count)
                    layer.Paste()
                    Dim afterCount As Integer = CInt(layerShapes.Count)
                    Dim added As Integer = afterCount - beforeCount

                    Dim pastedSelection As Object = Nothing
                    Try
                        pastedSelection = app.ActiveSelectionRange
                    Catch
                        pastedSelection = Nothing
                    End Try

                    ' Temuan A: ukur objek hasil paste yang BENAR, lalu verifikasi.
                    Dim pastedSet As MeasuredShapeSet = CopyMeasurement.MeasurePasted(
                        pastedSelection, layerShapes, added)
                    Dim verification As CopyVerificationResult = CopyVerification.VerifyPasted(
                        sourceSet, pastedSet, added, unitFactor, CopyVerification.ToleranceMm)
                    result.Summary.Add(pageIndex, verification.Status, verification.Reason)
                Catch pageEx As Exception
                    result.Summary.Add(pageIndex, CopyVerifyStatus.Failed,
                                       pageEx.GetType().Name & ": " & pageEx.Message)
                End Try
            Next
        Catch ex As Exception
            fatalError = ex
        Finally
            Dim document As Object = _connection.Document

            ' --- Temuan C: EndCommandGroup HANYA bila group benar-benar terbuka --
            If undoGuard.ShouldCallEnd AndAlso document IsNot Nothing Then
                Try
                    document.EndCommandGroup()
                    undoGuard.MarkEndSucceeded()
                Catch ex As Exception
                    undoGuard.MarkEndFailed(ex.Message)
                End Try
            End If

            ' --- Temuan B: pulihkan satuan asli, laporkan bila gagal ------------
            If unitScope.RestoreRequired AndAlso document IsNot Nothing Then
                Try
                    document.Unit = unitScope.OriginalUnit
                    unitScope.MarkRestored()
                Catch ex As Exception
                    unitScope.MarkRestoreFailed(ex.Message)
                End Try
            End If

            If sourcePageToRestore > 0 AndAlso document IsNot Nothing Then
                Try
                    document.Pages(sourcePageToRestore).Activate()
                Catch ex As Exception
                    result.Warnings.Add("Halaman sumber tidak dapat diaktifkan kembali: " & ex.Message)
                End Try
            End If
        End Try

        ' ---- Laporan dibuat SETELAH Finally supaya kegagalan pemulihan ikut
        ' ---- dilaporkan (Temuan B & C) dan tidak mungkin disembunyikan.
        If undoGuard.HasProblem Then result.Warnings.Add(undoGuard.Describe())
        result.RestoreFailed = unitScope.RestoreFailed
        If unitScope.RestoreFailed Then
            result.Warnings.Add("PEMULIHAN SATUAN DOKUMEN GAGAL (" & unitScope.RestoreError &
                                "). Satuan dokumen kemungkinan tetap mm — ubah manual di CorelDRAW.")
            Try
                _prompt.NotifyUnitRestoreFailed(DocumentUnits.UnitName(unitScope.OriginalUnit), unitScope.RestoreError)
            Catch
                ' Pemberitahuan bersifat informatif; kegagalan UI tidak boleh
                ' menyembunyikan hasil operasi.
            End Try
        End If

        stopwatch.Stop()
        result.DurationSeconds = stopwatch.Elapsed.TotalSeconds
        result.FatalError = fatalError
        If fatalError IsNot Nothing Then
            result.Warnings.Add("Operasi dihentikan oleh kesalahan: " & fatalError.GetType().Name & ": " & fatalError.Message)
        End If

        result.UnitNote = unitScope.Describe()
        result.UndoNote = undoGuard.Describe()
        Return result
    End Function

    Private Shared Function HasAny(items As IEnumerable(Of Integer)) As Boolean
        For Each item As Integer In items
            Return True
        Next
        Return False
    End Function
End Class
