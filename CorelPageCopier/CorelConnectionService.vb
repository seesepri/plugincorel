Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  CorelConnectionService.vb
'  Koneksi ke CorelDRAW, identitas instance/dokumen, dan validasi dokumen
'  target. Dipindahkan dari CopierForm (Program.vb) TANPA mengubah perilaku.
'
'  Batas kelas:
'    - Tidak memakai WinForms. Setiap interaksi pengguna lewat ICorelPrompt,
'      sehingga seluruh cabang logika bisa diuji tanpa CorelDRAW.
'    - Pencarian instance COM dan pembuatan instance baru lewat delegate yang
'      bisa diganti (default: CorelComProbe). Ini juga yang membuat test bisa
'      mensimulasikan ROT tanpa COM.
'
'  Perilaku yang dipertahankan (Temuan F audit reliabilitas):
'    - Instance yang SEDANG dipakai pengguna diutamakan (ROT).
'    - Instance kedua TIDAK pernah dibuka diam-diam: pengguna harus setuju.
'    - Objek COM yang sudah tidak valid dilepas, bukan dipakai ulang.
'    - Dokumen target divalidasi masih dokumen + instance yang sama.
' ============================================================================

Imports System

''' <summary>
''' Interaksi pengguna yang dibutuhkan lapisan koneksi. Implementasi UI
''' (CopierForm) menampilkan dialog WinForms; test memakai implementasi palsu.
''' </summary>
Public Interface ICorelPrompt

    ''' <summary>
    ''' Tanya izin membuka instance CorelDRAW baru. False = jangan buka.
    ''' Pesan lengkap (termasuk detail kegagalan) disusun service lewat detail,
    ''' UI hanya menambahkan pembungkus dialognya.
    ''' </summary>
    Function ConfirmOpenNewInstance(detail As String) As Boolean

    ''' <summary>Beri tahu bahwa instance baru sudah dibuka (dengan/tanpa dokumen aktif).</summary>
    Sub NotifyNewInstanceOpened(hasActiveDocument As Boolean)

    ''' <summary>Beri tahu bahwa dokumen aktif berbeda dari saat daftar halaman dimuat.</summary>
    Sub NotifyDocumentMismatch(loaded As DocumentIdentity, current As DocumentIdentity)
End Interface

''' <summary>Hasil satu percobaan koneksi (untuk status bar).</summary>
Public Class ConnectionOutcome
    Public Connected As Boolean = False
    ''' <summary>True bila koneksi memakai instance yang sudah berjalan (ROT).</summary>
    Public AttachedToRunning As Boolean = False
    ''' <summary>True bila instance baru dibuat (setelah disetujui pengguna).</summary>
    Public CreatedNewInstance As Boolean = False
    Public DocumentName As String = ""
    Public Message As String = ""
End Class

Public Class CorelConnectionService

    Private ReadOnly _prompt As ICorelPrompt
    Private ReadOnly _resolveRunning As Func(Of String, Object)
    Private ReadOnly _createInstance As Func(Of String, Object)
    Private ReadOnly _programIds As String()

    Private _app As Object = Nothing
    Private _document As Object = Nothing
    Private _loadedIdentity As DocumentIdentity = Nothing
    Private _lastOutcome As New ConnectionOutcome()

    Public Sub New(prompt As ICorelPrompt)
        Me.New(prompt, Nothing, Nothing, CorelComProbe.DefaultProgramIds())
    End Sub

    Public Sub New(prompt As ICorelPrompt, resolveRunning As Func(Of String, Object),
                   createInstance As Func(Of String, Object), Optional programIds As String() = Nothing)
        _prompt = prompt
        _resolveRunning = If(resolveRunning, AddressOf CorelComProbe.GetRunningObject)
        _createInstance = If(createInstance, AddressOf CorelComProbe.CreateInstance)
        _programIds = If(programIds, CorelComProbe.DefaultProgramIds())
    End Sub

    ''' <summary>Objek Application CorelDRAW (late bound) atau Nothing.</summary>
    Public ReadOnly Property App As Object
        Get
            Return _app
        End Get
    End Property

    ''' <summary>Objek Document aktif (late bound) atau Nothing.</summary>
    Public ReadOnly Property Document As Object
        Get
            Return _document
        End Get
    End Property

    ''' <summary>Identitas dokumen saat daftar halaman terakhir dimuat.</summary>
    Public ReadOnly Property LoadedIdentity As DocumentIdentity
        Get
            Return _loadedIdentity
        End Get
    End Property

    Public ReadOnly Property ProgramIds As String()
        Get
            Return _programIds
        End Get
    End Property

    Public ReadOnly Property LastOutcome As ConnectionOutcome
        Get
            Return _lastOutcome
        End Get
    End Property

    ''' <summary>
    ''' Hubungkan ke CorelDRAW dan siapkan Document aktif.
    ''' Melempar InvalidOperationException dengan pesan yang sama seperti versi
    ''' lama (pesan itu ditampilkan UI apa adanya).
    ''' </summary>
    Public Function Connect() As ConnectionOutcome
        Dim outcome As New ConnectionOutcome()
        Dim lastError As Exception = Nothing

        ' (1) Pakai koneksi yang sudah berhasil sebelumnya bila masih valid.
        If _app IsNot Nothing Then
            If ComObjectHealth.IsAlive(_app) Then
                Try
                    Dim existingDoc As Object = _app.ActiveDocument
                    If existingDoc IsNot Nothing Then
                        _document = existingDoc
                        outcome.Connected = True
                        outcome.DocumentName = SafeDocumentName(existingDoc)
                        outcome.Message = "Koneksi sebelumnya dipakai ulang."
                        _lastOutcome = outcome
                        Return outcome
                    End If
                    lastError = New InvalidOperationException("CorelDRAW terhubung, tetapi belum ada dokumen aktif.")
                Catch ex As Exception
                    lastError = ex
                End Try
            Else
                lastError = New InvalidOperationException("Koneksi CorelDRAW sebelumnya sudah tidak valid (instance ditutup?).")
            End If
            Release()
        End If

        ' (2) Lampirkan ke instance yang SUDAH berjalan (ROT).
        For Each progId As String In _programIds
            Dim running As Object = Nothing
            Try
                running = _resolveRunning(progId)
            Catch ex As Exception
                lastError = ex
            End Try

            If running IsNot Nothing Then
                Try
                    Dim runningDoc As Object = running.ActiveDocument
                    If runningDoc IsNot Nothing Then
                        _app = running
                        _document = runningDoc
                        outcome.Connected = True
                        outcome.AttachedToRunning = True
                        outcome.DocumentName = SafeDocumentName(runningDoc)
                        outcome.Message = "Terhubung ke instance CorelDRAW yang sedang berjalan."
                        _lastOutcome = outcome
                        Return outcome
                    End If
                    lastError = New InvalidOperationException(
                        "CorelDRAW (" & progId & ") ditemukan, tetapi tidak ada dokumen aktif. " &
                        "Buka dokumen CDR di jendela itu, lalu klik Muat ulang.")
                    ComObjectHealth.TryRelease(running)
                Catch ex As Exception
                    lastError = ex
                    ComObjectHealth.TryRelease(running)
                End Try
            End If
        Next

        ' (3) Tidak ada instance berjalan yang bisa dipakai: MINTA IZIN dulu.
        Dim detail As String = If(lastError Is Nothing, "Tidak ada objek aplikasi yang tersedia.", lastError.Message)
        If Not _prompt.ConfirmOpenNewInstance(detail) Then
            Throw New InvalidOperationException(
                "Koneksi dibatalkan oleh pengguna. Tidak ada instance CorelDRAW baru yang dibuka. " &
                "Buka dokumen CDR di CorelDRAW lalu klik Muat ulang. Detail: " & detail, lastError)
        End If

        For Each progId As String In _programIds
            Dim newApp As Object = Nothing
            Try
                newApp = _createInstance(progId)
                If newApp Is Nothing Then Continue For

                Try
                    newApp.Visible = True
                Catch
                    ' Visible bersifat kosmetik; kegagalan tidak membatalkan koneksi.
                End Try

                Dim newDoc As Object = Nothing
                Try
                    newDoc = newApp.ActiveDocument
                Catch
                    newDoc = Nothing
                End Try

                If newDoc IsNot Nothing Then
                    _app = newApp
                    _document = newDoc
                    _prompt.NotifyNewInstanceOpened(True)
                    outcome.Connected = True
                    outcome.CreatedNewInstance = True
                    outcome.DocumentName = SafeDocumentName(newDoc)
                    outcome.Message = "Instance CorelDRAW baru dibuka; dokumen aktifnya dipakai sebagai target."
                    _lastOutcome = outcome
                    Return outcome
                End If

                _prompt.NotifyNewInstanceOpened(False)
                ComObjectHealth.TryRelease(newApp)
                Throw New InvalidOperationException(
                    "CorelDRAW baru dibuka melalui COM, tetapi belum memiliki dokumen aktif.")
            Catch ex As Exception
                lastError = ex
            End Try
        Next

        Throw New InvalidOperationException(
            "Tidak dapat mengakses dokumen CorelDRAW. Pastikan CorelDRAW 2022 terpasang dan dokumen CDR aktif. " &
            "Detail: " & If(lastError Is Nothing, "tidak diketahui", lastError.Message), lastError)
    End Function

    ''' <summary>Lepas referensi COM lama supaya tidak dianggap target yang masih hidup.</summary>
    Public Sub Release()
        If _document IsNot Nothing Then ComObjectHealth.TryRelease(_document)
        If _app IsNot Nothing Then ComObjectHealth.TryRelease(_app)
        _document = Nothing
        _app = Nothing
    End Sub

    ''' <summary>Identitas instance + dokumen aktif saat ini (late bound, aman gagal).</summary>
    Public Function CurrentIdentity() As DocumentIdentity
        Return DocumentIdentity.FromDocument(_document, DocumentIdentity.InstanceTokenFor(_app))
    End Function

    ''' <summary>
    ''' Ambil dokumen aktif dari App yang sudah terhubung; melempar bila tidak ada.
    ''' Pesan sama dengan versi lama supaya UI tidak berubah.
    ''' </summary>
    Public Function EnsureDocument() As Object
        If _app Is Nothing Then
            Throw New InvalidOperationException(
                "Belum terhubung ke CorelDRAW. Buka CorelDRAW dan dokumen CDR terlebih dahulu.")
        End If
        _document = _app.ActiveDocument
        If _document Is Nothing Then
            Throw New InvalidOperationException(
                "Tidak ada dokumen aktif di CorelDRAW. Buka file CDR lalu klik Muat ulang.")
        End If
        Return _document
    End Function

    ''' <summary>Simpan identitas dokumen (dipanggil setelah daftar halaman dimuat).</summary>
    Public Function RememberLoadedIdentity() As DocumentIdentity
        _loadedIdentity = CurrentIdentity()
        Return _loadedIdentity
    End Function

    ''' <summary>
    ''' Pastikan target masih dokumen + instance yang sama seperti saat daftar
    ''' halaman dimuat. Ketidakcocokan dilaporkan lewat prompt lalu False.
    ''' </summary>
    Public Function EnsureSameDocument() As Boolean
        If _loadedIdentity Is Nothing Then
            _loadedIdentity = CurrentIdentity()
            Return True
        End If
        Dim current As DocumentIdentity = CurrentIdentity()
        If current.SameTarget(_loadedIdentity) Then
            ' Jumlah halaman boleh berubah (mis. halaman ditambah), cukup perbarui.
            _loadedIdentity = current
            Return True
        End If

        _prompt.NotifyDocumentMismatch(_loadedIdentity, current)
        Return False
    End Function

    ''' <summary>True bila objek Application yang dipegang masih hidup.</summary>
    Public Function IsAppAlive() As Boolean
        Return _app IsNot Nothing AndAlso ComObjectHealth.IsAlive(_app)
    End Function

    Private Shared Function SafeDocumentName(doc As Object) As String
        If doc Is Nothing Then Return ""
        Try
            Return CStr(doc.Name)
        Catch
            Return ""
        End Try
    End Function
End Class
