Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  SelectionNavigation.vb
'  State navigasi hasil "Auto Select — semua halaman" (murni, bisa diuji unit).
'
'  Bug yang diperbaiki (implementasi lama di Program.vb):
'    - Daftar halaman hasil (`matchingPageIndices`) dihitung dengan kategori
'      saat pemindaian, tetapi `SelectMatchingPageAtCurrentPosition` dipanggil
'      dengan `categoryBox.SelectedIndex` SAAT ITU. Bila pengguna mengganti
'      kategori sebelum menekan "Pilih hasil berikutnya", halaman yang dipilih
'      dihitung ulang dengan kategori berbeda -> halaman bisa terlewat atau
'      tampil 0 objek, sementara hitungan di layar tetap dari hasil lama.
'    - Tidak ada pencatatan kategori/cakupan pemindaian, sehingga tidak mungkin
'      mendeteksi perubahan itu.
'    - Posisi tetap maju walau halaman tujuan sudah tidak ada (dokumen diedit),
'      sehingga hasil berikutnya terlewat.
'
'  Solusi: snapshot (kategori + cakupan) disimpan bersama daftar halaman, dan
'  navigasi memakai snapshot tersebut. Perubahan UI terdeteksi lewat
'  MatchesSnapshot dan dilaporkan, bukan diam-diam dipakai.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Globalization

''' <summary>Hasil satu langkah navigasi.</summary>
Public Enum NavigationStep
    ''' <summary>Belum ada pemindaian (daftar halaman kosong).</summary>
    NotStarted = 0
    ''' <summary>Berpindah ke halaman berikutnya.</summary>
    Moved = 1
    ''' <summary>Sudah di hasil terakhir; navigasi dihentikan.</summary>
    AtEnd = 2
    ''' <summary>Halaman hasil tidak ada lagi di dokumen (dokumen berubah).</summary>
    PageMissing = 3
    ''' <summary>Snapshot pemindaian tidak cocok dengan pilihan UI sekarang.</summary>
    StaleSnapshot = 4
End Enum

''' <summary>Snapshot pemindaian + posisi navigasi. Satu instance per hasil pindai.</summary>
Public Class MatchNavigator

    Private ReadOnly _pages As New List(Of Integer)()
    Private _position As Integer = -1
    Private _category As ShapeCategory = ShapeCategory.Text
    Private _allPages As Boolean = False
    Private _scanCompleted As Boolean = False

    Public ReadOnly Property Pages As List(Of Integer)
        Get
            Return _pages
        End Get
    End Property

    Public ReadOnly Property Position As Integer
        Get
            Return _position
        End Get
    End Property

    Public ReadOnly Property Count As Integer
        Get
            Return _pages.Count
        End Get
    End Property

    Public ReadOnly Property Category As ShapeCategory
        Get
            Return _category
        End Get
    End Property

    Public ReadOnly Property AllPages As Boolean
        Get
            Return _allPages
        End Get
    End Property

    ''' <summary>True bila pemindaian selesai dan ada halaman hasil.</summary>
    Public ReadOnly Property HasResults As Boolean
        Get
            Return _scanCompleted AndAlso _pages.Count > 0
        End Get
    End Property

    ''' <summary>Halaman hasil saat ini (0 bila belum mulai).</summary>
    Public ReadOnly Property CurrentPage As Integer
        Get
            If _position < 0 OrElse _position >= _pages.Count Then Return 0
            Return _pages(_position)
        End Get
    End Property

    ''' <summary>"3/5" untuk status bar.</summary>
    Public ReadOnly Property PositionLabel As String
        Get
            If _pages.Count = 0 Then Return "0/0"
            Return (_position + 1).ToString(CultureInfo.InvariantCulture) & "/" &
                   _pages.Count.ToString(CultureInfo.InvariantCulture)
        End Get
    End Property

    ''' <summary>
    ''' Mulai sesi navigasi baru. Hanya halaman dengan hasil > 0 yang masuk, dan
    ''' daftar diurutkan + di-dedupe agar tidak ada halaman yang diulang.
    ''' </summary>
    Public Sub Start(scanCategory As ShapeCategory, scanAllPages As Boolean, pageIndices As IEnumerable(Of Integer))
        _pages.Clear()
        _category = scanCategory
        _allPages = scanAllPages
        _position = -1
        _scanCompleted = False

        If pageIndices IsNot Nothing Then
            For Each pageIndex As Integer In pageIndices
                If pageIndex >= 1 AndAlso Not _pages.Contains(pageIndex) Then _pages.Add(pageIndex)
            Next
        End If
        _pages.Sort()
        _scanCompleted = True
        If _pages.Count > 0 Then _position = 0
    End Sub

    ''' <summary>Batalkan sesi (mis. pemindaian gagal di tengah jalan).</summary>
    Public Sub Reset()
        _pages.Clear()
        _position = -1
        _scanCompleted = False
    End Sub

    ''' <summary>True bila pilihan UI sekarang berbeda dari saat pemindaian.</summary>
    Public Function MatchesSnapshot(category As ShapeCategory, allPages As Boolean) As Boolean
        Return category = _category AndAlso allPages = _allPages
    End Function

    ''' <summary>
    ''' Pindah ke hasil berikutnya. pageExists dipakai untuk memvalidasi halaman
    ''' masih ada di dokumen; bila tidak ada, posisi TIDAK dimajukan.
    ''' </summary>
    Public Function MoveNext(category As ShapeCategory, allPages As Boolean,
                             pageExists As Func(Of Integer, Boolean)) As NavigationStep
        If Not _scanCompleted OrElse _pages.Count = 0 Then Return NavigationStep.NotStarted
        If Not MatchesSnapshot(category, allPages) Then Return NavigationStep.StaleSnapshot
        If _position >= _pages.Count - 1 Then Return NavigationStep.AtEnd

        Dim nextPosition As Integer = _position + 1
        Dim nextPage As Integer = _pages(nextPosition)
        If pageExists IsNot Nothing AndAlso Not pageExists(nextPage) Then
            ' Halaman hilang (dokumen diedit): jangan majukan posisi agar hasil
            ' berikutnya tidak terlewat.
            Return NavigationStep.PageMissing
        End If

        _position = nextPosition
        Return NavigationStep.Moved
    End Function

    ''' <summary>True bila posisi sekarang adalah hasil terakhir.</summary>
    Public ReadOnly Property IsAtEnd As Boolean
        Get
            Return _pages.Count > 0 AndAlso _position >= _pages.Count - 1
        End Get
    End Property

    ''' <summary>Pesan siap tampil untuk pengguna.</summary>
    Public Function StatusText() As String
        If _pages.Count = 0 Then Return "Belum ada hasil pemindaian."
        Return "Hasil " & PositionLabel & " | halaman " & CurrentPage.ToString(CultureInfo.InvariantCulture) &
               " | kategori: " & ShapeClassifier.CategoryLabel(_category)
    End Function
End Class

''' <summary>
''' Hasil satu operasi pemilihan objek di CorelDRAW.
''' Dipisah dari COM/WinForms supaya aturan "verified / tidak" bisa diuji unit:
''' jumlah objek yang benar-benar terpilih dibandingkan dengan jumlah target,
''' karena AddToSelection() bisa tidak berefek (layer terkunci, objek di grup).
''' </summary>
Public Class SelectionOutcome
    ''' <summary>Jumlah target unik yang seharusnya terpilih.</summary>
    Public ExpectedCount As Integer = 0
    ''' <summary>Jumlah pemanggilan AddToSelection yang tidak melempar exception.</summary>
    Public AddedCalls As Integer = 0
    ''' <summary>Jumlah objek yang benar-benar terpilih menurut CorelDRAW.</summary>
    Public ActualCount As Integer = 0
    Public ActualCountKnown As Boolean = False
    ''' <summary>Jumlah objek yang cocok (bisa &gt; ExpectedCount untuk isi PowerClip).</summary>
    Public MatchCount As Integer = 0
    Public CategoryLabel As String = ""
    Public Failures As New List(Of String)()
    ''' <summary>Pesan bila operasi tidak dijalankan (mis. tidak ada hasil).</summary>
    Public NotRunReason As String = ""

    Public Shared Function NotRun(reason As String) As SelectionOutcome
        Dim outcome As New SelectionOutcome()
        outcome.NotRunReason = If(reason, "")
        Return outcome
    End Function

    Public ReadOnly Property WasAttempted As Boolean
        Get
            Return String.IsNullOrEmpty(NotRunReason)
        End Get
    End Property

    ''' <summary>
    ''' True hanya bila seleksi benar-benar terverifikasi: ada target dan jumlah
    ''' yang dilaporkan CorelDRAW sama dengan jumlah target.
    ''' </summary>
    Public ReadOnly Property SelectionVerified As Boolean
        Get
            Return WasAttempted AndAlso ExpectedCount > 0 AndAlso ActualCountKnown AndAlso
                   ActualCount = ExpectedCount AndAlso Failures.Count = 0
        End Get
    End Property

    ''' <summary>Angka yang boleh diklaim sebagai "terpilih".</summary>
    Public ReadOnly Property SelectionVerifiedCount As Integer
        Get
            If Not WasAttempted Then Return 0
            If ActualCountKnown Then Return ActualCount
            Return AddedCalls
        End Get
    End Property

    Public Function Describe() As String
        If Not WasAttempted Then Return "Seleksi tidak dijalankan: " & NotRunReason

        Dim sb As New System.Text.StringBuilder()
        If SelectionVerified Then
            sb.AppendLine("Seleksi terverifikasi: " & ActualCount.ToString() & " objek terpilih")
            sb.AppendLine("  (target unik " & ExpectedCount.ToString() & ", objek cocok " & MatchCount.ToString() & ").")
        Else
            sb.AppendLine("Seleksi BELUM terverifikasi.")
            sb.AppendLine("  Target unik     : " & ExpectedCount.ToString())
            sb.AppendLine("  AddToSelection  : " & AddedCalls.ToString() & " panggilan berhasil")
            If ActualCountKnown Then
                sb.AppendLine("  Terpilih (COM)  : " & ActualCount.ToString())
            Else
                sb.AppendLine("  Terpilih (COM)  : tidak dapat dibaca")
            End If
        End If
        If Failures.Count > 0 Then
            sb.AppendLine("  Catatan:")
            For Each failure As String In Failures
                sb.AppendLine("   • " & failure)
            Next
        End If
        Return sb.ToString().TrimEnd()
    End Function
End Class
