Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  UndoGroup.vb
'  Pencatat status grup undo CorelDRAW (Temuan C).
'
'  Masalah pada kode lama:
'    - `Try doc.BeginCommandGroup(name) : inCommandGroup = True Catch End Try`
'      menelan kegagalan, sehingga laporan tetap menjanjikan "satu Ctrl+Z".
'    - Kegagalan EndCommandGroup juga ditelan.
'
'  Aturan baru:
'    - Kegagalan BeginCommandGroup dicatat + dilaporkan, dan pemanggil wajib
'      bertanya ke pengguna sebelum melanjutkan tanpa jaminan satu langkah undo.
'    - EndCommandGroup HANYA dipanggil bila group benar-benar terbuka.
'    - Kegagalan EndCommandGroup dicatat dan dilaporkan (tidak ditelan).
'
'  Kelas ini murni (tanpa COM) sehingga bisa diuji unit.
' ============================================================================

Imports System

''' <summary>Pelacak state satu grup undo CorelDRAW.</summary>
Public Class UndoGroupGuard

    Private _beginAttempted As Boolean = False
    Private _beginSucceeded As Boolean = False
    Private _beginError As String = ""
    Private _endAttempted As Boolean = False
    Private _endSucceeded As Boolean = False
    Private _endError As String = ""

    Public Sub New(Optional name As String = "")
        Me.Name = If(name, "")
    End Sub

    ''' <summary>Nama grup yang dikirim ke CorelDRAW (muncul di daftar Undo).</summary>
    Public ReadOnly Property Name As String

    Public ReadOnly Property BeginAttempted As Boolean
        Get
            Return _beginAttempted
        End Get
    End Property

    ''' <summary>True bila BeginCommandGroup berhasil.</summary>
    Public ReadOnly Property IsOpen As Boolean
        Get
            Return _beginSucceeded
        End Get
    End Property

    ''' <summary>True bila BeginCommandGroup dipanggil tetapi gagal.</summary>
    Public ReadOnly Property BeginFailed As Boolean
        Get
            Return _beginAttempted AndAlso Not _beginSucceeded
        End Get
    End Property

    Public ReadOnly Property BeginError As String
        Get
            Return _beginError
        End Get
    End Property

    Public ReadOnly Property EndAttempted As Boolean
        Get
            Return _endAttempted
        End Get
    End Property

    Public ReadOnly Property EndFailed As Boolean
        Get
            Return _endAttempted AndAlso Not _endSucceeded
        End Get
    End Property

    Public ReadOnly Property EndError As String
        Get
            Return _endError
        End Get
    End Property

    ''' <summary>
    ''' True hanya bila group dibuka DAN ditutup dengan sukses.
    ''' Hanya dalam kondisi ini UI boleh menjanjikan "satu Ctrl+Z".
    ''' </summary>
    Public ReadOnly Property SingleUndoGuaranteed As Boolean
        Get
            Return _beginSucceeded AndAlso _endSucceeded
        End Get
    End Property

    ''' <summary>EndCommandGroup hanya boleh dipanggil bila group masih terbuka.</summary>
    Public ReadOnly Property ShouldCallEnd As Boolean
        Get
            Return _beginSucceeded AndAlso Not _endAttempted
        End Get
    End Property

    ' --- pencatatan hasil panggilan COM -------------------------------------

    Public Sub MarkBeginAttempted()
        _beginAttempted = True
    End Sub

    Public Sub MarkBeginSucceeded()
        _beginAttempted = True
        _beginSucceeded = True
        _beginError = ""
    End Sub

    Public Sub MarkBeginFailed(errorMessage As String)
        _beginAttempted = True
        _beginSucceeded = False
        _beginError = If(errorMessage, "")
    End Sub

    Public Sub MarkEndSucceeded()
        _endAttempted = True
        _endSucceeded = True
        _endError = ""
    End Sub

    Public Sub MarkEndFailed(errorMessage As String)
        _endAttempted = True
        _endSucceeded = False
        _endError = If(errorMessage, "")
    End Sub

    ''' <summary>True bila ada masalah undo yang wajib masuk laporan.</summary>
    Public ReadOnly Property HasProblem As Boolean
        Get
            Return BeginFailed OrElse EndFailed
        End Get
    End Property

    ''' <summary>Pesan siap tampil untuk laporan/peringatan.</summary>
    Public Function Describe() As String
        If BeginFailed Then
            Return "Grup undo TIDAK dapat dibuka (" & _beginError &
                   "). Perubahan tidak dapat dibatalkan dengan satu Ctrl+Z — " &
                   "batalkan per langkah bila perlu."
        End If
        If EndFailed Then
            Return "EndCommandGroup GAGAL (" & _endError &
                   "). Grup undo mungkin masih terbuka; simpan pekerjaan dan periksa " &
                   "daftar Undo CorelDRAW sebelum melanjutkan."
        End If
        If SingleUndoGuaranteed Then
            Return "Semua perubahan berada dalam satu grup undo """ & Me.Name &
                   """ — satu Ctrl+Z membatalkan seluruh operasi."
        End If
        If Not _beginAttempted Then
            Return "Grup undo tidak dipakai pada operasi ini."
        End If
        Return "Status grup undo tidak pasti (Begin berhasil, End belum dipanggil)."
    End Function
End Class
