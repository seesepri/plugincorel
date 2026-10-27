Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  ServicePorts.vb
'  Antarmuka (port) antara service dan UI.
'
'  Kenapa ada:
'    Sebelum refactor, logika koneksi/copy/Auto Select mencampur COM dengan
'    MessageBox dan status bar, sehingga tidak bisa diuji tanpa WinForms.
'    Service sekarang hanya bergantung pada interface di file ini; UI
'    (CopierForm) yang mengimplementasikannya dengan MessageBox/status label.
'
'  File ini murni: tidak ada COM, tidak ada WinForms, tidak ada I/O.
' ============================================================================

Imports System

''' <summary>Tujuan progress untuk operasi panjang (status bar + pump pesan Windows).</summary>
Public Interface IProgressSink
    ''' <summary>Tampilkan pesan status. Tidak boleh melempar exception.</summary>
    Sub Report(message As String)

    ''' <summary>
    ''' Beri kesempatan UI memproses antrean pesan (Application.DoEvents di UI).
    ''' Aman dipanggil dari test karena implementasi null tidak melakukan apa pun.
    ''' </summary>
    Sub Pump()
End Interface

''' <summary>Implementasi netral: dipakai test dan sebagai default aman.</summary>
Public Class NullProgressSink
    Implements IProgressSink

    Public Sub Report(message As String) Implements IProgressSink.Report
    End Sub

    Public Sub Pump() Implements IProgressSink.Pump
    End Sub
End Class

''' <summary>
''' Sink yang mengumpulkan pesan — dipakai test untuk memeriksa urutan status
''' yang dilihat pengguna tanpa WinForms.
''' </summary>
Public Class RecordingProgressSink
    Implements IProgressSink

    Private ReadOnly _messages As New System.Collections.Generic.List(Of String)()
    Private _pumpCount As Integer = 0

    Public ReadOnly Property Messages As System.Collections.Generic.List(Of String)
        Get
            Return _messages
        End Get
    End Property

    Public ReadOnly Property PumpCount As Integer
        Get
            Return _pumpCount
        End Get
    End Property

    Public ReadOnly Property LastMessage As String
        Get
            If _messages.Count = 0 Then Return ""
            Return _messages(_messages.Count - 1)
        End Get
    End Property

    Public Sub Report(message As String) Implements IProgressSink.Report
        _messages.Add(If(message, ""))
    End Sub

    Public Sub Pump() Implements IProgressSink.Pump
        _pumpCount += 1
    End Sub
End Class
