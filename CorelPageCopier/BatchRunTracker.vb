Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  BatchRunTracker.vb
'  Pembukuan jalannya batch: berapa file direncanakan, selesai, dan berapa
'  yang dibatalkan karena pengguna menekan "Hentikan" atau karena exception.
'
'  Kenapa perlu:
'    Kode lama tidak punya jalur berhenti sama sekali dan tidak membedakan
'    "belum diproses" dari "gagal". Bila batch dihentikan/di-exception-kan,
'    file yang belum disentuh hilang dari laporan dan angka ringkasan tidak
'    menjumlah ke jumlah file yang sebenarnya.
'
'  Kelas ini murni (tanpa COM/WinForms) sehingga bisa diuji unit.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Globalization

''' <summary>Pembukuan file batch (direncanakan / selesai / dibatalkan).</summary>
Public Class BatchRunTracker

    Private ReadOnly _planned As Integer
    Private _ok As Integer = 0
    Private _partial As Integer = 0
    Private _failed As Integer = 0
    Private _cancelled As Integer = 0
    Private _stopRequested As Boolean = False

    Public Sub New(planned As Integer)
        _planned = If(planned < 0, 0, planned)
    End Sub

    ''' <summary>Jumlah file yang direncanakan sejak awal.</summary>
    Public ReadOnly Property Planned As Integer
        Get
            Return _planned
        End Get
    End Property

    Public ReadOnly Property OkCount As Integer
        Get
            Return _ok
        End Get
    End Property

    Public ReadOnly Property PartialCount As Integer
        Get
            Return _partial
        End Get
    End Property

    Public ReadOnly Property FailedCount As Integer
        Get
            Return _failed
        End Get
    End Property

    Public ReadOnly Property CancelledCount As Integer
        Get
            Return _cancelled
        End Get
    End Property

    ''' <summary>File yang sudah punya hasil final (selesai atau dibatalkan).</summary>
    Public ReadOnly Property Processed As Integer
        Get
            Return _ok + _partial + _failed + _cancelled
        End Get
    End Property

    ''' <summary>File yang belum diproses.</summary>
    Public ReadOnly Property Remaining As Integer
        Get
            Dim leftover As Integer = _planned - Processed
            Return If(leftover < 0, 0, leftover)
        End Get
    End Property

    Public ReadOnly Property StopRequested As Boolean
        Get
            Return _stopRequested
        End Get
    End Property

    ''' <summary>Minta batch berhenti setelah file yang sedang dikerjakan selesai.</summary>
    Public Sub RequestStop()
        _stopRequested = True
    End Sub

    ''' <summary>Catat satu file yang sudah selesai (status diambil dari outcome).</summary>
    Public Sub RecordCompleted(outcome As BatchFileOutcome)
        If outcome Is Nothing Then Return
        Select Case outcome.Status
            Case BatchFileStatus.Ok
                _ok += 1
            Case BatchFileStatus.PartialSuccess
                _partial += 1
            Case BatchFileStatus.Cancelled
                _cancelled += 1
            Case Else
                _failed += 1
        End Select
    End Sub

    ''' <summary>Catat sejumlah file yang tidak jadi diproses karena batch dihentikan.</summary>
    Public Sub AddCancelled(items As Integer)
        If items > 0 Then _cancelled += items
    End Sub

    ''' <summary>
    ''' Pembukuan konsisten bila jumlah file yang punya status final sama dengan
    ''' jumlah yang direncanakan (tidak ada file yang hilang dari laporan).
    ''' </summary>
    Public Function IsConsistent() As Boolean
        Return Processed = _planned
    End Function

    ''' <summary>Pesan masalah konsistensi (kosong bila konsisten).</summary>
    Public Function ConsistencyMessage() As String
        If IsConsistent() Then Return ""
        Return "pembukuan batch tidak konsisten: " & Processed.ToString(CultureInfo.InvariantCulture) &
               " file berstatus final dari " & _planned.ToString(CultureInfo.InvariantCulture) & " yang direncanakan."
    End Function

    Public Function SummaryText() As String
        Return _planned.ToString(CultureInfo.InvariantCulture) & " file: " &
               _ok.ToString(CultureInfo.InvariantCulture) & " OK, " &
               _partial.ToString(CultureInfo.InvariantCulture) & " SEBAGIAN, " &
               _failed.ToString(CultureInfo.InvariantCulture) & " GAGAL, " &
               _cancelled.ToString(CultureInfo.InvariantCulture) & " DIBATALKAN"
    End Function

    ''' <summary>
    ''' Bangun outcome "dibatalkan" untuk file yang belum sempat diproses.
    ''' Outcome ini ikut ditulis ke CSV supaya setiap file pada batch muncul
    ''' tepat satu kali dengan status yang jelas.
    ''' </summary>
    Public Shared Function BuildCancelledOutcome(filePath As String, reason As String) As BatchFileOutcome
        Dim outcome As New BatchFileOutcome(filePath)
        outcome.Cancelled = True
        outcome.Messages.Add(If(String.IsNullOrEmpty(reason), "tidak diproses (batch dihentikan).", reason))
        Return outcome
    End Function
End Class
