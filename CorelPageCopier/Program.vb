Option Explicit On
Option Strict Off
Option Infer On
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms

Module Program
    <STAThread>
    Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New CopierForm())
    End Sub
End Module

Public Class CopierForm
    Inherits Form
    ' Port UI untuk service: koneksi, copy, dan progres operasi panjang.
    Implements ICorelPrompt, ICopyPrompt, IProgressSink
    Private ReadOnly pagesList As New CheckedListBox()
    Private ReadOnly categoryBox As New ComboBox()
    Private ReadOnly scopeBox As New ComboBox()
    Private ReadOnly statusLabel As New Label()
    Private ReadOnly sourceLabel As New Label()
    Private sourcePageIndex As Integer = -1

    ' ========================================================================
    '  ARSITEKTUR (refactor aman): FORM HANYA UI
    '    koneksi + identitas dokumen  -> CorelConnectionService
    '    copy + verifikasi + laporan  -> CopyService
    '    orkestrasi Auto Select       -> AutoSelectService
    '  Tidak ada lagi COM mentah / MessageBox di dalam logika operasi; form
    '  mengimplementasikan ICorelPrompt, ICopyPrompt, dan IProgressSink.
    ' ========================================================================
    Private ReadOnly _connection As CorelConnectionService
    Private ReadOnly _copyService As CopyService
    Private ReadOnly _autoSelect As AutoSelectService

    ' ========================================================================
    '  FASE 1.5 (#3): STATUS SIBUK GLOBAL
    '  Application.DoEvents() di dalam loop panjang (paste ke banyak halaman,
    '  pindai semua halaman) tetap memproses antrean pesan Windows, sehingga
    '  user bisa menekan tombol yang sama/lain saat operasi masih berjalan.
    '  Flag busy + penonaktifan kontrol menutup celah re-entrancy itu.
    ' ========================================================================
    Private ReadOnly busyGate As New OperationGate()
    Private ReadOnly busyControlState As New Dictionary(Of Control, Boolean)()

    ''' <summary>Status sibuk — sumber kebenaran ada di OperationGate (bisa diuji unit).</summary>
    Private ReadOnly Property busy As Boolean
        Get
            Return busyGate.IsBusy
        End Get
    End Property

    Private ReadOnly Property busyOperation As String
        Get
            Return busyGate.CurrentOperation
        End Get
    End Property

    ''' <summary>
    ''' Penjaga tunggal untuk semua operasi panjang. Mengembalikan False bila masih
    ''' ada operasi berjalan sehingga handler pemanggil harus langsung Return.
    ''' </summary>
    Private Function TryBeginOperation(operationName As String) As Boolean
        If Not busyGate.TryBegin(operationName) Then
            statusLabel.Text = busyGate.BusyMessage
            Return False
        End If

        busyControlState.Clear()
        ApplyBusyState(False)
        UseWaitCursor = True
        Cursor = Cursors.WaitCursor
        statusLabel.Text = operationName & " berjalan…"
        Application.DoEvents()
        Return True
    End Function

    ''' <summary>Selalu dipanggil di Finally agar UI tidak pernah terkunci permanen.</summary>
    Private Sub EndOperation()
        busyGate.Release()
        ApplyBusyState(True)
        busyControlState.Clear()
        UseWaitCursor = False
        Cursor = Cursors.Default
    End Sub

    Private Sub ApplyBusyState(enabled As Boolean)
        ApplyBusyStateTo(Me, enabled)
    End Sub

    ''' <summary>
    ''' Nonaktifkan/aktifkan ulang kontrol interaktif. Status awal tiap kontrol
    ''' diingat, jadi kontrol yang memang sudah nonaktif tetap nonaktif setelahnya.
    ''' </summary>
    Private Sub ApplyBusyStateTo(parent As Control, enabled As Boolean)
        For Each child As Control In parent.Controls
            If TypeOf child Is Button OrElse TypeOf child Is CheckedListBox OrElse TypeOf child Is ComboBox Then
                If Not busyControlState.ContainsKey(child) Then
                    busyControlState(child) = child.Enabled
                End If
                child.Enabled = enabled AndAlso busyControlState(child)
            End If
            If child.HasChildren Then ApplyBusyStateTo(child, enabled)
        Next
    End Sub

    ''' <summary>Mencegah jendela ditutup saat operasi panjang masih berjalan.</summary>
    Private Sub CopierFormClosing(sender As Object, e As FormClosingEventArgs)
        If Not busy Then Return
        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            statusLabel.Text = "Tidak bisa ditutup: " & busyOperation & " masih berjalan. Tunggu hingga selesai."
        End If
    End Sub

    Public Sub New()
        Text = "Corel Page Copier"
        StartPosition = FormStartPosition.CenterScreen
        Size = New Size(600, 820)
        MinimumSize = New Size(520, 700)
        BackColor = Color.FromArgb(245, 247, 250)
        Font = New Font("Segoe UI", 9.0F)
        ' Service dibuat sekali; form berlaku sebagai port UI untuk semuanya.
        _connection = New CorelConnectionService(Me)
        _copyService = New CopyService(_connection, Me, Me)
        _autoSelect = New AutoSelectService(_connection, Me)
        BuildUi()
        AddHandler Shown, AddressOf FormShown
        AddHandler FormClosing, AddressOf CopierFormClosing
    End Sub

    
Private Sub BuildUi()
    SuspendLayout()
    Controls.Clear()

    ' Layout utama: header, panel salin, panel Auto Select, dan status.
    Dim root As New TableLayoutPanel With {
        .Dock = DockStyle.Fill,
        .ColumnCount = 1,
        .RowCount = 4,
        .Padding = New Padding(16),
        .BackColor = BackColor,
        .GrowStyle = TableLayoutPanelGrowStyle.FixedSize
    }

    root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
    root.RowStyles.Add(New RowStyle(SizeType.Absolute, 76.0F))
    root.RowStyles.Add(New RowStyle(SizeType.Percent, 57.0F))
    root.RowStyles.Add(New RowStyle(SizeType.Percent, 43.0F))
    root.RowStyles.Add(New RowStyle(SizeType.Absolute, 56.0F))
    Controls.Add(root)

    ' HEADER
    Dim header As New Panel With {
        .Dock = DockStyle.Fill,
        .BackColor = BackColor,
        .Margin = New Padding(0, 0, 0, 6)
    }

    Dim title As New Label With {
        .Text = "Corel Page Copier",
        .Font = New Font("Segoe UI", 18.0F, FontStyle.Bold),
        .ForeColor = Color.FromArgb(25, 35, 52),
        .AutoSize = True,
        .Location = New Point(0, 0)
    }

    Dim subtitle As New Label With {
        .Text = "Salin objek dan temukan desain antarhalaman",
        .Font = New Font("Segoe UI", 9.0F),
        .ForeColor = Color.FromArgb(102, 115, 132),
        .AutoSize = True,
        .Location = New Point(2, 34)
    }

    sourceLabel.SetBounds(2, 55, 540, 18)
    sourceLabel.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
    sourceLabel.AutoEllipsis = True
    sourceLabel.Font = New Font("Segoe UI", 8.0F)
    sourceLabel.ForeColor = Color.FromArgb(45, 100, 150)

    ' Tombol masuk ke fitur watermark. Cakupan "Halaman yang dicentang"
    ' otomatis terisi dari halaman yang sedang dicentang di daftar bawah.
    Dim wmButton As Button = CreateButton("Watermark...", False, 124)
    wmButton.Location = New Point(430, 22)
    wmButton.Anchor = AnchorStyles.Top Or AnchorStyles.Right
    AddHandler wmButton.Click, AddressOf OpenWatermarkDialog

    header.Controls.AddRange(New Control() {
        title, subtitle, sourceLabel, wmButton
    })

    root.Controls.Add(header, 0, 0)

    ' PANEL 1: COPY HALAMAN
    Dim copyCard As Panel = CreateCard()
    copyCard.Margin = New Padding(0, 0, 0, 10)

    Dim copyLayout As New TableLayoutPanel With {
        .Dock = DockStyle.Fill,
        .ColumnCount = 1,
        .RowCount = 5,
        .Padding = New Padding(14),
        .BackColor = Color.White
    }

    copyLayout.ColumnStyles.Add(
        New ColumnStyle(SizeType.Percent, 100.0F)
    )
    copyLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 27.0F))
    copyLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 32.0F))
    copyLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
    copyLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 36.0F))
    copyLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 40.0F))

    Dim copyTitle As Label = CreateSectionTitle("01  Salin ke halaman")
    copyTitle.Dock = DockStyle.Fill
    copyTitle.TextAlign = ContentAlignment.MiddleLeft
    copyLayout.Controls.Add(copyTitle, 0, 0)

    Dim hint As New Label With {
        .Text = "Pilih halaman tujuan. Halaman aktif merupakan sumber.",
        .Dock = DockStyle.Fill,
        .ForeColor = Color.FromArgb(102, 115, 132),
        .TextAlign = ContentAlignment.MiddleLeft,
        .AutoEllipsis = True,
        .Font = New Font("Segoe UI", 8.5F)
    }
    copyLayout.Controls.Add(hint, 0, 1)

    pagesList.Dock = DockStyle.Fill
    pagesList.Margin = New Padding(0, 2, 0, 8)
    pagesList.CheckOnClick = True
    pagesList.IntegralHeight = False
    pagesList.BorderStyle = BorderStyle.FixedSingle
    pagesList.BackColor = Color.White
    pagesList.Font = New Font("Segoe UI", 9.0F)
    copyLayout.Controls.Add(pagesList, 0, 2)

    Dim pageActions As New FlowLayoutPanel With {
        .Dock = DockStyle.Fill,
        .FlowDirection = FlowDirection.LeftToRight,
        .WrapContents = False,
        .AutoScroll = True,
        .BackColor = Color.White,
        .Margin = New Padding(0, 0, 0, 5),
        .Padding = New Padding(0)
    }

    Dim refresh As Button = CreateButton("Muat ulang", False, 100)
    AddHandler refresh.Click, AddressOf RefreshClicked

    Dim all As Button = CreateButton("Pilih semua", False, 100)
    AddHandler all.Click, Sub() SetAllChecks(True)

    Dim none As Button = CreateButton("Kosongkan", False, 100)
    AddHandler none.Click, Sub() SetAllChecks(False)

    pageActions.Controls.AddRange(New Control() {
        refresh, all, none
    })
    copyLayout.Controls.Add(pageActions, 0, 3)

    Dim copy As Button = CreateButton(
        "Salin ke halaman terpilih", True, 200
    )
    copy.Dock = DockStyle.Fill
    copy.Margin = New Padding(0, 2, 0, 0)
    AddHandler copy.Click, AddressOf CopyClicked
    copyLayout.Controls.Add(copy, 0, 4)

    copyCard.Controls.Add(copyLayout)
    root.Controls.Add(copyCard, 0, 1)

    ' PANEL 2: AUTO SELECT
    Dim selectCard As Panel = CreateCard()
    selectCard.Margin = New Padding(0, 0, 0, 8)

    Dim selectLayout As New TableLayoutPanel With {
        .Dock = DockStyle.Fill,
        .ColumnCount = 1,
        .RowCount = 8,
        .Padding = New Padding(14),
        .BackColor = Color.White
    }

    selectLayout.ColumnStyles.Add(
        New ColumnStyle(SizeType.Percent, 100.0F)
    )
    selectLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 27.0F))
    selectLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 19.0F))
    selectLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 31.0F))
    selectLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 19.0F))
    selectLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 31.0F))
    selectLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 37.0F))
    selectLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 35.0F))
    selectLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

    Dim selectTitle As Label = CreateSectionTitle("02  Auto Select objek")
    selectTitle.Dock = DockStyle.Fill
    selectTitle.TextAlign = ContentAlignment.MiddleLeft
    selectLayout.Controls.Add(selectTitle, 0, 0)

    Dim categoryLabel As New Label With {
        .Text = "Jenis objek",
        .Dock = DockStyle.Fill,
        .ForeColor = Color.FromArgb(70, 82, 99),
        .TextAlign = ContentAlignment.MiddleLeft
    }
    selectLayout.Controls.Add(categoryLabel, 0, 1)

    categoryBox.DropDownStyle = ComboBoxStyle.DropDownList
    categoryBox.Dock = DockStyle.Fill
    categoryBox.Margin = New Padding(0, 0, 0, 3)
    categoryBox.Items.Clear()
    categoryBox.Items.AddRange(New Object() {
        "Teks (Artistik + Paragraf)",
        "Gambar / Bitmap",
        "Garis / Kurva terbuka",
        "Objek vektor (tanpa teks/bitmap/grup)",
        "Semua objek"
    })
    categoryBox.SelectedIndex = 0
    selectLayout.Controls.Add(categoryBox, 0, 2)

    Dim scopeLabel As New Label With {
        .Text = "Cakupan pencarian",
        .Dock = DockStyle.Fill,
        .ForeColor = Color.FromArgb(70, 82, 99),
        .TextAlign = ContentAlignment.MiddleLeft
    }
    selectLayout.Controls.Add(scopeLabel, 0, 3)

    scopeBox.DropDownStyle = ComboBoxStyle.DropDownList
    scopeBox.Dock = DockStyle.Fill
    scopeBox.Margin = New Padding(0, 0, 0, 3)
    scopeBox.Items.Clear()
    scopeBox.Items.AddRange(New Object() {
        "Halaman aktif",
        "Pindai semua halaman"
    })
    scopeBox.SelectedIndex = 0
    selectLayout.Controls.Add(scopeBox, 0, 4)

    Dim autoSelect As Button = CreateButton(
        "Select otomatis", True, 200
    )
    autoSelect.Dock = DockStyle.Fill
    autoSelect.Margin = New Padding(0, 3, 0, 3)
    AddHandler autoSelect.Click, AddressOf AutoSelectClicked
    selectLayout.Controls.Add(autoSelect, 0, 5)

    Dim nextMatch As Button = CreateButton(
        "Pilih hasil berikutnya", False, 200
    )
    nextMatch.Dock = DockStyle.Fill
    nextMatch.Margin = New Padding(0, 2, 0, 3)
    AddHandler nextMatch.Click, AddressOf NextMatchingPageClicked
    selectLayout.Controls.Add(nextMatch, 0, 6)

    Dim note As New Label With {
        .Text = "Pindai semua halaman untuk menemukan objek yang cocok. " &
                "Gunakan tombol hasil berikutnya untuk berpindah halaman. " &
                "Navigasi berhenti pada hasil terakhir.",
        .Dock = DockStyle.Fill,
        .ForeColor = Color.FromArgb(102, 115, 132),
        .Font = New Font("Segoe UI", 8.0F),
        .TextAlign = ContentAlignment.MiddleLeft,
        .AutoEllipsis = True,
        .Margin = New Padding(0, 3, 0, 0)
    }
    selectLayout.Controls.Add(note, 0, 7)

    selectCard.Controls.Add(selectLayout)
    root.Controls.Add(selectCard, 0, 2)

    ' STATUS
    Dim statusCard As New Panel With {
        .Dock = DockStyle.Fill,
        .BackColor = Color.FromArgb(232, 241, 250),
        .Padding = New Padding(12, 6, 12, 6),
        .Margin = New Padding(0, 2, 0, 0)
    }

    statusLabel.Dock = DockStyle.Fill
    statusLabel.TextAlign = ContentAlignment.MiddleLeft
    statusLabel.Font = New Font("Segoe UI", 8.5F)
    statusLabel.ForeColor = Color.FromArgb(35, 71, 105)
    statusLabel.AutoEllipsis = True
    statusLabel.Text = "Status: siap. Buka dokumen di CorelDRAW."

    statusCard.Controls.Add(statusLabel)
    root.Controls.Add(statusCard, 0, 3)

    ResumeLayout(True)
End Sub


    Private Function CreateCard() As Panel
        Return New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White, .Padding = New Padding(0), .Margin = New Padding(0, 0, 0, 8), .BorderStyle = BorderStyle.FixedSingle}
    End Function

    Private Function CreateSectionTitle(caption As String) As Label
        Return New Label With {.Text = caption, .Font = New Font("Segoe UI Semibold", 11.0F, FontStyle.Bold), .ForeColor = Color.FromArgb(35, 48, 67), .AutoSize = True}
    End Function

    Private Function CreateButton(caption As String, primary As Boolean, width As Integer) As Button
        Dim b As New Button With {
            .Text = caption, .Width = width, .Height = 30, .Margin = New Padding(0, 0, 8, 0),
            .FlatStyle = FlatStyle.Flat, .UseVisualStyleBackColor = False,
            .Font = New Font("Segoe UI Semibold", 9.0F, FontStyle.Bold),
            .BackColor = If(primary, Color.FromArgb(28, 103, 190), Color.FromArgb(243, 246, 249)),
            .ForeColor = If(primary, Color.White, Color.FromArgb(47, 61, 79)),
            .Cursor = Cursors.Hand
        }
        b.FlatAppearance.BorderSize = If(primary, 0, 1)
        b.FlatAppearance.BorderColor = Color.FromArgb(220, 226, 233)
        b.FlatAppearance.MouseOverBackColor = If(primary, Color.FromArgb(22, 85, 160), Color.FromArgb(231, 237, 243))
        Return b
    End Function

    Private Sub FormShown(sender As Object, e As EventArgs)
        Try
            ConnectToCorel()
            LoadPages()
        Catch ex As Exception
            ShowError("Tidak dapat terhubung ke CorelDRAW 2022. Pastikan CorelDRAW terbuka dan dokumen CDR aktif." & Microsoft.VisualBasic.Constants.vbCrLf & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Menghubungkan ke CorelDRAW (delegasi tipis).
    ''' Seluruh logika — instance berjalan diutamakan, persetujuan sebelum membuka
    ''' instance baru, pelepasan COM yang tidak valid — ada di
    ''' CorelConnectionService sehingga bisa diuji tanpa CorelDRAW.
    ''' </summary>
    Private Sub ConnectToCorel()
        _connection.Connect()
    End Sub

    ''' <summary>
    ''' Pastikan target masih dokumen + instance yang sama seperti saat daftar
    ''' halaman dimuat (Temuan F). Pesan peringatan ditampilkan lewat
    ''' NotifyDocumentMismatch di bawah.
    ''' </summary>
    Private Function EnsureSameDocument() As Boolean
        Return _connection.EnsureSameDocument()
    End Function

    ' ========================================================================
    '  PORT UI (ICorelPrompt / ICopyPrompt / IProgressSink)
    '  Service tidak pernah memanggil MessageBox sendiri. Teks dialog di bawah
    '  SENGAJA sama persis dengan versi sebelum refactor.
    ' ========================================================================

    ''' <summary>Izin membuka instance CorelDRAW baru (jangan pernah diam-diam).</summary>
    Public Function ConfirmOpenNewInstance(detail As String) As Boolean Implements ICorelPrompt.ConfirmOpenNewInstance
        Dim question As String =
            "Tidak dapat memakai jendela CorelDRAW yang sedang berjalan." &
            Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
            "Detail: " & detail & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
            "Buka instance CorelDRAW BARU melalui COM sekarang?" & Microsoft.VisualBasic.Constants.vbCrLf &
            "PENTING: dokumen yang diproses adalah dokumen yang Anda buka di jendela BARU itu, " &
            "bukan dokumen di jendela lain."
        Return MessageBox.Show(question, "Koneksi CorelDRAW", MessageBoxButtons.YesNo,
                               MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes
    End Function

    ''' <summary>Beri tahu pengguna tentang instance baru yang sudah dibuka.</summary>
    Public Sub NotifyNewInstanceOpened(hasActiveDocument As Boolean) Implements ICorelPrompt.NotifyNewInstanceOpened
        If hasActiveDocument Then
            MessageBox.Show(
                "Instance CorelDRAW baru sudah dibuka dan dokumen aktifnya dipakai sebagai target.",
                "Koneksi CorelDRAW", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        MessageBox.Show(
            "Instance CorelDRAW BARU sudah dibuka melalui COM, tetapi belum memiliki dokumen aktif." &
            Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
            "Buka file CDR di jendela CorelDRAW yang baru terbuka, lalu klik Muat ulang.",
            "Koneksi CorelDRAW", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ''' <summary>Dokumen aktif berbeda dari dokumen saat daftar halaman dimuat (Temuan F).</summary>
    Public Sub NotifyDocumentMismatch(loaded As DocumentIdentity,
                                      current As DocumentIdentity) Implements ICorelPrompt.NotifyDocumentMismatch
        MessageBox.Show(
            "Dokumen aktif berbeda dari dokumen saat daftar halaman dimuat." &
            Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
            "Saat dimuat : " & loaded.Describe() & Microsoft.VisualBasic.Constants.vbCrLf &
            "Sekarang    : " & current.Describe() & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
            "Operasi dibatalkan agar tidak mengubah dokumen yang salah. Klik Muat ulang untuk memakai dokumen aktif.",
            "Dokumen target berubah", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        statusLabel.Text = "Dokumen aktif berubah — klik Muat ulang."
    End Sub

    ''' <summary>Konfirmasi sebelum dokumen diubah (identitas + halaman + satuan).</summary>
    Public Function ConfirmCopy(plan As CopyPlan) As Boolean Implements ICopyPrompt.ConfirmCopy
        Dim confirmText As String =
            "Dokumen     : " & plan.DocumentIdentity & Microsoft.VisualBasic.Constants.vbCrLf &
            "Halaman asal: " & plan.SourcePage & Microsoft.VisualBasic.Constants.vbCrLf &
            "Halaman tuju: " & String.Join(", ", plan.Targets.ToArray()) & Microsoft.VisualBasic.Constants.vbCrLf &
            "Satuan      : " & plan.UnitName &
            " (dinormalkan ke mm, lalu dipulihkan)" & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
            "Salin objek terpilih ke " & plan.TargetCount & " halaman? Objek sumber tidak diubah/dihapus."
        Return MessageBox.Show(confirmText, "Konfirmasi salin objek", MessageBoxButtons.YesNo,
                               MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes
    End Function

    ''' <summary>Konfirmasi lanjut tanpa grup undo (Temuan C).</summary>
    Public Function ConfirmContinueWithoutUndoGroup(reason As String) As Boolean Implements ICopyPrompt.ConfirmContinueWithoutUndoGroup
        Dim question As String =
            "Grup undo CorelDRAW tidak dapat dibuka:" & Microsoft.VisualBasic.Constants.vbCrLf &
            reason & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
            "Tanpa grup undo, satu Ctrl+Z TIDAK akan membatalkan seluruh operasi " &
            "(harus dibatalkan per halaman)." & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
            "Lanjutkan tanpa jaminan satu langkah Undo?"
        Return MessageBox.Show(question, "Grup undo gagal dibuka", MessageBoxButtons.YesNo,
                               MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes
    End Function

    ''' <summary>Satuan dokumen gagal dipulihkan (Temuan B) — wajib terlihat.</summary>
    Public Sub NotifyUnitRestoreFailed(targetUnitName As String,
                                       errorMessage As String) Implements ICopyPrompt.NotifyUnitRestoreFailed
        MessageBox.Show("Satuan dokumen gagal dipulihkan ke " & targetUnitName &
                        "." & Microsoft.VisualBasic.Constants.vbCrLf & errorMessage &
                        Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
                        "Ubah satuan dokumen kembali secara manual.",
                        "Pemulihan satuan gagal", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    ' --- IProgressSink: progres operasi panjang masuk ke status bar ----------

    ''' <summary>Pesan progres dari service (teks sama seperti versi lama).</summary>
    Public Sub Report(message As String) Implements IProgressSink.Report
        statusLabel.Text = message
    End Sub

    ''' <summary>Biarkan Windows memproses antrean pesan; aman karena busy guard aktif.</summary>
    Public Sub Pump() Implements IProgressSink.Pump
        Application.DoEvents()
    End Sub

    Private Sub LoadPages()
        ConnectToCorel()
        ' Melempar pesan yang sama seperti versi lama bila tidak ada dokumen aktif.
        _connection.EnsureDocument()
        pagesList.Items.Clear()

        ' Temuan F: simpan identitas dokumen + instance aktif. Dipakai sebelum
        ' operasi yang mengubah dokumen dan ditampilkan ke pengguna.
        Dim loadedIdentity As DocumentIdentity = _connection.RememberLoadedIdentity()

        sourcePageIndex = CInt(_connection.Document.ActivePage.Index)
        sourceLabel.Text = "Dokumen: " & loadedIdentity.Describe() &
                           " | Halaman aktif: " & sourcePageIndex &
                           " | Pilih objek sumber sebelum menjalankan Copy."
        For i As Integer = 1 To CInt(_connection.Document.Pages.Count)
            Dim caption = "Halaman " & i
            If i = sourcePageIndex Then caption &= " (AKTIF)"
            Dim itemIndex = pagesList.Items.Add(New PageItem(i, caption))
            If i = sourcePageIndex Then pagesList.SetItemChecked(itemIndex, False)
        Next
        statusLabel.Text = "Dokumen terdeteksi: " & loadedIdentity.Describe() & "."
    End Sub

    Private Sub SetAllChecks(checked As Boolean)
        For i As Integer = 0 To pagesList.Items.Count - 1
            Dim p = DirectCast(pagesList.Items(i), PageItem)
            pagesList.SetItemChecked(i, checked AndAlso p.Index <> sourcePageIndex)
        Next
    End Sub

    Private Sub RefreshClicked(sender As Object, e As EventArgs)
        If Not TryBeginOperation("Muat ulang halaman") Then Return
        Try
            RefreshClickedCore()
        Finally
            EndOperation()
        End Try
    End Sub

    Private Sub RefreshClickedCore()
        Try
            LoadPages()
        Catch ex As Exception
            ShowError("Gagal memuat ulang daftar halaman: " & ex.Message)
        End Try
    End Sub

    Private Sub CopyClicked(sender As Object, e As EventArgs)
        If Not TryBeginOperation("Salin objek") Then Return
        Try
            CopyClickedCore()
        Finally
            EndOperation()
        End Try
    End Sub

    ''' <summary>
    ''' Menyalin objek terpilih dari halaman sumber ke halaman tujuan.
    '''
    ''' Seluruh pipeline (validasi target, normalisasi satuan, grup undo, paste +
    ''' verifikasi geometri) ada di CopyService. Form hanya:
    '''   - mengumpulkan halaman yang dicentang,
    '''   - menyediakan dialog konfirmasi/peringatan lewat ICopyPrompt,
    '''   - menampilkan laporan hasil.
    ''' Objek sumber tidak pernah dihapus/diubah, dan tidak ada fallback Delete.
    ''' </summary>
    Private Sub CopyClickedCore()
        Dim targets As New List(Of Integer)()
        For Each obj As Object In pagesList.CheckedItems
            Dim p = DirectCast(obj, PageItem)
            If p.Index <> sourcePageIndex Then targets.Add(p.Index)
        Next
        If targets.Count = 0 Then
            MessageBox.Show("Centang setidaknya satu halaman tujuan selain halaman sumber.", "Corel Page Copier", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim result As CopyServiceResult = _copyService.Run(sourcePageIndex, targets.ToArray())
        RenderCopyResult(result)
    End Sub

    ''' <summary>
    ''' Terjemahkan CopyServiceResult menjadi status bar + dialog, dengan teks
    ''' yang sama seperti sebelum refactor. Setiap alasan batal punya pesannya
    ''' sendiri (tidak ada kegagalan yang ditelan).
    ''' </summary>
    Private Sub RenderCopyResult(result As CopyServiceResult)
        If result Is Nothing Then Return

        Select Case result.AbortReason
            Case CopyAbortReason.NoTargetPages
                MessageBox.Show("Centang setidaknya satu halaman tujuan selain halaman sumber.", "Corel Page Copier", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return

            Case CopyAbortReason.NoValidTargets
                MessageBox.Show("Tidak ada halaman tujuan yang valid." & Microsoft.VisualBasic.Constants.vbCrLf &
                                result.AbortDetail,
                                "Validasi tujuan", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return

            Case CopyAbortReason.NoSelection
                MessageBox.Show("Pilih objek pada halaman sumber terlebih dahulu.", "Tidak ada seleksi", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return

            Case CopyAbortReason.UnitUnreadable
                MessageBox.Show(result.AbortDetail, "Satuan dokumen tidak terbaca",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                statusLabel.Text = "Copy dibatalkan: satuan dokumen tidak terbaca."
                Return

            Case CopyAbortReason.UnitNotNormalizable
                MessageBox.Show(result.AbortDetail, "Satuan dokumen tidak dapat dinormalkan",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                statusLabel.Text = "Copy dibatalkan: satuan dokumen tidak dapat dinormalkan."
                Return

            Case CopyAbortReason.UserCancelled
                statusLabel.Text = "Copy dibatalkan oleh pengguna."
                Return

            Case CopyAbortReason.UndoGroupDeclined
                statusLabel.Text = "Copy dibatalkan: grup undo tidak tersedia."
                Return

            Case CopyAbortReason.DocumentChanged
                ' Pesan + status sudah ditampilkan oleh NotifyDocumentMismatch.
                Return

            Case CopyAbortReason.SourcePageChanged
                MessageBox.Show("Halaman aktif berubah. Klik Muat Ulang dan ulangi.", "Halaman sumber berubah", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
        End Select

        ' ---- Laporan (dibuat setelah pemulihan satuan agar kegagalannya ikut) ----
        statusLabel.Text = result.Headline
        If result.Summary.TotalCount = 0 Then
            MessageBox.Show("Copy tidak berjalan." & Microsoft.VisualBasic.Constants.vbCrLf &
                            String.Join(Microsoft.VisualBasic.Constants.vbCrLf, result.Warnings.ToArray()),
                            "Copy tidak berjalan", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        MessageBox.Show(result.BuildReport(), "Hasil Corel Page Copier", MessageBoxButtons.OK,
                        If(result.AllVerified, MessageBoxIcon.Information, MessageBoxIcon.Warning))
    End Sub

    ' ========================================================================
    '  AUTO SELECT
    '  Klasifikasi objek (ShapeClassifier), penelusuran pohon (ShapeScanner),
    '  navigasi hasil (MatchNavigator) tetap murni di modul masing-masing.
    '  Orkestrasi — pindai semua halaman, snapshot navigasi, terapkan seleksi,
    '  verifikasi jumlah terpilih — ada di AutoSelectService.
    '  Form hanya menyusun laporan & status bar.
    ' ========================================================================

    Private Sub AutoSelectClicked(sender As Object, e As EventArgs)
        If Not TryBeginOperation("Auto Select") Then Return
        Try
            AutoSelectClickedCore()
        Finally
            EndOperation()
        End Try
    End Sub

    ''' <summary>
    ''' Pilih otomatis objek berdasarkan kategori & cakupan halaman.
    ''' Aturan keselamatan yang dipertahankan:
    '''   - tidak ada objek yang dihapus/diubah,
    '''   - seleksi lama tidak dibersihkan bila tidak ada hasil,
    '''   - jumlah objek terpilih diverifikasi ulang dari CorelDRAW,
    '''   - navigasi memakai snapshot kategori/cakupan saat pemindaian.
    ''' </summary>
    Private Sub AutoSelectClickedCore()
        Dim scanTimer As Stopwatch = Stopwatch.StartNew()
        Dim category As ShapeCategory = ShapeClassifier.CategoryFromIndex(categoryBox.SelectedIndex)
        Dim allPages As Boolean = (scopeBox.SelectedIndex = 1)

        ' Sesi sebelumnya dibatalkan lebih dulu supaya kegagalan di tengah jalan
        ' tidak meninggalkan daftar halaman yang setengah jadi.
        _autoSelect.Reset()

        Try
            ConnectToCorel()
            If Not EnsureSameDocument() Then Return

            Dim pageCount As Integer = _autoSelect.DocumentPageCount()
            Dim originalPage As Integer = CInt(_connection.Document.ActivePage.Index)

            If Not allPages Then
                ScanSinglePage(originalPage, category, scanTimer)
                Return
            End If

            ' --- Pindai semua halaman (progres masuk ke status bar lewat sink) ---
            Dim scanReport As AutoSelectScanReport = _autoSelect.ScanAllPages(category)

            If Not scanReport.HasMatches Then
                ' Tidak ada hasil: JANGAN menyentuh seleksi yang ada.
                statusLabel.Text = "Auto Select: tidak ada objek cocok (" &
                                   ShapeClassifier.CategoryLabel(category) & ")."
                MessageBox.Show("Tidak ditemukan objek yang cocok untuk kategori """ &
                                ShapeClassifier.CategoryLabel(category) & """ pada " &
                                pageCount & " halaman." & Microsoft.VisualBasic.Constants.vbCrLf &
                                ShapeClassifier.CategoryHelp(category) & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
                                "Seleksi yang ada tidak diubah.",
                                "Auto Select", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim selectReport As SelectionOutcome = SelectPageAtCurrentMatch(category)

            Dim text As New StringBuilder()
            text.AppendLine("Pemindaian semua halaman selesai.")
            text.AppendLine("Kategori       : " & ShapeClassifier.CategoryLabel(category))
            text.AppendLine("Definisi       : " & ShapeClassifier.CategoryHelp(category))
            text.AppendLine("Total objek    : " & scanReport.TotalMatches & " (" & scanReport.PagesWithMatches.Count & " halaman memiliki hasil)")
            text.AppendLine("Halaman dipindai: " & scanReport.PagesScanned & " dari " & pageCount)
            text.AppendLine("Durasi         : " & scanReport.DurationSeconds.ToString("0.0") & " detik")
            text.AppendLine()
            text.AppendLine("Per halaman:")
            For Each line As String In scanReport.DescribePages()
                text.AppendLine("  " & line)
            Next
            text.AppendLine()
            text.AppendLine(selectReport.Describe())
            If scanReport.TotalDiagnostics > 0 Then
                text.AppendLine()
                text.AppendLine("Catatan pemindaian (" & scanReport.TotalDiagnostics & "):")
                text.Append(scanReport.DescribeDiagnostics())
            End If
            text.AppendLine()
            text.AppendLine("Navigasi: tombol ""Pilih hasil berikutnya"" berpindah halaman hasil " &
                            "sesuai snapshot pemindaian ini dan berhenti di hasil terakhir.")

            statusLabel.Text = "Pindai semua halaman: " & scanReport.TotalMatches & " objek; " &
                               scanReport.PagesWithMatches.Count & " halaman cocok (" & scanReport.DurationSeconds.ToString("0.0") & " detik)."
            MessageBox.Show(text.ToString(), "Auto Select — Semua Halaman", MessageBoxButtons.OK,
                            If(selectReport.SelectionVerified, MessageBoxIcon.Information, MessageBoxIcon.Warning))
        Catch ex As Exception
            _autoSelect.Reset()
            ShowError("Auto Select gagal: " & ex.Message)
        End Try
    End Sub

    ''' <summary>Cakupan "Halaman aktif": pindai satu halaman lalu pilih hasilnya.</summary>
    Private Sub ScanSinglePage(pageIndex As Integer, category As ShapeCategory, scanTimer As Stopwatch)
        Dim scan As ShapeScanResult = _autoSelect.ScanSinglePage(pageIndex, category)
        If scan Is Nothing Then Return

        If scan.MatchCount = 0 Then
            ' Tidak ada hasil -> seleksi lama dibiarkan apa adanya.
            statusLabel.Text = "Auto Select: tidak ada objek cocok pada halaman " & pageIndex.ToString() & "."
            MessageBox.Show("Tidak ditemukan objek yang cocok untuk kategori """ &
                            ShapeClassifier.CategoryLabel(category) & """ pada halaman " & pageIndex.ToString() & "." &
                            Microsoft.VisualBasic.Constants.vbCrLf &
                            ShapeClassifier.CategoryHelp(category) & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
                            "Seleksi yang ada tidak diubah.",
                            "Auto Select", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim outcome As SelectionOutcome = SelectPageAtCurrentMatch(category)

        Dim text As New StringBuilder()
        text.AppendLine("Auto Select — halaman " & pageIndex.ToString())
        text.AppendLine("Kategori : " & ShapeClassifier.CategoryLabel(category))
        text.AppendLine("Definisi : " & ShapeClassifier.CategoryHelp(category))
        text.AppendLine()
        text.AppendLine(scan.DescribeHits(20))
        text.AppendLine()
        text.AppendLine(outcome.Describe())
        Dim diagnostics As String = scan.DescribeDiagnostics()
        If Not String.IsNullOrEmpty(diagnostics) Then
            text.AppendLine()
            text.AppendLine("Catatan pemindaian:")
            text.Append(diagnostics)
        End If
        text.AppendLine()
        text.AppendLine("Durasi: " & scanTimer.Elapsed.TotalSeconds.ToString("0.0") & " detik.")

        statusLabel.Text = "Auto Select: " & outcome.SelectionVerifiedCount & " objek terpilih pada halaman " &
                           pageIndex.ToString() & " (" & scan.MatchCount & " objek cocok)."
        MessageBox.Show(text.ToString(), "Auto Select", MessageBoxButtons.OK,
                        If(outcome.SelectionVerified, MessageBoxIcon.Information, MessageBoxIcon.Warning))
    End Sub

    ''' <summary>
    ''' Pilih objek pada halaman hasil yang sedang aktif di navigator.
    ''' Kategori SNAPSHOT dipakai (bukan pilihan ComboBox saat ini).
    ''' </summary>
    Private Function SelectPageAtCurrentMatch(scanCategory As ShapeCategory) As SelectionOutcome
        Dim outcome As SelectionOutcome = _autoSelect.SelectPageAtCurrentMatch(scanCategory)
        UpdateStatusForCurrentMatch(outcome)
        Return outcome
    End Function

    ''' <summary>Status bar untuk hasil navigasi saat ini (teks sama seperti versi lama).</summary>
    Private Sub UpdateStatusForCurrentMatch(outcome As SelectionOutcome)
        If outcome Is Nothing Then Return
        Dim navigator As MatchNavigator = _autoSelect.Navigator
        If Not navigator.HasResults Then Return
        statusLabel.Text = "Hasil " & navigator.PositionLabel & " | halaman " & navigator.CurrentPage.ToString() &
                           " | " & outcome.SelectionVerifiedCount & " objek terpilih | " &
                           ShapeClassifier.CategoryLabel(navigator.Category) & "."
    End Sub

    Private Sub NextMatchingPageClicked(sender As Object, e As EventArgs)
        If Not TryBeginOperation("Pindah hasil") Then Return
        Try
            NextMatchingPageClickedCore()
        Finally
            EndOperation()
        End Try
    End Sub

    ''' <summary>
    ''' Berpindah ke halaman hasil berikutnya tanpa melewati/mengulang halaman.
    ''' Kategori & cakupan dari snapshot pemindaian; perubahan UI terdeteksi.
    ''' </summary>
    Private Sub NextMatchingPageClickedCore()
        Try
            ConnectToCorel()
            If Not EnsureSameDocument() Then Return

            Dim navigator As MatchNavigator = _autoSelect.Navigator
            If Not navigator.HasResults Then
                MessageBox.Show("Jalankan ""Select otomatis"" dengan cakupan ""Pindai semua halaman"" terlebih dahulu.",
                                "Tidak ada hasil pemindaian", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim categoryNow As ShapeCategory = ShapeClassifier.CategoryFromIndex(categoryBox.SelectedIndex)
            Dim allPagesNow As Boolean = (scopeBox.SelectedIndex = 1)

            Dim navStep As NavigationStep = _autoSelect.MoveNext(categoryNow, allPagesNow)

            Select Case navStep
                Case NavigationStep.StaleSnapshot
                    ' Kategori/cakupan di UI berubah setelah pemindaian: JANGAN
                    ' memakai daftar lama dengan aturan baru (penyebab halaman
                    ' terlewat di versi sebelumnya).
                    statusLabel.Text = "Pilihan kategori/cakupan berubah — jalankan Select otomatis lagi."
                    MessageBox.Show(
                        "Kategori atau cakupan diubah setelah pemindaian terakhir." & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
                        "Hasil pemindaian disimpan untuk: " & ShapeClassifier.CategoryLabel(navigator.Category) &
                        If(navigator.AllPages, " (semua halaman)", " (halaman aktif)") & "." & Microsoft.VisualBasic.Constants.vbCrLf &
                        "Pilihan sekarang: " & ShapeClassifier.CategoryLabel(categoryNow) &
                        If(allPagesNow, " (semua halaman)", " (halaman aktif)") & "." & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
                        "Navigasi dihentikan agar hasil tidak melompat ke halaman yang salah. " &
                        "Jalankan Select otomatis lagi untuk memindai dengan pilihan baru.",
                        "Hasil pemindaian tidak sesuai pilihan", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return

                Case NavigationStep.AtEnd
                    statusLabel.Text = "Navigasi selesai: hasil terakhir " & navigator.PositionLabel &
                                       ", halaman aktif " & CInt(_connection.Document.ActivePage.Index).ToString() & "."
                    MessageBox.Show("Anda sudah berada pada hasil terakhir (" & navigator.PositionLabel &
                                    "). Navigasi dihentikan dan tidak kembali ke hasil pertama.",
                                    "Hasil terakhir", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return

                Case NavigationStep.PageMissing
                    statusLabel.Text = "Halaman hasil tidak ditemukan lagi di dokumen."
                    MessageBox.Show("Halaman hasil berikutnya sudah tidak ada di dokumen (dokumen berubah)." &
                                    Microsoft.VisualBasic.Constants.vbCrLf &
                                    "Posisi navigasi tidak dimajukan. Jalankan Select otomatis lagi untuk memindai ulang.",
                                    "Halaman hasil hilang", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return

                Case NavigationStep.NotStarted
                    statusLabel.Text = "Belum ada hasil pemindaian."
                    Return
            End Select

            Dim outcome As SelectionOutcome = _autoSelect.LastOutcome
            UpdateStatusForCurrentMatch(outcome)
            If outcome Is Nothing OrElse Not outcome.SelectionVerified Then
                MessageBox.Show("Berpindah ke hasil " & navigator.PositionLabel & ", tetapi seleksi tidak sepenuhnya berhasil:" &
                                Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
                                If(outcome Is Nothing, "Hasil seleksi tidak tersedia.", outcome.Describe()),
                                "Seleksi parsial", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            ShowError("Gagal berpindah ke halaman cocok: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Buka dialog watermark. Halaman yang sedang dicentang diteruskan sebagai
    ''' cakupan siap pakai ("Halaman yang dicentang (N)"), jadi alur
    ''' "centang halaman -> Watermark" langsung nyambung dengan panel 01.
    ''' </summary>
    Private Sub OpenWatermarkDialog(sender As Object, e As EventArgs)
        Dim selectedPages As New List(Of Integer)()
        For Each obj As Object In pagesList.CheckedItems
            Dim p = DirectCast(obj, PageItem)
            If p.Index <> sourcePageIndex Then selectedPages.Add(p.Index)
        Next
        selectedPages.Sort()

        Try
            ConnectToCorel()
            Using dlg As New WatermarkDialog(_connection.App, selectedPages)
                dlg.ShowDialog(Me)
            End Using
            ' Dokumen mungkin berubah (halaman/layer baru) -> segarkan daftar.
            Try
                LoadPages()
            Catch
            End Try
        Catch ex As Exception
            ShowError("Fitur watermark tidak dapat dibuka: " & ex.Message)
        End Try
    End Sub

    Private Sub ShowError(message As String)
        statusLabel.Text = "Terjadi kesalahan."
        MessageBox.Show(message, "Corel Page Copier", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    Private Class PageItem
        Public ReadOnly Index As Integer
        Private ReadOnly Caption As String
        Public Sub New(pageIndex As Integer, label As String)
            Index = pageIndex
            Caption = label
        End Sub
        Public Overrides Function ToString() As String
            Return Caption
        End Function
    End Class
End Class
