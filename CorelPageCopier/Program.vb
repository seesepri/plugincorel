Option Explicit On
Option Strict Off
Option Infer On
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing
Imports System.Runtime.InteropServices
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
    Private ReadOnly pagesList As New CheckedListBox()
    Private ReadOnly categoryBox As New ComboBox()
    Private ReadOnly scopeBox As New ComboBox()
    Private ReadOnly statusLabel As New Label()
    Private ReadOnly sourceLabel As New Label()
    Private corelApp As Object
    Private corelDoc As Object
    Private sourcePageIndex As Integer = -1
    Private Const CDR_UNIT_MM As Integer = 3
    ''' <summary>Toleransi geometri terdokumentasi (lihat CopyVerification.ToleranceMm).</summary>
    Private Const GEOMETRY_TOLERANCE_MM As Double = CopyVerification.ToleranceMm

    ' Temuan F: identitas dokumen yang terakhir dimuat. Dipakai untuk memastikan
    ' dokumen target pada saat Copy adalah dokumen (dan instance) yang sama.
    Private loadedIdentity As DocumentIdentity = Nothing

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
    ''' Menghubungkan ke CorelDRAW.
    '''
    ''' Temuan F (audit reliabilitas):
    '''   - Instance yang sedang dipakai pengguna diutamakan (ROT / GetActiveObject).
    '''   - Instance KEDUA tidak pernah dibuka diam-diam: pengguna harus menyetujui
    '''     lebih dulu, dan diberi tahu jendela CorelDRAW baru sudah terbuka.
    '''   - Objek COM yang sudah tidak valid (instance ditutup) dilepas, bukan dipakai ulang.
    ''' </summary>
    Private Sub ConnectToCorel()
        Dim lastError As Exception = Nothing

        ' (1) Pakai koneksi yang sudah berhasil sebelumnya bila masih valid.
        If corelApp IsNot Nothing Then
            If ComObjectHealth.IsAlive(corelApp) Then
                Try
                    Dim existingDoc As Object = corelApp.ActiveDocument
                    If existingDoc IsNot Nothing Then
                        corelDoc = existingDoc
                        Return
                    End If
                    lastError = New InvalidOperationException("CorelDRAW terhubung, tetapi belum ada dokumen aktif.")
                Catch ex As Exception
                    lastError = ex
                End Try
            Else
                lastError = New InvalidOperationException("Koneksi CorelDRAW sebelumnya sudah tidak valid (instance ditutup?).")
            End If
            ReleaseCorelReferences()
        End If

        ' (2) Lampirkan ke instance CorelDRAW yang SUDAH berjalan (Running Object Table).
        For Each progId As String In New String() {"CorelDRAW.Application.24", "CorelDRAW.Application"}
            Try
                Dim running As Object = Marshal.GetActiveObject(progId)
                If running IsNot Nothing Then
                    Try
                        Dim runningDoc As Object = running.ActiveDocument
                        If runningDoc IsNot Nothing Then
                            corelApp = running
                            corelDoc = runningDoc
                            Return
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
            Catch ex As Exception
                lastError = ex
            End Try
        Next

        ' (3) Tidak ada instance berjalan yang bisa dipakai: MINTA IZIN sebelum
        '     membuka instance baru (jangan diam-diam membuat instance kedua).
        Dim detail As String = If(lastError Is Nothing, "Tidak ada objek aplikasi yang tersedia.", lastError.Message)
        Dim prompt As String =
            "Tidak dapat memakai jendela CorelDRAW yang sedang berjalan." &
            Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
            "Detail: " & detail & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
            "Buka instance CorelDRAW BARU melalui COM sekarang?" & Microsoft.VisualBasic.Constants.vbCrLf &
            "PENTING: dokumen yang diproses adalah dokumen yang Anda buka di jendela BARU itu, " &
            "bukan dokumen di jendela lain."
        If MessageBox.Show(prompt, "Koneksi CorelDRAW", MessageBoxButtons.YesNo,
                           MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then
            Throw New InvalidOperationException(
                "Koneksi dibatalkan oleh pengguna. Tidak ada instance CorelDRAW baru yang dibuka. " &
                "Buka dokumen CDR di CorelDRAW lalu klik Muat ulang. Detail: " & detail, lastError)
        End If

        For Each progId As String In New String() {"CorelDRAW.Application.24", "CorelDRAW.Application"}
            Try
                Dim appType As Type = Type.GetTypeFromProgID(progId, False)
                If appType Is Nothing Then Continue For
                Dim newApp As Object = Activator.CreateInstance(appType)
                If newApp Is Nothing Then Continue For

                Try
                    newApp.Visible = True
                Catch
                End Try

                Dim newDoc As Object = Nothing
                Try
                    newDoc = newApp.ActiveDocument
                Catch
                    newDoc = Nothing
                End Try

                If newDoc IsNot Nothing Then
                    corelApp = newApp
                    corelDoc = newDoc
                    MessageBox.Show(
                        "Instance CorelDRAW baru sudah dibuka dan dokumen aktifnya dipakai sebagai target.",
                        "Koneksi CorelDRAW", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ' Instance baru terbuka tetapi belum ada dokumen: beri tahu eksplisit
                ' (jangan biarkan pengguna mengira dokumen lamanya yang diproses).
                MessageBox.Show(
                    "Instance CorelDRAW BARU sudah dibuka melalui COM, tetapi belum memiliki dokumen aktif." &
                    Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
                    "Buka file CDR di jendela CorelDRAW yang baru terbuka, lalu klik Muat ulang.",
                    "Koneksi CorelDRAW", MessageBoxButtons.OK, MessageBoxIcon.Information)
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
    End Sub

    ''' <summary>Lepas referensi COM lama supaya tidak dianggap sebagai target yang masih hidup.</summary>
    Private Sub ReleaseCorelReferences()
        If corelDoc IsNot Nothing Then ComObjectHealth.TryRelease(corelDoc)
        If corelApp IsNot Nothing Then ComObjectHealth.TryRelease(corelApp)
        corelDoc = Nothing
        corelApp = Nothing
    End Sub

    ''' <summary>Identitas instance + dokumen aktif saat ini (late bound, aman gagal).</summary>
    Private Function CurrentIdentity() As DocumentIdentity
        Return DocumentIdentity.FromDocument(corelDoc, DocumentIdentity.InstanceTokenFor(corelApp))
    End Function

    ''' <summary>
    ''' Pastikan target masih dokumen yang sama seperti saat daftar halaman dimuat.
    ''' Temuan F: mencegah dokumen dari instance lain dianggap sebagai target yang sama.
    ''' </summary>
    Private Function EnsureSameDocument() As Boolean
        If loadedIdentity Is Nothing Then
            loadedIdentity = CurrentIdentity()
            Return True
        End If
        Dim current As DocumentIdentity = CurrentIdentity()
        If current.SameTarget(loadedIdentity) Then
            ' Jumlah halaman boleh berubah (mis. halaman ditambah), cukup perbarui.
            loadedIdentity = current
            Return True
        End If

        MessageBox.Show(
            "Dokumen aktif berbeda dari dokumen saat daftar halaman dimuat." &
            Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
            "Saat dimuat : " & loadedIdentity.Describe() & Microsoft.VisualBasic.Constants.vbCrLf &
            "Sekarang    : " & current.Describe() & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
            "Operasi dibatalkan agar tidak mengubah dokumen yang salah. Klik Muat ulang untuk memakai dokumen aktif.",
            "Dokumen target berubah", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        statusLabel.Text = "Dokumen aktif berubah — klik Muat ulang."
        Return False
    End Function

    Private Sub LoadPages()
        ConnectToCorel()
        pagesList.Items.Clear()
        corelDoc = corelApp.ActiveDocument
        If corelDoc Is Nothing Then
            Throw New InvalidOperationException("Tidak ada dokumen aktif di CorelDRAW. Buka file CDR lalu klik Muat ulang.")
        End If

        ' Temuan F: simpan identitas dokumen + instance aktif. Dipakai sebelum
        ' operasi yang mengubah dokumen dan ditampilkan ke pengguna.
        loadedIdentity = CurrentIdentity()

        sourcePageIndex = CInt(corelDoc.ActivePage.Index)
        sourceLabel.Text = "Dokumen: " & loadedIdentity.Describe() &
                           " | Halaman aktif: " & sourcePageIndex &
                           " | Pilih objek sumber sebelum menjalankan Copy."
        For i As Integer = 1 To CInt(corelDoc.Pages.Count)
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
    ''' Audit reliabilitas — yang diperbaiki di sini:
    '''   A. Verifikasi copy: status dipisah Verified / Partial / Failed, geometri
    '''      sumber dibandingkan dengan OBJEK HASIL PASTE (z-order atas atau
    '''      selection setelah paste), bukan shape terbawah pada layer.
    '''   B. Normalisasi satuan: satuan asli wajib terbaca; bila penyetelan mm
    '''      gagal dipakai konversi eksplisit; pemulihan di Finally dan
    '''      kegagalannya dilaporkan.
    '''   C. Undo grouping: kegagalan BeginCommandGroup dilaporkan dan pengguna
    '''      diminta konfirmasi; EndCommandGroup hanya dipanggil bila group terbuka.
    '''   F. Identitas dokumen + jumlah halaman ditampilkan sebelum mengubah dokumen.
    ''' Objek sumber tidak pernah dihapus/diubah, dan tidak ada fallback Delete.
    ''' </summary>
    Private Sub CopyClickedCore()
        Dim stopwatch As Stopwatch = Stopwatch.StartNew()

        Dim targets As New List(Of Integer)()
        For Each obj As Object In pagesList.CheckedItems
            Dim p = DirectCast(obj, PageItem)
            If p.Index <> sourcePageIndex Then targets.Add(p.Index)
        Next
        If targets.Count = 0 Then
            MessageBox.Show("Centang setidaknya satu halaman tujuan selain halaman sumber.", "Corel Page Copier", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        ' Status dokumen yang wajib dipulihkan (Temuan B & C).
        Dim unitScope As New DocumentUnitScope(CDR_UNIT_MM)
        Dim undoGuard As New UndoGroupGuard("Salin objek ke halaman")
        Dim summary As New CopyRunSummary()
        Dim sourcePageToRestore As Integer = -1
        Dim fatalError As Exception = Nothing
        Dim identityText As String = ""

        Try
            ConnectToCorel()
            If Not EnsureSameDocument() Then Return
            identityText = loadedIdentity.Describe()

            Dim actualSourcePage As Integer = CInt(corelDoc.ActivePage.Index)
            If actualSourcePage <> sourcePageIndex Then
                MessageBox.Show("Halaman aktif berubah. Klik Muat Ulang dan ulangi.", "Halaman sumber berubah", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            sourcePageToRestore = actualSourcePage

            Dim documentPageCount As Integer = CInt(corelDoc.Pages.Count)
            Dim validatedTargets As New List(Of Integer)()
            For Each pageIndex In targets
                If pageIndex < 1 OrElse pageIndex > documentPageCount Then
                    summary.Warnings.Add("Halaman " & pageIndex & ": indeks halaman tidak valid — dilewati.")
                ElseIf pageIndex = actualSourcePage Then
                    summary.Warnings.Add("Halaman " & pageIndex & ": halaman sumber tidak boleh menjadi tujuan — dilewati.")
                ElseIf Not validatedTargets.Contains(pageIndex) Then
                    validatedTargets.Add(pageIndex)
                End If
            Next
            If validatedTargets.Count = 0 Then
                MessageBox.Show("Tidak ada halaman tujuan yang valid." & Microsoft.VisualBasic.Constants.vbCrLf &
                                String.Join(Microsoft.VisualBasic.Constants.vbCrLf, summary.Warnings.ToArray()),
                                "Validasi tujuan", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' --- Seleksi: jumlah objek dibaca dengan status eksplisit ---------
            ' Review kedua: `CInt(selection.Count)` membuat kegagalan membaca
            ' Count berubah menjadi exception umum ("Operasi dihentikan oleh
            ' kesalahan"), dan kegagalan COM saat membaca seleksi dilaporkan
            ' sebagai "tidak ada seleksi". Keduanya menyesatkan. Sekarang:
            '   - seleksi tidak terbaca      -> dibatalkan dengan pesan COM jelas,
            '   - Count terbaca & 0          -> "tidak ada seleksi",
            '   - Count gagal dibaca         -> dibatalkan SEBELUM dokumen diubah
            '                                  (jumlah objek tidak boleh ditebak),
            '   - tidak punya Count (Shape)  -> diteruskan; pengukuran yang
            '                                  menandai Partial bila tak terbukti.
            Dim selection As Object = Nothing
            Dim selectionReadError As String = ""
            Try
                selection = corelApp.ActiveSelectionRange
            Catch ex As Exception
                selection = Nothing
                selectionReadError = ex.GetType().Name & ": " & ex.Message
            End Try

            If selection Is Nothing Then
                If selectionReadError <> "" Then
                    MessageBox.Show("Seleksi tidak dapat dibaca dari CorelDRAW:" &
                                    Microsoft.VisualBasic.Constants.vbCrLf & selectionReadError &
                                    Microsoft.VisualBasic.Constants.vbCrLf &
                                    Microsoft.VisualBasic.Constants.vbCrLf &
                                    "Copy dibatalkan; dokumen tidak diubah.",
                                    "Seleksi tidak dapat dibaca", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    statusLabel.Text = "Copy dibatalkan: seleksi tidak dapat dibaca."
                Else
                    MessageBox.Show("Pilih objek pada halaman sumber terlebih dahulu.", "Tidak ada seleksi", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
                Return
            End If

            Dim selectionCount As CountReadResult = CopyMeasurement.TryReadCount(selection)
            If selectionCount.IsOk AndAlso selectionCount.Count = 0 Then
                MessageBox.Show("Pilih objek pada halaman sumber terlebih dahulu.", "Tidak ada seleksi", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If Not selectionCount.IsOk AndAlso Not selectionCount.IsMemberMissing Then
                MessageBox.Show("Jumlah objek seleksi tidak dapat diverifikasi:" &
                                Microsoft.VisualBasic.Constants.vbCrLf & selectionCount.Describe() &
                                Microsoft.VisualBasic.Constants.vbCrLf &
                                Microsoft.VisualBasic.Constants.vbCrLf &
                                "Copy dibatalkan sebelum dokumen diubah, karena hasil tidak akan bisa diverifikasi.",
                                "Seleksi tidak dapat diverifikasi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                statusLabel.Text = "Copy dibatalkan: jumlah objek seleksi tidak dapat diverifikasi."
                Return
            End If

            ' --- Temuan B: baca satuan asli SEBELUM menyentuh dokumen ----------
            Try
                unitScope.MarkOriginalKnown(CInt(corelDoc.Unit))
            Catch ex As Exception
                unitScope.MarkOriginalUnknown(ex.Message)
            End Try
            If Not unitScope.OriginalUnitKnown Then
                MessageBox.Show(unitScope.BlockingReason, "Satuan dokumen tidak terbaca",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                statusLabel.Text = "Copy dibatalkan: satuan dokumen tidak terbaca."
                Return
            End If

            ' --- Temuan F: identitas dokumen + halaman ditampilkan dulu --------
            Dim confirmText As String =
                "Dokumen     : " & identityText & Microsoft.VisualBasic.Constants.vbCrLf &
                "Halaman asal: " & actualSourcePage & Microsoft.VisualBasic.Constants.vbCrLf &
                "Halaman tuju: " & String.Join(", ", validatedTargets.ToArray()) & Microsoft.VisualBasic.Constants.vbCrLf &
                "Satuan      : " & DocumentUnits.UnitName(unitScope.OriginalUnit) &
                " (dinormalkan ke mm, lalu dipulihkan)" & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
                "Salin objek terpilih ke " & validatedTargets.Count & " halaman? Objek sumber tidak diubah/dihapus."
            If MessageBox.Show(confirmText, "Konfirmasi salin objek", MessageBoxButtons.YesNo,
                               MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then
                statusLabel.Text = "Copy dibatalkan oleh pengguna."
                Return
            End If

            ' --- Temuan B: normalisasi satuan --------------------------------
            Try
                corelDoc.Unit = unitScope.TargetUnit
                unitScope.MarkChangeApplied()
            Catch ex As Exception
                unitScope.MarkChangeFailed(ex.Message)
            End Try
            If Not unitScope.CanVerifyInMillimeters Then
                MessageBox.Show(unitScope.BlockingReason, "Satuan dokumen tidak dapat dinormalkan",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                statusLabel.Text = "Copy dibatalkan: satuan dokumen tidak dapat dinormalkan."
                Return
            End If
            Dim unitFactor As Double = unitScope.MeasurementFactorMm

            ' Ukur objek SUMBER (gabungan shape range) sebagai acuan verifikasi.
            ' Konversi ke mm dilakukan di CopyVerification.VerifyPasted.
            Dim sourceSet As MeasuredShapeSet = CopyMeasurement.MeasureSelection(selection)

            ' --- Temuan C: grup undo ----------------------------------------
            undoGuard.MarkBeginAttempted()
            Try
                corelDoc.BeginCommandGroup(undoGuard.Name)
                undoGuard.MarkBeginSucceeded()
            Catch ex As Exception
                undoGuard.MarkBeginFailed(ex.Message)
            End Try
            If undoGuard.BeginFailed Then
                Dim question As String =
                    "Grup undo CorelDRAW tidak dapat dibuka:" & Microsoft.VisualBasic.Constants.vbCrLf &
                    undoGuard.BeginError & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
                    "Tanpa grup undo, satu Ctrl+Z TIDAK akan membatalkan seluruh operasi " &
                    "(harus dibatalkan per halaman)." & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
                    "Lanjutkan tanpa jaminan satu langkah Undo?"
                If MessageBox.Show(question, "Grup undo gagal dibuka", MessageBoxButtons.YesNo,
                                   MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then
                    summary.Warnings.Add(undoGuard.Describe())
                    statusLabel.Text = "Copy dibatalkan: grup undo tidak tersedia."
                    Return
                End If
                summary.Warnings.Add(undoGuard.Describe())
            End If

            ' Objek sumber TIDAK PERNAH dihapus sebagai bagian dari validasi.
            selection.Copy()

            For targetPosition As Integer = 0 To validatedTargets.Count - 1
                Dim pageIndex As Integer = validatedTargets(targetPosition)
                statusLabel.Text = "Copy progres: " & (targetPosition + 1) & "/" & validatedTargets.Count &
                                   " — halaman tujuan " & pageIndex & "."
                Application.DoEvents()   ' aman: flag busy + kontrol nonaktif mencegah re-entrancy
                Try
                    Dim destination As Object = corelDoc.Pages(pageIndex)
                    destination.Activate()
                    Dim layer As Object = destination.ActiveLayer
                    Dim layerShapes As Object = layer.Shapes

                    ' Review kedua: jumlah objek layer dibaca dengan status
                    ' eksplisit. Kegagalan membaca Count TIDAK dianggap 0 (dan
                    ' tidak boleh menghasilkan "paste tidak menambah objek" yang
                    ' menyesatkan maupun exception mentah).
                    Dim beforeRead As CountReadResult = CopyMeasurement.TryReadCount(layerShapes)
                    If Not beforeRead.IsOk Then
                        summary.Add(pageIndex, CopyVerifyStatus.Failed,
                                    "jumlah objek pada layer tujuan tidak dapat dibaca sebelum paste (" &
                                    beforeRead.Describe() & "); halaman dilewati, tidak ada klaim berhasil.")
                        Continue For
                    End If

                    layer.Paste()

                    Dim afterRead As CountReadResult = CopyMeasurement.TryReadCount(layerShapes)
                    If Not afterRead.IsOk Then
                        ' Paste sudah dijalankan tetapi hasilnya tak terhitung:
                        ' itu "paste terjadi, verifikasi tidak lengkap" = Partial,
                        ' bukan sukses dan bukan pula klaim bahwa paste gagal.
                        summary.Add(pageIndex, CopyVerifyStatus.PartiallyVerified,
                                    "jumlah objek pada layer tujuan tidak dapat dibaca setelah paste (" &
                                    afterRead.Describe() & "); hasil paste tidak dapat diverifikasi.")
                        Continue For
                    End If

                    Dim added As Integer = afterRead.Count - beforeRead.Count

                    Dim pastedSelection As Object = Nothing
                    Try
                        pastedSelection = corelApp.ActiveSelectionRange
                    Catch
                        pastedSelection = Nothing
                    End Try

                    ' Temuan A: ukur objek hasil paste yang BENAR (bukan shape terbawah),
                    ' lalu verifikasi lewat satu fungsi yang juga diuji unit.
                    Dim pastedSet As MeasuredShapeSet = CopyMeasurement.MeasurePasted(pastedSelection, layerShapes, added)
                    Dim verification As CopyVerificationResult = CopyVerification.VerifyPasted(
                        sourceSet, pastedSet, added, unitFactor, GEOMETRY_TOLERANCE_MM)
                    summary.Add(pageIndex, verification.Status, verification.Reason)
                Catch pageEx As Exception
                    summary.Add(pageIndex, CopyVerifyStatus.Failed,
                                pageEx.GetType().Name & ": " & pageEx.Message)
                End Try
            Next
        Catch ex As Exception
            fatalError = ex
        Finally
            ' --- Temuan C: EndCommandGroup HANYA bila group benar-benar terbuka ---
            If undoGuard.ShouldCallEnd Then
                Try
                    corelDoc.EndCommandGroup()
                    undoGuard.MarkEndSucceeded()
                Catch ex As Exception
                    undoGuard.MarkEndFailed(ex.Message)
                End Try
            End If

            ' --- Temuan B: pulihkan satuan asli, laporkan bila gagal ------------
            If unitScope.RestoreRequired Then
                Try
                    corelDoc.Unit = unitScope.OriginalUnit
                    unitScope.MarkRestored()
                Catch ex As Exception
                    unitScope.MarkRestoreFailed(ex.Message)
                End Try
            End If

            If sourcePageToRestore > 0 Then
                Try
                    corelDoc.Pages(sourcePageToRestore).Activate()
                Catch ex As Exception
                    summary.Warnings.Add("Halaman sumber tidak dapat diaktifkan kembali: " & ex.Message)
                End Try
            End If
        End Try

        ' ---- Laporan dibuat SETELAH Finally supaya kegagalan pemulihan ikut
        ' ---- dilaporkan (Temuan B & C) dan tidak mungkin disembunyikan.
        If undoGuard.HasProblem Then summary.Warnings.Add(undoGuard.Describe())
        If unitScope.RestoreFailed Then
            summary.Warnings.Add("PEMULIHAN SATUAN DOKUMEN GAGAL (" & unitScope.RestoreError &
                                 "). Satuan dokumen kemungkinan tetap mm — ubah manual di CorelDRAW.")
            MessageBox.Show("Satuan dokumen gagal dipulihkan ke " & DocumentUnits.UnitName(unitScope.OriginalUnit) &
                            "." & Microsoft.VisualBasic.Constants.vbCrLf & unitScope.RestoreError &
                            Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
                            "Ubah satuan dokumen kembali secara manual.",
                            "Pemulihan satuan gagal", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        Dim elapsed As Double = stopwatch.Elapsed.TotalSeconds

        If fatalError IsNot Nothing Then
            summary.Warnings.Add("Operasi dihentikan oleh kesalahan: " & fatalError.GetType().Name & ": " & fatalError.Message)
        End If

        statusLabel.Text = summary.Headline(elapsed)
        If summary.TotalCount = 0 Then
            MessageBox.Show("Copy tidak berjalan." & Microsoft.VisualBasic.Constants.vbCrLf &
                            String.Join(Microsoft.VisualBasic.Constants.vbCrLf, summary.Warnings.ToArray()),
                            "Copy tidak berjalan", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim report As String = summary.BuildReport(identityText, unitScope.Describe(), undoGuard.Describe(), elapsed)
        MessageBox.Show(report, "Hasil Corel Page Copier", MessageBoxButtons.OK,
                        If(summary.AllVerified, MessageBoxIcon.Information, MessageBoxIcon.Warning))
    End Sub

    ' ========================================================================
    '  AUTO SELECT
    '  Klasifikasi objek dipisah ke ShapeClassifier (murni), penelusuran pohon
    '  ke ShapeScanner (murni), navigasi hasil ke MatchNavigator (murni).
    '  Bagian di bawah ini hanya menjembatani COM + UI.
    ' ========================================================================

    Private ReadOnly matchNavigator As New MatchNavigator()

    ''' <summary>Hasil pemindaian terakhir per halaman (untuk laporan & navigasi).</summary>
    Private ReadOnly lastScanByPage As New Dictionary(Of Integer, ShapeScanResult)()

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
    '''
    ''' Audit Auto Select — yang diperbaiki:
    '''   1. Kategori "vektor" tidak lagi "apa pun yang bukan teks/bitmap/grup"
    '''      (dulu ikut memilih garis bantu/OLE); definisinya eksplisit di
    '''      ShapeClassifier.
    '''   2. Status kurva dibaca dari Curve.Closed; bila tidak terbaca objek TIDAK
    '''      dianggap kurva terbuka dan kegagalannya dilaporkan.
    '''   3. Grup bersarang: batas kedalaman + kegagalan baca anak dicatat.
    '''   4. PowerClip: isi container dipindai lewat jalur khusus; hit di dalamnya
    '''      memilih container (bukan objek tersembunyi yang tak bisa dipilih).
    '''   5. Watermark tidak pernah terpilih (termasuk di dalam grup/PowerClip).
    '''   6. Navigasi memakai snapshot kategori/cakupan saat pemindaian.
    '''   7. Seleksi kosong: seleksi lama TIDAK dibersihkan (tidak ada operasi
    '''      destruktif saat tidak ada hasil).
    '''   8. Jumlah objek terpilih divalidasi ulang dari CorelDRAW.
    ''' </summary>
    Private Sub AutoSelectClickedCore()
        Dim scanTimer As Stopwatch = Stopwatch.StartNew()
        Dim category As ShapeCategory = ShapeClassifier.CategoryFromIndex(categoryBox.SelectedIndex)
        Dim allPages As Boolean = (scopeBox.SelectedIndex = 1)

        ' Sesi sebelumnya dibatalkan lebih dulu supaya kegagalan di tengah jalan
        ' tidak meninggalkan daftar halaman yang setengah jadi.
        matchNavigator.Reset()
        lastScanByPage.Clear()

        Try
            ConnectToCorel()
            If Not EnsureSameDocument() Then Return

            Dim pageCount As Integer = CInt(corelDoc.Pages.Count)
            Dim originalPage As Integer = CInt(corelDoc.ActivePage.Index)

            If Not allPages Then
                ScanSinglePage(originalPage, category, scanTimer)
                Return
            End If

            ' --- Pindai semua halaman --------------------------------
            Dim pagesWithMatches As New List(Of Integer)()
            Dim totalMatches As Integer = 0
            Dim totalDiagnostics As Integer = 0

            For pageIndex As Integer = 1 To pageCount
                statusLabel.Text = "Memindai progres: halaman " & pageIndex & "/" & pageCount &
                                   " | total objek cocok sementara: " & totalMatches & "."
                Application.DoEvents()   ' aman: flag busy + kontrol nonaktif mencegah re-entrancy

                Dim scan As ShapeScanResult = ScanPageForUi(pageIndex, category)
                If scan Is Nothing Then Continue For

                lastScanByPage(pageIndex) = scan
                totalMatches += scan.MatchCount
                totalDiagnostics += scan.Diagnostics.Count
                If scan.MatchCount > 0 Then pagesWithMatches.Add(pageIndex)
            Next

            If pagesWithMatches.Count = 0 Then
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

            ' Snapshot kategori + cakupan ikut disimpan untuk navigasi.
            matchNavigator.Start(category, allPages, pagesWithMatches)

            Dim selectReport As SelectionOutcome = SelectPageAtCurrentMatch(category)

            Dim details As New List(Of String)()
            For Each pageIndex As Integer In pagesWithMatches
                Dim scan As ShapeScanResult = lastScanByPage(pageIndex)
                details.Add("Halaman " & pageIndex.ToString() & ": " & scan.MatchCount.ToString() &
                            " objek (" & scan.DistinctTargetCount.ToString() & " target seleksi)")
            Next
            For Each pageIndex As Integer In lastScanByPage.Keys
                If Not pagesWithMatches.Contains(pageIndex) Then
                    details.Add("Halaman " & pageIndex.ToString() & ": 0 objek")
                End If
            Next
            details.Sort()

            Dim report As New StringBuilder()
            report.AppendLine("Pemindaian semua halaman selesai.")
            report.AppendLine("Kategori       : " & ShapeClassifier.CategoryLabel(category))
            report.AppendLine("Definisi       : " & ShapeClassifier.CategoryHelp(category))
            report.AppendLine("Total objek    : " & totalMatches & " (" & pagesWithMatches.Count & " halaman memiliki hasil)")
            report.AppendLine("Halaman dipindai: " & lastScanByPage.Count & " dari " & pageCount)
            report.AppendLine("Durasi         : " & scanTimer.Elapsed.TotalSeconds.ToString("0.0") & " detik")
            report.AppendLine()
            report.AppendLine("Per halaman:")
            For Each line As String In details
                report.AppendLine("  " & line)
            Next
            report.AppendLine()
            report.AppendLine(selectReport.Describe())
            If totalDiagnostics > 0 Then
                report.AppendLine()
                report.AppendLine("Catatan pemindaian (" & totalDiagnostics & "):")
                For Each pageIndex As Integer In lastScanByPage.Keys
                    Dim diagnostics As String = lastScanByPage(pageIndex).DescribeDiagnostics()
                    If Not String.IsNullOrEmpty(diagnostics) Then
                        report.AppendLine("  Halaman " & pageIndex.ToString() & ":")
                        report.Append(diagnostics)
                    End If
                Next
            End If
            report.AppendLine()
            report.AppendLine("Navigasi: tombol ""Pilih hasil berikutnya"" berpindah halaman hasil " &
                              "sesuai snapshot pemindaian ini dan berhenti di hasil terakhir.")

            statusLabel.Text = "Pindai semua halaman: " & totalMatches & " objek; " &
                               pagesWithMatches.Count & " halaman cocok (" & scanTimer.Elapsed.TotalSeconds.ToString("0.0") & " detik)."
            MessageBox.Show(report.ToString(), "Auto Select — Semua Halaman", MessageBoxButtons.OK,
                            If(selectReport.SelectionVerified, MessageBoxIcon.Information, MessageBoxIcon.Warning))
        Catch ex As Exception
            matchNavigator.Reset()
            lastScanByPage.Clear()
            ShowError("Auto Select gagal: " & ex.Message)
        End Try
    End Sub

    ''' <summary>Cakupan "Halaman aktif": pindai satu halaman lalu pilih hasilnya.</summary>
    Private Sub ScanSinglePage(pageIndex As Integer, category As ShapeCategory, scanTimer As Stopwatch)
        Dim scan As ShapeScanResult = ScanPageForUi(pageIndex, category)
        If scan Is Nothing Then Return
        lastScanByPage(pageIndex) = scan

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

        matchNavigator.Start(category, False, New Integer() {pageIndex})
        Dim outcome As SelectionOutcome = SelectPageAtCurrentMatch(category)

        Dim report As New StringBuilder()
        report.AppendLine("Auto Select — halaman " & pageIndex.ToString())
        report.AppendLine("Kategori : " & ShapeClassifier.CategoryLabel(category))
        report.AppendLine("Definisi : " & ShapeClassifier.CategoryHelp(category))
        report.AppendLine()
        report.AppendLine(scan.DescribeHits(20))
        report.AppendLine()
        report.AppendLine(outcome.Describe())
        Dim diagnostics As String = scan.DescribeDiagnostics()
        If Not String.IsNullOrEmpty(diagnostics) Then
            report.AppendLine()
            report.AppendLine("Catatan pemindaian:")
            report.Append(diagnostics)
        End If
        report.AppendLine()
        report.AppendLine("Durasi: " & scanTimer.Elapsed.TotalSeconds.ToString("0.0") & " detik.")

        statusLabel.Text = "Auto Select: " & outcome.SelectionVerifiedCount & " objek terpilih pada halaman " &
                           pageIndex.ToString() & " (" & scan.MatchCount & " objek cocok)."
        MessageBox.Show(report.ToString(), "Auto Select", MessageBoxButtons.OK,
                        If(outcome.SelectionVerified, MessageBoxIcon.Information, MessageBoxIcon.Warning))
    End Sub

    ''' <summary>
    ''' Pindai satu halaman lewat adapter COM. Mengembalikan Nothing bila objek
    ''' halaman tidak dapat diakses (dicatat ke status).
    ''' </summary>
    Private Function ScanPageForUi(pageIndex As Integer, category As ShapeCategory) As ShapeScanResult
        Dim page As Object
        Try
            page = corelDoc.Pages(pageIndex)
        Catch ex As Exception
            statusLabel.Text = "Halaman " & pageIndex.ToString() & " tidak dapat diakses: " & ex.Message
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
    ''' Pilih objek pada halaman hasil yang sedang aktif di navigator.
    ''' Selalu memakai kategori SNAPSHOT, bukan pilihan ComboBox saat ini.
    ''' </summary>
    Private Function SelectPageAtCurrentMatch(scanCategory As ShapeCategory) As SelectionOutcome
        If Not matchNavigator.HasResults Then Return SelectionOutcome.NotRun("Belum ada hasil pemindaian.")

        Dim pageIndex As Integer = matchNavigator.CurrentPage
        If pageIndex <= 0 Then Return SelectionOutcome.NotRun("Posisi hasil tidak valid.")

        Dim scan As ShapeScanResult = Nothing
        If Not lastScanByPage.TryGetValue(pageIndex, scan) Then
            ' Hasil lama tidak tersedia (mis. dokumen berubah): pindai ulang
            ' dengan kategori snapshot agar jumlahnya konsisten.
            scan = ScanPageForUi(pageIndex, scanCategory)
            If scan Is Nothing Then Return SelectionOutcome.NotRun("Halaman " & pageIndex.ToString() & " tidak dapat diakses.")
            lastScanByPage(pageIndex) = scan
        End If

        Try
            corelDoc.Pages(pageIndex).Activate()
        Catch ex As Exception
            Return SelectionOutcome.NotRun("Halaman " & pageIndex.ToString() & " tidak dapat diaktifkan: " & ex.Message)
        End Try

        Dim outcome As SelectionOutcome = ApplySelection(scan, expectedPageIndex:=pageIndex)
        statusLabel.Text = "Hasil " & matchNavigator.PositionLabel & " | halaman " & pageIndex.ToString() &
                           " | " & outcome.SelectionVerifiedCount & " objek terpilih | " &
                           ShapeClassifier.CategoryLabel(scanCategory) & "."
        Return outcome
    End Function

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

            If Not matchNavigator.HasResults Then
                MessageBox.Show("Jalankan ""Select otomatis"" dengan cakupan ""Pindai semua halaman"" terlebih dahulu.",
                                "Tidak ada hasil pemindaian", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim categoryNow As ShapeCategory = ShapeClassifier.CategoryFromIndex(categoryBox.SelectedIndex)
            Dim allPagesNow As Boolean = (scopeBox.SelectedIndex = 1)

            Dim pageCount As Integer = CInt(corelDoc.Pages.Count)
            ' 'Step' adalah kata kunci VB -> pakai navStep. Lambda dijadikan variabel
            ' agar tidak bergantung pada parsing lanjutan baris.
            Dim pageExists As Func(Of Integer, Boolean) =
                Function(pageIdx As Integer) pageIdx >= 1 AndAlso pageIdx <= pageCount
            Dim navStep As NavigationStep = matchNavigator.MoveNext(categoryNow, allPagesNow, pageExists)

            Select Case navStep
                Case NavigationStep.StaleSnapshot
                    ' Kategori/cakupan di UI berubah setelah pemindaian: JANGAN
                    ' memakai daftar lama dengan aturan baru (penyebab halaman
                    ' terlewat di versi sebelumnya).
                    statusLabel.Text = "Pilihan kategori/cakupan berubah — jalankan Select otomatis lagi."
                    MessageBox.Show(
                        "Kategori atau cakupan diubah setelah pemindaian terakhir." & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
                        "Hasil pemindaian disimpan untuk: " & ShapeClassifier.CategoryLabel(matchNavigator.Category) &
                        If(matchNavigator.AllPages, " (semua halaman)", " (halaman aktif)") & "." & Microsoft.VisualBasic.Constants.vbCrLf &
                        "Pilihan sekarang: " & ShapeClassifier.CategoryLabel(categoryNow) &
                        If(allPagesNow, " (semua halaman)", " (halaman aktif)") & "." & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
                        "Navigasi dihentikan agar hasil tidak melompat ke halaman yang salah. " &
                        "Jalankan Select otomatis lagi untuk memindai dengan pilihan baru.",
                        "Hasil pemindaian tidak sesuai pilihan", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return

                Case NavigationStep.AtEnd
                    statusLabel.Text = "Navigasi selesai: hasil terakhir " & matchNavigator.PositionLabel &
                                       ", halaman aktif " & CInt(corelDoc.ActivePage.Index).ToString() & "."
                    MessageBox.Show("Anda sudah berada pada hasil terakhir (" & matchNavigator.PositionLabel &
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

            Dim outcome As SelectionOutcome = SelectPageAtCurrentMatch(matchNavigator.Category)
            If Not outcome.SelectionVerified Then
                MessageBox.Show("Berpindah ke hasil " & matchNavigator.PositionLabel & ", tetapi seleksi tidak sepenuhnya berhasil:" &
                                Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
                                outcome.Describe(), "Seleksi parsial", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            ShowError("Gagal berpindah ke halaman cocok: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Terapkan seleksi untuk hasil pemindaian satu halaman.
    ''' Aturan keselamatan:
    '''   - Tidak ada target -> seleksi lama TIDAK disentuh.
    '''   - Tidak ada fallback Delete; kegagalan ClearSelection membatalkan operasi.
    '''   - Jumlah hasil dibaca ulang dari CorelDRAW (bukan sekadar hitungan
    '''     pemanggilan AddToSelection yang tidak melempar).
    ''' </summary>
    Private Function ApplySelection(scan As ShapeScanResult, expectedPageIndex As Integer) As SelectionOutcome
        If scan Is Nothing Then Return SelectionOutcome.NotRun("Hasil pemindaian tidak tersedia.")

        Dim targets As List(Of Object) = scan.SelectionTargets()
        If targets.Count = 0 Then
            ' Penting: jangan ClearSelection saat kosong (tidak ada operasi destruktif).
            Return SelectionOutcome.NotRun("Tidak ada target seleksi pada halaman " & expectedPageIndex.ToString() & ".")
        End If

        ' Tidak ada fallback Delete: kegagalan ClearSelection menghentikan operasi demi keselamatan.
        Try
            corelDoc.ClearSelection()
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
            Dim selectionRange As Object = corelApp.ActiveSelectionRange
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
            Using dlg As New WatermarkDialog(corelApp, selectedPages)
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
