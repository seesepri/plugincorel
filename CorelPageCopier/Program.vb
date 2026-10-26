Option Explicit On
Option Strict Off
Option Infer On
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing
Imports System.Runtime.InteropServices
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
    Private matchingPageIndices As New List(Of Integer)()
    Private currentMatchPagePosition As Integer = -1
    Private Const CDR_CURVE As Integer = 3
    Private Const CDR_BITMAP As Integer = 5
    Private Const CDR_TEXT As Integer = 6
    Private Const CDR_GROUP As Integer = 7
    Private Const CDR_UNIT_MM As Integer = 3
    Private Const GEOMETRY_TOLERANCE_MM As Double = 0.5

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

    Private Sub ConnectToCorel()
        Dim lastError As Exception = Nothing

        ' Gunakan koneksi yang sudah berhasil sebelumnya jika masih valid.
        If corelApp IsNot Nothing Then
            Try
                corelDoc = corelApp.ActiveDocument
                If corelDoc IsNot Nothing Then Return
            Catch
                corelApp = Nothing
                corelDoc = Nothing
            End Try
        End If

        ' Coba ambil instance CorelDRAW yang sudah berjalan.
        For Each progId As String In New String() {"CorelDRAW.Application.24", "CorelDRAW.Application"}
            Try
                corelApp = Marshal.GetActiveObject(progId)
                If corelApp IsNot Nothing Then
                    corelDoc = corelApp.ActiveDocument
                    If corelDoc IsNot Nothing Then Return
                    lastError = New InvalidOperationException("CorelDRAW ditemukan, tetapi belum ada dokumen aktif.")
                End If
            Catch ex As Exception
                lastError = ex
                corelApp = Nothing
                corelDoc = Nothing
            End Try
        Next

        ' Jika instance aktif tidak terdaftar di ROT, buat/ambil instance baru melalui COM.
        ' Ini mungkin membuka instance CorelDRAW kedua. Pengguna harus membuka CDR di instance tersebut.
        For Each progId As String In New String() {"CorelDRAW.Application.24", "CorelDRAW.Application"}
            Try
                Dim appType As Type = Type.GetTypeFromProgID(progId, False)
                If appType Is Nothing Then Continue For
                corelApp = Activator.CreateInstance(appType)
                If corelApp IsNot Nothing Then
                    Try
                        corelApp.Visible = True
                    Catch
                    End Try
                    Try
                        corelDoc = corelApp.ActiveDocument
                    Catch
                        corelDoc = Nothing
                    End Try
                    If corelDoc IsNot Nothing Then Return
                    Throw New InvalidOperationException("CorelDRAW berhasil dibuka melalui COM, tetapi belum memiliki dokumen aktif. Buka file CDR di jendela CorelDRAW yang baru terbuka, lalu klik Muat Ulang.")
                End If
            Catch ex As Exception
                lastError = ex
                corelApp = Nothing
                corelDoc = Nothing
            End Try
        Next

        Dim detail As String = If(lastError Is Nothing, "Tidak ada objek aplikasi yang tersedia.", lastError.Message)
        Throw New InvalidOperationException("Tidak dapat mengakses dokumen CorelDRAW. Pastikan CorelDRAW 2022 terpasang. Jika aplikasi baru terbuka melalui COM, buka dokumen CDR di jendela tersebut lalu klik Muat Ulang. Detail: " & detail, lastError)
    End Sub

    Private Sub LoadPages()
        ConnectToCorel()
        pagesList.Items.Clear()
        sourcePageIndex = CInt(corelDoc.ActivePage.Index)
        sourceLabel.Text = "Halaman aktif: " & sourcePageIndex & " | Pilih objek sumber sebelum menjalankan Copy."
        For i As Integer = 1 To CInt(corelDoc.Pages.Count)
            Dim caption = "Halaman " & i
            If i = sourcePageIndex Then caption &= " (AKTIF)"
            Dim itemIndex = pagesList.Items.Add(New PageItem(i, caption))
            If i = sourcePageIndex Then pagesList.SetItemChecked(itemIndex, False)
        Next
        statusLabel.Text = "Dokumen terdeteksi."
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

        ' Status dokumen yang wajib dipulihkan (Temuan #1 & #3).
        Dim inCommandGroup As Boolean = False
        Dim previousUnit As Integer = -1
        Dim unitStored As Boolean = False
        Dim sourceMeasured As Boolean = False
        Dim srcCenterX As Double = 0
        Dim srcCenterY As Double = 0
        Dim srcWidth As Double = 0
        Dim srcHeight As Double = 0

        Try
            ConnectToCorel()
            Dim actualSourcePage As Integer = CInt(corelDoc.ActivePage.Index)
            If actualSourcePage <> sourcePageIndex Then
                MessageBox.Show("Halaman aktif berubah. Klik Muat Ulang dan ulangi.", "Halaman sumber berubah", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            Dim documentPageCount As Integer = CInt(corelDoc.Pages.Count)
            Dim validatedTargets As New List(Of Integer)()
            Dim validationFailures As New List(Of String)()
            For Each pageIndex In targets
                If pageIndex < 1 OrElse pageIndex > documentPageCount Then
                    validationFailures.Add("Halaman " & pageIndex & ": indeks halaman tidak valid.")
                ElseIf pageIndex = actualSourcePage Then
                    validationFailures.Add("Halaman " & pageIndex & ": halaman sumber tidak boleh menjadi tujuan.")
                ElseIf Not validatedTargets.Contains(pageIndex) Then
                    validatedTargets.Add(pageIndex)
                End If
            Next
            If validatedTargets.Count = 0 Then
                MessageBox.Show("Tidak ada halaman tujuan yang valid." & Microsoft.VisualBasic.Constants.vbCrLf & String.Join(Microsoft.VisualBasic.Constants.vbCrLf, validationFailures.ToArray()), "Validasi tujuan", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim selection As Object = corelApp.ActiveSelectionRange
            If selection Is Nothing OrElse CInt(selection.Count) = 0 Then
                MessageBox.Show("Pilih objek pada halaman sumber terlebih dahulu.", "Tidak ada seleksi", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' --- Temuan #3: normalisasi satuan dokumen -------------------------
            ' CenterX/CenterY/SizeWidth/SizeHeight dibaca & ditulis dalam Document.Unit,
            ' yang bisa inch, pixel, atau point tergantung dokumen. Satuan dikunci ke mm
            ' selama operasi agar perbandingan geometri sumber vs hasil paste konsisten,
            ' lalu nilai aslinya dipulihkan di Finally.
            Try
                previousUnit = CInt(corelDoc.Unit)
                unitStored = True
            Catch
            End Try
            Try
                corelDoc.Unit = CDR_UNIT_MM
            Catch
            End Try

            ' Ukur objek sumber (gabungan ShapeRange) sebagai acuan verifikasi.
            Try
                srcCenterX = CDbl(selection.CenterX)
                srcCenterY = CDbl(selection.CenterY)
                srcWidth = CDbl(selection.SizeWidth)
                srcHeight = CDbl(selection.SizeHeight)
                sourceMeasured = True
            Catch
                sourceMeasured = False
            End Try

            ' Copy berjalan langsung tanpa dialog konfirmasi. Objek sumber tidak pernah dihapus.
            corelDoc.Pages(actualSourcePage).Activate()
            selection.Copy()

            ' --- Temuan #1: seluruh paste dalam satu grup undo ------------------
            ' Satu Ctrl+Z membatalkan semua paste ke semua halaman tujuan, bukan
            ' satu per satu. EndCommandGroup selalu dipanggil di Finally.
            Try
                corelDoc.BeginCommandGroup("Salin objek ke halaman")
                inCommandGroup = True
            Catch
            End Try

            Dim succeeded As Integer = 0
            Dim verifiedByCountOnly As Integer = 0
            Dim failures As New List(Of String)(validationFailures)

            For targetPosition As Integer = 0 To validatedTargets.Count - 1
                Dim pageIndex As Integer = validatedTargets(targetPosition)
                statusLabel.Text = "Copy progres: " & (targetPosition + 1) & "/" & validatedTargets.Count & " — halaman tujuan " & pageIndex & "."
                Application.DoEvents()   ' aman: flag busy + kontrol nonaktif mencegah re-entrancy
                Try
                    Dim destination As Object = corelDoc.Pages(pageIndex)
                    destination.Activate()
                    Dim layer As Object = destination.ActiveLayer
                    Dim beforeCount As Integer = CInt(layer.Shapes.Count)
                    layer.Paste()
                    Dim added As Integer = CInt(layer.Shapes.Count) - beforeCount

                    If added <= 0 Then
                        failures.Add("Halaman " & pageIndex & ": Paste tidak menambah shape pada layer aktif; periksa hasil secara manual.")
                        Continue For
                    End If

                    ' Verifikasi geometri (bukan hanya jumlah shape).
                    If Not sourceMeasured Then
                        succeeded += 1
                        verifiedByCountOnly += 1
                        Continue For
                    End If

                    Dim pastedCenterX As Double = 0
                    Dim pastedCenterY As Double = 0
                    Dim pastedWidth As Double = 0
                    Dim pastedHeight As Double = 0
                    If Not TryMeasureTopShapes(layer, added, pastedCenterX, pastedCenterY, pastedWidth, pastedHeight) Then
                        succeeded += 1
                        verifiedByCountOnly += 1
                        Continue For
                    End If

                    Dim delta As WatermarkLayout.GeometryDelta = WatermarkLayout.CompareGeometry(
                        srcCenterX, srcCenterY, srcWidth, srcHeight,
                        pastedCenterX, pastedCenterY, pastedWidth, pastedHeight)

                    If delta.Matches(GEOMETRY_TOLERANCE_MM) Then
                        succeeded += 1
                    Else
                        failures.Add("Halaman " & pageIndex & ": objek tertempel tetapi posisi/ukuran berbeda dari sumber (" &
                                     delta.Describe() & ") — periksa hasil secara manual.")
                    End If
                Catch pageEx As Exception
                    failures.Add("Halaman " & pageIndex & ": " & pageEx.Message)
                End Try
            Next

            Try
                corelDoc.Pages(actualSourcePage).Activate()
            Catch
                ' Jangan menutupi laporan utama bila pemulihan halaman gagal.
            End Try

            Dim report As String = "Copy selesai." & Microsoft.VisualBasic.Constants.vbCrLf &
                "Halaman tujuan valid diproses: " & validatedTargets.Count & "." & Microsoft.VisualBasic.Constants.vbCrLf &
                "Berhasil diverifikasi (geometri cocok): " & (succeeded - verifiedByCountOnly) & " halaman." & Microsoft.VisualBasic.Constants.vbCrLf
            If verifiedByCountOnly > 0 Then
                report &= "Terverifikasi lewat jumlah shape saja: " & verifiedByCountOnly & " halaman (geometri tidak dapat diukur)." & Microsoft.VisualBasic.Constants.vbCrLf
            End If
            report &= "Gagal/perlu diperiksa: " & failures.Count & " halaman."
            If failures.Count > 0 Then report &= Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf & String.Join(Microsoft.VisualBasic.Constants.vbCrLf, failures.ToArray())
            report &= Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf &
                "Catatan: posisi & ukuran objek hasil paste dibandingkan dengan objek sumber dalam mm (toleransi " &
                GEOMETRY_TOLERANCE_MM & " mm). Satuan dokumen dipulihkan otomatis. Semua paste berada dalam satu grup undo — satu Ctrl+Z membatalkan seluruh operasi."

            report &= Microsoft.VisualBasic.Constants.vbCrLf & "Durasi: " & stopwatch.Elapsed.TotalSeconds.ToString("0.0") & " detik."
            statusLabel.Text = "Copy selesai: " & succeeded & "/" & validatedTargets.Count & " halaman terverifikasi; " & failures.Count & " gagal/perlu diperiksa (" & stopwatch.Elapsed.TotalSeconds.ToString("0.0") & " detik)."
            MessageBox.Show(report, "Hasil Corel Page Copier", MessageBoxButtons.OK, If(failures.Count = 0, MessageBoxIcon.Information, MessageBoxIcon.Warning))
        Catch ex As Exception
            ShowError("Proses copy gagal: " & ex.Message)
        Finally
            If inCommandGroup Then
                Try
                    corelDoc.EndCommandGroup()
                Catch
                End Try
            End If
            If unitStored Then
                Try
                    corelDoc.Unit = previousUnit
                Catch
                End Try
            End If
        End Try
    End Sub

    Private Sub AutoSelectClicked(sender As Object, e As EventArgs)
        If Not TryBeginOperation("Auto Select") Then Return
        Try
            AutoSelectClickedCore()
        Finally
            EndOperation()
        End Try
    End Sub

    Private Sub AutoSelectClickedCore()
        Dim stopwatch As Stopwatch = Stopwatch.StartNew()
        Try
            ConnectToCorel()
            Dim category As Integer = categoryBox.SelectedIndex
            Dim allPages As Boolean = (scopeBox.SelectedIndex = 1)
            Dim pageCount As Integer = CInt(corelDoc.Pages.Count)
            Dim originalPage As Integer = CInt(corelDoc.ActivePage.Index)
            matchingPageIndices.Clear()
            currentMatchPagePosition = -1
            If Not allPages Then
                Dim matches As List(Of Object) = GetMatchingShapes(corelDoc.ActivePage, category)
                Dim selectedCount As Integer = ApplySelection(matches)
                statusLabel.Text = "Auto Select: " & selectedCount & " objek pada halaman " & originalPage & "."
                MessageBox.Show("Terpilih " & selectedCount & " objek pada halaman " & originalPage & ".", "Auto Select", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim counts As New Dictionary(Of Integer, Integer)()
            Dim total As Integer = 0
            For pageIndex As Integer = 1 To pageCount
                statusLabel.Text = "Memindai progres: halaman " & pageIndex & "/" & pageCount & " | total objek cocok sementara: " & total & " | halaman aktif saat ini: " & CInt(corelDoc.ActivePage.Index) & "."
                Application.DoEvents()   ' aman: flag busy + kontrol nonaktif mencegah re-entrancy
                Dim matches As List(Of Object) = GetMatchingShapes(corelDoc.Pages(pageIndex), category)
                counts(pageIndex) = matches.Count
                total += matches.Count
                If matches.Count > 0 Then matchingPageIndices.Add(pageIndex)
            Next
            If matchingPageIndices.Count = 0 Then
                MessageBox.Show("Tidak ditemukan objek yang cocok pada dokumen.", "Auto Select", MessageBoxButtons.OK, MessageBoxIcon.Information)
                statusLabel.Text = "Auto Select: tidak ada objek cocok."
                Return
            End If
            currentMatchPagePosition = 0
            SelectMatchingPageAtCurrentPosition(category)
            Dim details As New List(Of String)()
            For Each kvp In counts
                details.Add("Halaman " & kvp.Key & ": " & kvp.Value & " objek")
            Next
            Dim report As String = "Pemindaian semua halaman selesai." & Microsoft.VisualBasic.Constants.vbCrLf & "Total objek cocok: " & total & Microsoft.VisualBasic.Constants.vbCrLf & "Durasi pemindaian: " & stopwatch.Elapsed.TotalSeconds.ToString("0.0") & " detik." & Microsoft.VisualBasic.Constants.vbCrLf & "Halaman yang memiliki kecocokan: " & matchingPageIndices.Count & Microsoft.VisualBasic.Constants.vbCrLf & "Seleksi saat ini: halaman " & matchingPageIndices(0) & ". Gunakan tombol 'Pilih Halaman Cocok Berikutnya' untuk berpindah." & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf & String.Join(Microsoft.VisualBasic.Constants.vbCrLf, details.ToArray())
            statusLabel.Text = "Pindai semua halaman: " & total & " objek; " & matchingPageIndices.Count & " halaman cocok."
            MessageBox.Show(report, "Auto Select — Semua Halaman", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            ShowError("Auto Select gagal: " & ex.Message)
        End Try
    End Sub

    Private Sub NextMatchingPageClicked(sender As Object, e As EventArgs)
        If Not TryBeginOperation("Pindah hasil") Then Return
        Try
            NextMatchingPageClickedCore()
        Finally
            EndOperation()
        End Try
    End Sub

    Private Sub NextMatchingPageClickedCore()
        Try
            ConnectToCorel()
            If matchingPageIndices.Count = 0 Then
                MessageBox.Show("Jalankan 'Select Otomatis' dengan cakupan 'Pindai semua halaman' terlebih dahulu.", "Tidak ada hasil pemindaian", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If
            If currentMatchPagePosition >= matchingPageIndices.Count - 1 Then
                statusLabel.Text = "Navigasi selesai: hasil terakhir " & matchingPageIndices.Count & "/" & matchingPageIndices.Count & ", halaman aktif " & CInt(corelDoc.ActivePage.Index) & "."
                MessageBox.Show("Anda sudah berada pada hasil terakhir (" & matchingPageIndices.Count & "/" & matchingPageIndices.Count & "). Navigasi dihentikan dan tidak kembali ke hasil pertama.", "Hasil terakhir", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If
            currentMatchPagePosition += 1
            SelectMatchingPageAtCurrentPosition(categoryBox.SelectedIndex)
        Catch ex As Exception
            ShowError("Gagal berpindah ke halaman cocok: " & ex.Message)
        End Try
    End Sub

    Private Sub SelectMatchingPageAtCurrentPosition(category As Integer)
        If currentMatchPagePosition < 0 OrElse currentMatchPagePosition >= matchingPageIndices.Count Then Return
        Dim pageIndex As Integer = matchingPageIndices(currentMatchPagePosition)
        Dim page As Object = corelDoc.Pages(pageIndex)
        page.Activate()
        Dim matches As List(Of Object) = GetMatchingShapes(page, category)
        Dim selectedCount As Integer = ApplySelection(matches)
        statusLabel.Text = "Hasil " & (currentMatchPagePosition + 1) & "/" & matchingPageIndices.Count & " | halaman cocok " & pageIndex & " | " & selectedCount & " objek terpilih | halaman aktif " & CInt(corelDoc.ActivePage.Index) & "."
    End Sub

    Private Function GetMatchingShapes(page As Object, category As Integer) As List(Of Object)
        Dim matches As New List(Of Object)()
        Dim shapes As Object = page.Shapes
        For i As Integer = 1 To CInt(shapes.Count)
            CollectMatchingShape(shapes(i), category, matches, 0)
        Next
        Return matches
    End Function

    Private Sub CollectMatchingShape(shape As Object, category As Integer, matches As List(Of Object), depth As Integer)
        ' Batas kedalaman mencegah rekursi tidak berujung jika COM mengembalikan struktur tak terduga.
        If shape Is Nothing OrElse depth > 32 Then Return

        ' Temuan #2: objek watermark (nama WM_* / layer WATERMARK) tidak boleh ikut
        ' terpilih. Tanpa ini, "Salin ke halaman" akan menyalin watermark ke halaman
        ' tujuan sehingga watermark menumpuk.
        If IsWatermarkShape(shape, depth) Then Return
        Dim shapeType As Integer
        Try
            shapeType = CInt(shape.Type)
        Catch
            Return
        End Try

        ' Untuk kategori Semua Objek, pilih grup sebagai satu kesatuan agar struktur grup terjaga.
        If category = 4 Then
            matches.Add(shape)
            Return
        End If

        If shapeType = CDR_GROUP Then
            Try
                Dim children As Object = shape.Shapes
                For i As Integer = 1 To CInt(children.Count)
                    CollectMatchingShape(children(i), category, matches, depth + 1)
                Next
            Catch
                ' Grup yang isinya tidak dapat dibaca dilewati secara aman.
            End Try
            Return
        End If

        If MatchesCategory(shape, category) Then matches.Add(shape)
    End Sub

    Private Function MatchesCategory(shape As Object, category As Integer) As Boolean
        Dim shapeType As Integer
        Try
            shapeType = CInt(shape.Type)
        Catch
            Return False
        End Try
        Select Case category
            Case 0
                Return shapeType = CDR_TEXT
            Case 1
                Return shapeType = CDR_BITMAP
            Case 2
                If shapeType <> CDR_CURVE Then Return False
                Try
                    ' Jika status Closed tidak dapat dibaca, jangan menganggap objek sebagai kurva terbuka.
                    Return CBool(shape.Curve.Closed) = False
                Catch
                    Return False
                End Try
            Case 3
                Return shapeType <> CDR_TEXT AndAlso shapeType <> CDR_BITMAP AndAlso shapeType <> CDR_GROUP
            Case Else
                Return True
        End Select
    End Function

    Private Function ApplySelection(matches As List(Of Object)) As Integer
        ' Tidak ada fallback Delete: kegagalan ClearSelection menghentikan operasi demi keselamatan.
        Try
            corelDoc.ClearSelection()
        Catch ex As Exception
            Throw New InvalidOperationException("Seleksi lama tidak dapat dibersihkan. Tidak ada objek yang dihapus; proses dibatalkan.", ex)
        End Try

        Dim selectedCount As Integer = 0
        Dim failures As Integer = 0
        For Each shape As Object In matches
            Try
                shape.AddToSelection()
                selectedCount += 1
            Catch
                failures += 1
            End Try
        Next
        If failures > 0 Then
            statusLabel.Text = "Seleksi parsial: " & selectedCount & " berhasil, " & failures & " gagal."
        End If
        Return selectedCount
    End Function

    ''' <summary>
    ''' Deteksi objek watermark: nama shape berawalan WM_, berada di layer WATERMARK,
    ''' atau grup yang berisi salah satunya. Dipakai Auto Select agar watermark tidak
    ''' pernah ikut terpilih (Temuan #2).
    ''' </summary>
    Private Function IsWatermarkShape(shape As Object, Optional depth As Integer = 0) As Boolean
        If shape Is Nothing OrElse depth > 32 Then Return False

        Dim shapeName As String = Nothing
        Dim layerName As String = Nothing
        Try
            shapeName = CStr(shape.Name)
        Catch
        End Try
        Try
            layerName = CStr(shape.Layer.Name)
        Catch
        End Try
        If WatermarkLayout.IsWatermarkIdentity(shapeName, layerName) Then Return True

        Try
            If CInt(shape.Type) = CDR_GROUP Then
                Dim children As Object = shape.Shapes
                For i As Integer = 1 To CInt(children.Count)
                    If IsWatermarkShape(children(i), depth + 1) Then Return True
                Next
            End If
        Catch
        End Try

        Return False
    End Function

    ''' <summary>
    ''' Ukur kotak pembatas gabungan dari N shape teratas pada sebuah layer (hasil Paste
    ''' selalu berada paling atas). Mengembalikan False bila pengukuran tidak mungkin.
    ''' Semua nilai dalam satuan dokumen saat ini (mm selama operasi copy).
    ''' </summary>
    Private Function TryMeasureTopShapes(layer As Object, count As Integer,
                                         ByRef centerX As Double, ByRef centerY As Double,
                                         ByRef width As Double, ByRef height As Double) As Boolean
        If count <= 0 Then Return False
        Try
            Dim shapes As Object = layer.Shapes
            If CInt(shapes.Count) < count Then Return False

            Dim minX As Double = Double.MaxValue
            Dim maxX As Double = Double.MinValue
            Dim minY As Double = Double.MaxValue
            Dim maxY As Double = Double.MinValue
            Dim measured As Integer = 0

            For i As Integer = 1 To count
                Dim sh As Object = shapes(i)
                Dim cx As Double = CDbl(sh.CenterX)
                Dim cy As Double = CDbl(sh.CenterY)
                Dim w As Double = CDbl(sh.SizeWidth)
                Dim h As Double = CDbl(sh.SizeHeight)
                minX = Math.Min(minX, cx - w / 2.0)
                maxX = Math.Max(maxX, cx + w / 2.0)
                minY = Math.Min(minY, cy - h / 2.0)
                maxY = Math.Max(maxY, cy + h / 2.0)
                measured += 1
            Next

            If measured = 0 Then Return False
            centerX = (minX + maxX) / 2.0
            centerY = (minY + maxY) / 2.0
            width = maxX - minX
            height = maxY - minY
            Return True
        Catch
            Return False
        End Try
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
