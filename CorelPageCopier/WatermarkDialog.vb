Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  WatermarkDialog.vb
'  UI fitur watermark (WinForms, gaya sama dengan CopierForm: kartu putih,
'  tombol datar, biru sebagai aksi utama).
'
'  Dipakai dari CopierForm:
'      Dim dlg As New WatermarkDialog(corelApp)
'      dlg.ShowDialog(Me)
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

Public Class WatermarkDialog
    Inherits Form

    Private ReadOnly corelApp As Object

    ' --- Kontrol -------------------------------------------------------------
    Private ReadOnly kindBox As New ComboBox()
    Private ReadOnly textBox As New ComboBox()
    Private ReadOnly fontBox As New ComboBox()
    Private ReadOnly fontSizeBox As New NumericUpDown()
    Private ReadOnly colorPanel As New Panel()
    Private ReadOnly logoPathBox As New TextBox()
    Private ReadOnly logoWidthBox As New NumericUpDown()
    Private ReadOnly opacityBox As New NumericUpDown()
    Private ReadOnly rotationBox As New NumericUpDown()
    Private ReadOnly anchorBox As New ComboBox()
    Private ReadOnly marginBox As New NumericUpDown()
    Private ReadOnly tileCheck As New CheckBox()
    Private ReadOnly tileXBox As New NumericUpDown()
    Private ReadOnly tileYBox As New NumericUpDown()
    Private ReadOnly replaceCheck As New CheckBox()
    Private ReadOnly lockLayerCheck As New CheckBox()
    Private ReadOnly perPageCheck As New CheckBox()
    Private ReadOnly layerNameBox As New TextBox()
    Private ReadOnly scopeBox As New ComboBox()
    Private ReadOnly folderBox As New TextBox()
    Private ReadOnly subfolderCheck As New CheckBox()
    Private ReadOnly saveInPlaceCheck As New CheckBox()
    Private ReadOnly reportBox As New TextBox()
    Private ReadOnly statusLabel As New Label()

    Private ReadOnly bundledLogoCheck As New CheckBox()
    Private ReadOnly logoVariantBox As New ComboBox()
    Private ReadOnly preselectedPages As New List(Of Integer)()

    Private selectedColor As Color = Color.FromArgb(120, 120, 120)
    Private busy As Boolean = False

    Public Sub New(app As Object, Optional preselected As List(Of Integer) = Nothing)
        corelApp = app
        If preselected IsNot Nothing Then
            For Each p As Integer In preselected
                If Not preselectedPages.Contains(p) Then preselectedPages.Add(p)
            Next
            preselectedPages.Sort()
        End If

        Text = "Watermark — Corel Page Copier"
        StartPosition = FormStartPosition.CenterParent
        Size = New Size(760, 780)
        MinimumSize = New Size(680, 640)
        BackColor = Color.FromArgb(245, 247, 250)
        Font = New Font("Segoe UI", 9.0F)

        BuildUi()
        LoadDefaults()
    End Sub

    ' ========================================================================
    '  UI
    ' ========================================================================

    Private Sub BuildUi()
        SuspendLayout()

        Dim root As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 4,
            .Padding = New Padding(16),
            .BackColor = BackColor
        }
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 62.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 150.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 92.0F))
        Controls.Add(root)

        ' --- Header ----------------------------------------------------------
        Dim header As New Panel With {.Dock = DockStyle.Fill, .BackColor = BackColor}
        Dim title As New Label With {
            .Text = "Watermark",
            .Font = New Font("Segoe UI", 17.0F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(25, 35, 52),
            .AutoSize = True,
            .Location = New Point(0, 0)
        }
        Dim subtitle As New Label With {
            .Text = "Teks dan/atau logo, satu halaman, semua halaman, atau batch folder",
            .Font = New Font("Segoe UI", 8.5F),
            .ForeColor = Color.FromArgb(102, 115, 132),
            .AutoSize = True,
            .Location = New Point(2, 32)
        }
        header.Controls.Add(title)
        header.Controls.Add(subtitle)
        root.Controls.Add(header, 0, 0)

        ' --- Kartu isi (scroll) ----------------------------------------------
        Dim card As New Panel With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle,
            .Padding = New Padding(0, 0, 0, 0)
        }
        Dim scroll As New Panel With {.Dock = DockStyle.Fill, .AutoScroll = True, .BackColor = Color.White}
        Dim grid As New TableLayoutPanel With {
            .Dock = DockStyle.Top,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .ColumnCount = 2,
            .Padding = New Padding(14),
            .BackColor = Color.White
        }
        grid.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 160.0F))
        grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

        ' Jenis
        kindBox.DropDownStyle = ComboBoxStyle.DropDownList
        kindBox.Items.AddRange(New Object() {"Teks saja", "Logo saja", "Teks + logo"})
        kindBox.SelectedIndex = 0
        AddRow(grid, "Jenis watermark", kindBox)

        ' Teks (editable ComboBox supaya ada saran cepat)
        textBox.DropDownStyle = ComboBoxStyle.DropDown
        textBox.Items.AddRange(New Object() {"CONTOH", "DRAFT", "RAHASIA", "SALINAN", "PROOF ONLY"})
        AddRow(grid, "Teks watermark", textBox)

        ' Font + ukuran
        Dim fontRow As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .AutoSize = True, .WrapContents = False, .Margin = New Padding(0)}
        fontBox.DropDownStyle = ComboBoxStyle.DropDown
        fontBox.Width = 220
        fontBox.Items.AddRange(New Object() {"Segoe UI", "Arial", "Calibri", "Times New Roman", "Verdana", "Tahoma", "Georgia", "Impact"})
        fontSizeBox.Minimum = 4
        fontSizeBox.Maximum = 500
        fontSizeBox.DecimalPlaces = 0
        fontSizeBox.Width = 80
        Dim ptLabel As New Label With {.Text = "pt", .AutoSize = True, .Margin = New Padding(6, 7, 0, 0), .ForeColor = Color.FromArgb(102, 115, 132)}
        fontRow.Controls.Add(fontBox)
        fontRow.Controls.Add(fontSizeBox)
        fontRow.Controls.Add(ptLabel)
        AddRow(grid, "Font & ukuran", fontRow)

        ' Warna
        Dim colorRow As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .AutoSize = True, .WrapContents = False, .Margin = New Padding(0)}
        colorPanel.Width = 60
        colorPanel.Height = 24
        colorPanel.BackColor = selectedColor
        colorPanel.BorderStyle = BorderStyle.FixedSingle
        colorPanel.Cursor = Cursors.Hand
        AddHandler colorPanel.Click, AddressOf PickColor
        Dim colorHint As New Label With {
            .Text = "klik untuk pilih warna",
            .AutoSize = True,
            .Margin = New Padding(8, 6, 0, 0),
            .ForeColor = Color.FromArgb(102, 115, 132)
        }
        colorRow.Controls.Add(colorPanel)
        colorRow.Controls.Add(colorHint)
        AddRow(grid, "Warna teks", colorRow)

        ' Logo
        Dim logoRow As New TableLayoutPanel With {.Dock = DockStyle.Fill, .AutoSize = True, .ColumnCount = 2, .Margin = New Padding(0)}
        logoRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        logoRow.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 92.0F))
        logoPathBox.Dock = DockStyle.Fill
        logoPathBox.Margin = New Padding(0, 3, 6, 3)
        Dim browseLogo As Button = FlatButton("Pilih…", False)
        browseLogo.Dock = DockStyle.Fill
        browseLogo.Margin = New Padding(0, 3, 0, 3)
        AddHandler browseLogo.Click, AddressOf PickLogo
        logoRow.Controls.Add(logoPathBox, 0, 0)
        logoRow.Controls.Add(browseLogo, 1, 0)
        AddRow(grid, "File logo (PNG/JPG)", logoRow)

        bundledLogoCheck.Text = "Pakai logo bawaan SETA (PNG transparan, sudah tertanam di aplikasi)"
        bundledLogoCheck.AutoSize = True
        bundledLogoCheck.Margin = New Padding(0, 6, 0, 0)
        bundledLogoCheck.Enabled = HasBundledLogo()
        If HasBundledLogo() Then bundledLogoCheck.Checked = True
        AddHandler bundledLogoCheck.CheckedChanged, AddressOf BundledLogoToggled
        AddRow(grid, "Logo bawaan", bundledLogoCheck)

        ' Varian warna logo bawaan: hitam (default) atau putih, tanpa perlu cari file.
        logoVariantBox.DropDownStyle = ComboBoxStyle.DropDownList
        logoVariantBox.Items.Add("Hitam — untuk desain berlatar terang")
        logoVariantBox.Items.Add("Putih — untuk desain berlatar gelap")
        logoVariantBox.SelectedIndex = 0
        logoVariantBox.Dock = DockStyle.Fill
        AddHandler logoVariantBox.SelectedIndexChanged, AddressOf LogoTyped
        AddRow(grid, "Warna logo", logoVariantBox)

        logoWidthBox.Minimum = 0
        logoWidthBox.Maximum = 5000
        logoWidthBox.DecimalPlaces = 1
        logoWidthBox.Width = 90
        AddRow(grid, "Lebar logo (mm, 0 = asli)", logoWidthBox)

        ' Opasitas + rotasi
        Dim styleRow As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .AutoSize = True, .WrapContents = False, .Margin = New Padding(0)}
        opacityBox.Minimum = 1
        opacityBox.Maximum = 100
        opacityBox.Width = 70
        Dim pctLabel As New Label With {.Text = "% opasitas", .AutoSize = True, .Margin = New Padding(6, 7, 18, 0), .ForeColor = Color.FromArgb(102, 115, 132)}
        rotationBox.Minimum = -180
        rotationBox.Maximum = 180
        rotationBox.Width = 70
        Dim degLabel As New Label With {.Text = "° rotasi", .AutoSize = True, .Margin = New Padding(6, 7, 0, 0), .ForeColor = Color.FromArgb(102, 115, 132)}
        styleRow.Controls.Add(opacityBox)
        styleRow.Controls.Add(pctLabel)
        styleRow.Controls.Add(rotationBox)
        styleRow.Controls.Add(degLabel)
        AddRow(grid, "Opasitas & rotasi", styleRow)

        ' Posisi
        anchorBox.DropDownStyle = ComboBoxStyle.DropDownList
        For Each anchorLabel As String In WatermarkSettings.AnchorLabels()
            anchorBox.Items.Add(anchorLabel)
        Next
        anchorBox.SelectedIndex = 4
        AddRow(grid, "Posisi", anchorBox)

        marginBox.Minimum = 0
        marginBox.Maximum = 200
        marginBox.DecimalPlaces = 1
        marginBox.Width = 90
        AddRow(grid, "Margin dari tepi (mm)", marginBox)

        ' Tile
        tileCheck.Text = "Ulangi watermark sebagai pola (tile) menutupi halaman"
        tileCheck.AutoSize = True
        tileCheck.Margin = New Padding(0, 4, 0, 0)
        AddRow(grid, "Mode tile", tileCheck)

        Dim tileRow As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .AutoSize = True, .WrapContents = False, .Margin = New Padding(0)}
        tileXBox.Minimum = 1
        tileXBox.Maximum = 2000
        tileXBox.DecimalPlaces = 1
        tileXBox.Width = 70
        tileYBox.Minimum = 1
        tileYBox.Maximum = 2000
        tileYBox.DecimalPlaces = 1
        tileYBox.Width = 70
        Dim tileXLabel As New Label With {.Text = "jarak X (mm)", .AutoSize = True, .Margin = New Padding(6, 7, 18, 0), .ForeColor = Color.FromArgb(102, 115, 132)}
        Dim tileYLabel As New Label With {.Text = "jarak Y (mm)", .AutoSize = True, .Margin = New Padding(6, 7, 0, 0), .ForeColor = Color.FromArgb(102, 115, 132)}
        tileRow.Controls.Add(tileXBox)
        tileRow.Controls.Add(tileXLabel)
        tileRow.Controls.Add(tileYBox)
        tileRow.Controls.Add(tileYLabel)
        AddRow(grid, "Jarak tile", tileRow)

        ' Cakupan halaman
        scopeBox.DropDownStyle = ComboBoxStyle.DropDownList
        scopeBox.Items.Add("Semua halaman")
        scopeBox.Items.Add("Halaman aktif saja")
        If preselectedPages.Count > 0 Then
            scopeBox.Items.Add("Halaman yang dicentang (" & preselectedPages.Count & ")")
            scopeBox.SelectedIndex = 2
        Else
            scopeBox.SelectedIndex = 0
        End If
        AddRow(grid, "Cakupan halaman", scopeBox)

        ' Layer & opsi
        layerNameBox.Text = "WATERMARK"
        AddRow(grid, "Nama layer", layerNameBox)

        replaceCheck.Text = "Ganti watermark lama (aman dijalankan berulang)"
        replaceCheck.AutoSize = True
        replaceCheck.Checked = True
        replaceCheck.Margin = New Padding(0, 4, 0, 0)
        AddRow(grid, "Idempoten", replaceCheck)

        perPageCheck.Text = "Tambahkan nomor halaman pada teks"
        perPageCheck.AutoSize = True
        perPageCheck.Margin = New Padding(0, 4, 0, 0)
        AddRow(grid, "Nomor halaman", perPageCheck)

        lockLayerCheck.Text = "Kunci layer setelah diterapkan (tidak bisa diklik tak sengaja)"
        lockLayerCheck.AutoSize = True
        lockLayerCheck.Margin = New Padding(0, 4, 0, 0)
        AddRow(grid, "Kunci", lockLayerCheck)

        ' --- Batch -----------------------------------------------------------
        Dim batchHeader As New Label With {
            .Text = "BATCH — terapkan ke semua file .cdr dalam folder",
            .Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(35, 48, 67),
            .AutoSize = True,
            .Margin = New Padding(0, 14, 0, 6)
        }
        Dim batchHeaderRow As Integer = grid.RowCount
        grid.RowCount = batchHeaderRow + 1
        grid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        grid.Controls.Add(batchHeader, 0, batchHeaderRow)
        grid.SetColumnSpan(batchHeader, 2)

        Dim folderRow As New TableLayoutPanel With {.Dock = DockStyle.Fill, .AutoSize = True, .ColumnCount = 2, .Margin = New Padding(0)}
        folderRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        folderRow.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 92.0F))
        folderBox.Dock = DockStyle.Fill
        folderBox.Margin = New Padding(0, 3, 6, 3)
        Dim browseFolder As Button = FlatButton("Folder…", False)
        browseFolder.Dock = DockStyle.Fill
        browseFolder.Margin = New Padding(0, 3, 0, 3)
        AddHandler browseFolder.Click, AddressOf PickFolder
        folderRow.Controls.Add(folderBox, 0, 0)
        folderRow.Controls.Add(browseFolder, 1, 0)
        AddRow(grid, "Folder .cdr", folderRow)

        subfolderCheck.Text = "Sertakan subfolder"
        subfolderCheck.AutoSize = True
        subfolderCheck.Margin = New Padding(0, 4, 0, 0)
        AddRow(grid, "Subfolder", subfolderCheck)

        saveInPlaceCheck.Text = "Timpa file asli (jika tidak dicentang, hasil = nama_wm.cdr)"
        saveInPlaceCheck.AutoSize = True
        saveInPlaceCheck.Margin = New Padding(0, 4, 0, 0)
        AddRow(grid, "Simpan", saveInPlaceCheck)

        Dim batchRow As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .AutoSize = True, .WrapContents = False, .Margin = New Padding(0, 6, 0, 0)}
        Dim batchButton As Button = FlatButton("Jalankan batch", False)
        batchButton.Width = 160
        batchButton.Height = 30
        AddHandler batchButton.Click, AddressOf BatchClicked
        batchRow.Controls.Add(batchButton)
        AddRow(grid, "", batchRow)

        scroll.Controls.Add(grid)
        card.Controls.Add(scroll)
        root.Controls.Add(card, 0, 1)

        ' --- Laporan ---------------------------------------------------------
        Dim reportCard As New Panel With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle,
            .Padding = New Padding(10),
            .Margin = New Padding(0, 10, 0, 10)
        }
        reportBox.Dock = DockStyle.Fill
        reportBox.Multiline = True
        reportBox.ReadOnly = True
        reportBox.ScrollBars = ScrollBars.Vertical
        reportBox.BorderStyle = BorderStyle.None
        reportBox.BackColor = Color.White
        reportBox.Font = New Font("Consolas", 8.5F)
        reportBox.Text = "Laporan akan muncul di sini."
        reportCard.Controls.Add(reportBox)
        root.Controls.Add(reportCard, 0, 2)

        ' --- Status + tombol -------------------------------------------------
        Dim footer As New Panel With {.Dock = DockStyle.Fill, .BackColor = BackColor}
        statusLabel.Dock = DockStyle.Top
        statusLabel.Height = 34
        statusLabel.TextAlign = ContentAlignment.MiddleLeft
        statusLabel.ForeColor = Color.FromArgb(35, 71, 105)
        statusLabel.Text = "Siap. Default: teks di tengah halaman, opasitas 50%."

        Dim buttonRow As New FlowLayoutPanel With {
            .Dock = DockStyle.Bottom,
            .Height = 44,
            .FlowDirection = FlowDirection.RightToLeft,
            .WrapContents = False
        }
        Dim closeButton As Button = FlatButton("Tutup", False)
        closeButton.Width = 100
        AddHandler closeButton.Click, Sub(s, e) Close()
        Dim removeButton As Button = FlatButton("Hapus watermark", False)
        removeButton.Width = 150
        AddHandler removeButton.Click, AddressOf RemoveClicked
        Dim applyButton As Button = FlatButton("Terapkan", True)
        applyButton.Width = 130
        AddHandler applyButton.Click, AddressOf ApplyClicked
        buttonRow.Controls.Add(closeButton)
        buttonRow.Controls.Add(removeButton)
        buttonRow.Controls.Add(applyButton)

        footer.Controls.Add(buttonRow)
        footer.Controls.Add(statusLabel)
        root.Controls.Add(footer, 0, 3)

        ResumeLayout(True)
    End Sub

    Private Function FlatButton(caption As String, primary As Boolean) As Button
        Dim b As New Button With {
            .Text = caption,
            .Height = 32,
            .Width = 120,
            .Margin = New Padding(8, 4, 0, 0),
            .FlatStyle = FlatStyle.Flat,
            .UseVisualStyleBackColor = False,
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

    Private Sub AddRow(grid As TableLayoutPanel, caption As String, control As Control)
        Dim label As New Label With {
            .Text = caption,
            .AutoSize = True,
            .ForeColor = Color.FromArgb(70, 82, 99),
            .Margin = New Padding(0, 7, 8, 7),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        Dim index As Integer = grid.RowCount
        grid.RowCount = index + 1
        grid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        grid.Controls.Add(label, 0, index)
        grid.Controls.Add(control, 1, index)
    End Sub

    Private Sub LoadDefaults()
        If bundledLogoCheck.Checked Then
            logoPathBox.Text = BundledLogoCaption()
            logoPathBox.ReadOnly = True
            logoVariantBox.Enabled = True
        Else
            logoVariantBox.Enabled = False
        End If
        textBox.Text = "CONTOH"
        fontBox.Text = "Segoe UI"
        fontSizeBox.Value = 24
        logoWidthBox.Value = 35
        opacityBox.Value = 50
        rotationBox.Value = 0
        marginBox.Value = 8
        tileXBox.Value = 60
        tileYBox.Value = 45
        anchorBox.SelectedIndex = 4
    End Sub

    ' ========================================================================
    '  PEMILIH FILE / WARNA
    ' ========================================================================

    ''' <summary>Nama file resource untuk varian terpilih (0 = hitam, 1 = putih).</summary>
    Private Function BundledLogoFileName() As String
        Return If(logoVariantBox.SelectedIndex = 1, "seta-logo-white.png", "seta-logo-black.png")
    End Function

    ''' <summary>Cari resource logo bawaan berdasarkan nama file (mis. seta-logo-black.png).</summary>
    Private Function FindBundledLogoResource(fileName As String) As String
        Try
            Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
            For Each candidate As String In asm.GetManifestResourceNames()
                If candidate.EndsWith("." & fileName, StringComparison.OrdinalIgnoreCase) OrElse
                   candidate.EndsWith(fileName, StringComparison.OrdinalIgnoreCase) Then
                    Return candidate
                End If
            Next
        Catch
        End Try
        Return ""
    End Function

    ''' <summary>True bila salah satu varian logo bawaan tertanam di assembly.</summary>
    Private Function HasBundledLogo() As Boolean
        Return FindBundledLogoResource("seta-logo-black.png").Length > 0 OrElse
               FindBundledLogoResource("seta-logo-white.png").Length > 0
    End Function

    ''' <summary>
    ''' Ekstrak logo bawaan ke folder temp (CorelDRAW mengimpor dari path file,
    ''' bukan dari stream) dan kembalikan path-nya. String kosong bila gagal.
    ''' </summary>
    Private Function EnsureBundledLogoFile() As String
        Dim fileName As String = BundledLogoFileName()
        Dim resourceName As String = FindBundledLogoResource(fileName)
        If resourceName.Length = 0 Then Return ""
        Try
            Dim folder As String = Path.Combine(Path.GetTempPath(), "CorelPageCopier")
            Directory.CreateDirectory(folder)
            Dim target As String = Path.Combine(folder, fileName)
            Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
            Using stream As IO.Stream = asm.GetManifestResourceStream(resourceName)
                If stream Is Nothing Then Return ""
                Using fileStream As IO.FileStream = IO.File.Create(target)
                    stream.CopyTo(fileStream)
                End Using
            End Using
            Return target
        Catch
            Return ""
        End Try
    End Function

    Private Sub BundledLogoToggled(sender As Object, e As EventArgs)
        Try
            logoPathBox.ReadOnly = bundledLogoCheck.Checked
            logoVariantBox.Enabled = bundledLogoCheck.Checked
            If bundledLogoCheck.Checked Then
                logoPathBox.Text = BundledLogoCaption()
            ElseIf IsBundledCaption(logoPathBox.Text) Then
                logoPathBox.Text = ""
            End If
        Catch
        End Try
    End Sub

    ''' <summary>Ganti varian logo bawaan -> perbarui keterangan di kolom path.</summary>
    Private Sub LogoTyped(sender As Object, e As EventArgs)
        Try
            If bundledLogoCheck.Checked Then logoPathBox.Text = BundledLogoCaption()
        Catch
        End Try
    End Sub

    Private Function BundledLogoCaption() As String
        Return If(logoVariantBox.SelectedIndex = 1, "(logo bawaan SETA — putih)", "(logo bawaan SETA — hitam)")
    End Function

    Private Function IsBundledCaption(value As String) As Boolean
        Return value IsNot Nothing AndAlso value.StartsWith("(logo bawaan SETA", StringComparison.Ordinal)
    End Function

    Private Sub PickColor(sender As Object, e As EventArgs)
        Using dlg As New ColorDialog()
            dlg.Color = selectedColor
            dlg.FullOpen = True
            If dlg.ShowDialog(Me) = DialogResult.OK Then
                selectedColor = dlg.Color
                colorPanel.BackColor = selectedColor
            End If
        End Using
    End Sub

    Private Sub PickLogo(sender As Object, e As EventArgs)
        Using dlg As New OpenFileDialog()
            dlg.Title = "Pilih logo watermark"
            dlg.Filter = "Gambar|*.png;*.jpg;*.jpeg;*.tif;*.tiff;*.bmp;*.gif;*.emf|Semua file|*.*"
            If dlg.ShowDialog(Me) = DialogResult.OK Then logoPathBox.Text = dlg.FileName
        End Using
    End Sub

    Private Sub PickFolder(sender As Object, e As EventArgs)
        Using dlg As New FolderBrowserDialog()
            dlg.Description = "Folder berisi file .cdr"
            If dlg.ShowDialog(Me) = DialogResult.OK Then folderBox.Text = dlg.SelectedPath
        End Using
    End Sub

    ' ========================================================================
    '  AKSI
    ' ========================================================================

    Private Sub ApplyClicked(sender As Object, e As EventArgs)
        If busy Then Return
        Dim settings As WatermarkSettings
        Try
            settings = ReadSettings()
        Catch ex As Exception
            MessageBox.Show("Pengaturan tidak valid: " & ex.Message, "Watermark", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End Try

        Dim errors As List(Of String) = settings.Validate()
        If errors.Count > 0 Then
            MessageBox.Show(String.Join(Environment.NewLine, errors.ToArray()), "Pengaturan belum lengkap", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim doc As Object = Nothing
        Try
            doc = GetActiveDocument()
            ' Cakupan "halaman aktif saja": isi target dari halaman yang sedang aktif.
            ' Cakupan "halaman yang dicentang" sudah terisi dari CopierForm.
            If Not settings.AllPages AndAlso settings.TargetPages.Count = 0 Then
                settings.TargetPages.Add(CInt(doc.ActivePage.Index))
            End If
        Catch ex As Exception
            MessageBox.Show("Tidak dapat mengakses dokumen CorelDRAW: " & ex.Message, "Watermark", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try

        SetBusy(True)
        Try
            Dim engine As New WatermarkEngine()
            AddHandler engine.Progress, Sub(message As String) UpdateStatus(message)
            Dim result As WatermarkResult = engine.ApplyToDocument(doc, settings)
            reportBox.Text = "Pengaturan: " & settings.Describe() & Environment.NewLine & Environment.NewLine & result.ToReport()
            reportBox.SelectionStart = 0
            reportBox.ScrollToCaret()
            UpdateStatus(If(result.Ok, "Watermark selesai.", "Watermark selesai dengan catatan — lihat laporan."))
        Catch ex As Exception
            MessageBox.Show("Watermark gagal: " & ex.Message, "Watermark", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            SetBusy(False)
        End Try
    End Sub

    Private Sub RemoveClicked(sender As Object, e As EventArgs)
        If busy Then Return
        If MessageBox.Show("Hapus semua watermark (shape bernama WM_*) dari dokumen aktif?",
                           "Hapus watermark", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
            Return
        End If

        SetBusy(True)
        Try
            Dim doc As Object = GetActiveDocument()
            Dim engine As New WatermarkEngine()
            Dim result As WatermarkResult = engine.RemoveFromDocument(doc)
            reportBox.Text = result.ToReport()
            UpdateStatus(If(result.PagesPassed > 0, "Watermark dihapus.", "Tidak ada watermark yang ditemukan."))
        Catch ex As Exception
            MessageBox.Show("Gagal menghapus watermark: " & ex.Message, "Watermark", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            SetBusy(False)
        End Try
    End Sub

    Private Sub BatchClicked(sender As Object, e As EventArgs)
        If busy Then Return

        Dim settings As WatermarkSettings
        Try
            settings = ReadSettings()
        Catch ex As Exception
            MessageBox.Show("Pengaturan tidak valid: " & ex.Message, "Watermark", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End Try

        Dim errors As List(Of String) = settings.Validate()
        If errors.Count > 0 Then
            MessageBox.Show(String.Join(Environment.NewLine, errors.ToArray()), "Pengaturan belum lengkap", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If folderBox.Text.Trim().Length = 0 OrElse Not Directory.Exists(folderBox.Text.Trim()) Then
            MessageBox.Show("Pilih folder yang berisi file .cdr terlebih dahulu.", "Batch", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim mode As String = If(saveInPlaceCheck.Checked, "file asli akan DITIMPA", "hasil disimpan sebagai nama_wm.cdr")
        If MessageBox.Show("Jalankan watermark ke semua .cdr di:" & Environment.NewLine & folderBox.Text & Environment.NewLine & Environment.NewLine &
                           "Mode simpan: " & mode & Environment.NewLine & Environment.NewLine & "Lanjutkan?",
                           "Batch watermark", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.Yes Then
            Return
        End If

        SetBusy(True)
        Try
            Dim engine As New WatermarkEngine()
            AddHandler engine.Progress, Sub(message As String) UpdateStatus(message)
            Dim result As WatermarkResult = engine.ApplyToFolder(
                corelApp, folderBox.Text.Trim(), settings, subfolderCheck.Checked, saveInPlaceCheck.Checked)
            reportBox.Text = "BATCH — " & settings.Describe() & Environment.NewLine & Environment.NewLine & result.ToReport()
            UpdateStatus("Batch selesai: " & result.DocumentsTouched & " dokumen tersimpan, " & result.Failures.Count & " gagal.")
        Catch ex As Exception
            MessageBox.Show("Batch gagal: " & ex.Message, "Watermark", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            SetBusy(False)
        End Try
    End Sub

    ' ========================================================================
    '  HELPER
    ' ========================================================================

    Private Function ReadSettings() As WatermarkSettings
        Dim s As New WatermarkSettings()
        s.Kind = CType(kindBox.SelectedIndex, WatermarkKind)
        s.Text = textBox.Text
        s.FontName = If(fontBox.Text.Trim().Length = 0, "Segoe UI", fontBox.Text.Trim())
        s.FontSize = CSng(fontSizeBox.Value)
        s.Color = selectedColor
        s.LogoPath = If(bundledLogoCheck.Checked, EnsureBundledLogoFile(), logoPathBox.Text.Trim())
        s.LogoWidthMm = CDbl(logoWidthBox.Value)
        s.OpacityPercent = CInt(opacityBox.Value)
        s.RotationDeg = CDbl(rotationBox.Value)
        s.Anchor = CType(If(anchorBox.SelectedIndex < 0, 4, anchorBox.SelectedIndex), WatermarkAnchor)
        s.MarginMm = CDbl(marginBox.Value)
        s.Tile = tileCheck.Checked
        s.TileSpacingXMm = CDbl(tileXBox.Value)
        s.TileSpacingYMm = CDbl(tileYBox.Value)
        s.ReplaceExisting = replaceCheck.Checked
        s.PerPageText = perPageCheck.Checked
        s.LockLayer = lockLayerCheck.Checked
        s.LayerName = If(layerNameBox.Text.Trim().Length = 0, "WATERMARK", layerNameBox.Text.Trim())
        ' index 0 = semua halaman, 1 = halaman aktif, 2 = halaman yang dicentang
        s.AllPages = (scopeBox.SelectedIndex = 0)
        s.TargetPages.Clear()
        If scopeBox.SelectedIndex = 2 Then
            For Each p As Integer In preselectedPages
                s.TargetPages.Add(p)
            Next
        End If
        Return s
    End Function

    Private Function GetActiveDocument() As Object
        If corelApp Is Nothing Then
            Throw New InvalidOperationException("Belum terhubung ke CorelDRAW. Buka CorelDRAW dan dokumen CDR terlebih dahulu.")
        End If
        Dim doc As Object = corelApp.ActiveDocument
        If doc Is Nothing Then
            Throw New InvalidOperationException("Tidak ada dokumen aktif di CorelDRAW.")
        End If
        Return doc
    End Function

    Private Sub SetBusy(value As Boolean)
        busy = value
        Cursor = If(value, Cursors.WaitCursor, Cursors.Default)
        Application.DoEvents()
    End Sub

    Private Sub UpdateStatus(message As String)
        statusLabel.Text = message
        Application.DoEvents()
    End Sub
End Class
