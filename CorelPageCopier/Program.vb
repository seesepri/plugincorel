Option Explicit On
Option Strict Off
Option Infer On
Imports System
Imports System.Collections.Generic
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

    Public Sub New()
        Text = "Corel Page Copier"
        StartPosition = FormStartPosition.CenterScreen
        Size = New Size(600, 820)
        MinimumSize = New Size(560, 740)
        BackColor = Color.FromArgb(245, 247, 250)
        Font = New Font("Segoe UI", 9.0F)
        BuildUi()
        AddHandler Shown, AddressOf FormShown
    End Sub

    Private Sub BuildUi()
        SuspendLayout()
        Controls.Clear()

        Dim root As New TableLayoutPanel With {
            .Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 4,
            .Padding = New Padding(18), .BackColor = BackColor
        }
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 86.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 56.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 44.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 62.0F))
        Controls.Add(root)

        Dim header As New Panel With {.Dock = DockStyle.Fill, .BackColor = BackColor}
        Dim title As New Label With {.Text = "Corel Page Copier", .Font = New Font("Segoe UI Semibold", 19.0F, FontStyle.Bold), .ForeColor = Color.FromArgb(25, 35, 52), .AutoSize = True, .Location = New Point(0, 0)}
        Dim subtitle As New Label With {.Text = "Salin desain dan pilih objek dengan lebih cepat", .Font = New Font("Segoe UI", 9.5F), .ForeColor = Color.FromArgb(102, 115, 132), .AutoSize = True, .Location = New Point(2, 36)}
        sourceLabel.SetBounds(2, 59, 520, 22)
        sourceLabel.Font = New Font("Segoe UI", 8.5F)
        sourceLabel.ForeColor = Color.FromArgb(65, 95, 125)
        header.Controls.AddRange(New Control() {title, subtitle, sourceLabel})
        root.Controls.Add(header, 0, 0)

        Dim copyCard As Panel = CreateCard()
        copyCard.Margin = New Padding(0, 0, 0, 12)
        Dim copyTitle As Label = CreateSectionTitle("01  Salin ke halaman")
        copyTitle.Location = New Point(16, 14)
        copyCard.Controls.Add(copyTitle)
        Dim hint As New Label With {.Text = "Pilih halaman tujuan. Halaman aktif adalah sumber dan tidak bisa dipilih.", .AutoSize = False, .Location = New Point(18, 43), .Size = New Size(500, 32), .ForeColor = Color.FromArgb(102, 115, 132)}
        copyCard.Controls.Add(hint)
        pagesList.SetBounds(18, 78, 500, 145)
        pagesList.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        pagesList.CheckOnClick = True
        pagesList.IntegralHeight = False
        pagesList.BorderStyle = BorderStyle.FixedSingle
        pagesList.BackColor = Color.White
        pagesList.Font = New Font("Segoe UI", 9.0F)
        copyCard.Controls.Add(pagesList)

        Dim pageActions As New FlowLayoutPanel With {.Location = New Point(16, 230), .Size = New Size(500, 34), .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right, .FlowDirection = FlowDirection.LeftToRight, .WrapContents = False, .BackColor = Color.White}
        Dim refresh As Button = CreateButton("Muat ulang", False, 92)
        AddHandler refresh.Click, AddressOf RefreshClicked
        Dim all As Button = CreateButton("Pilih semua", False, 98)
        AddHandler all.Click, Sub() SetAllChecks(True)
        Dim none As Button = CreateButton("Kosongkan", False, 94)
        AddHandler none.Click, Sub() SetAllChecks(False)
        pageActions.Controls.AddRange(New Control() {refresh, all, none})
        copyCard.Controls.Add(pageActions)
        Dim copy As Button = CreateButton("Salin ke halaman terpilih", True, 500)
        copy.SetBounds(18, 270, 500, 38)
        copy.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        AddHandler copy.Click, AddressOf CopyClicked
        copyCard.Controls.Add(copy)
        root.Controls.Add(copyCard, 0, 1)

        Dim selectCard As Panel = CreateCard()
        Dim selectTitle As Label = CreateSectionTitle("02  Auto Select objek")
        selectTitle.Location = New Point(16, 14)
        selectCard.Controls.Add(selectTitle)
        Dim categoryLabel As New Label With {.Text = "Jenis objek", .AutoSize = True, .Location = New Point(18, 48), .ForeColor = Color.FromArgb(70, 82, 99)}
        selectCard.Controls.Add(categoryLabel)
        categoryBox.DropDownStyle = ComboBoxStyle.DropDownList
        categoryBox.SetBounds(18, 69, 500, 30)
        categoryBox.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        categoryBox.Items.Clear()
        categoryBox.Items.AddRange(New Object() {"Teks (Artistik + Paragraf)", "Gambar / Bitmap", "Garis / Kurva terbuka", "Objek vektor (tanpa teks/bitmap/grup)", "Semua objek"})
        categoryBox.SelectedIndex = 0
        selectCard.Controls.Add(categoryBox)
        Dim scopeLabel As New Label With {.Text = "Cakupan pencarian", .AutoSize = True, .Location = New Point(18, 108), .ForeColor = Color.FromArgb(70, 82, 99)}
        selectCard.Controls.Add(scopeLabel)
        scopeBox.DropDownStyle = ComboBoxStyle.DropDownList
        scopeBox.SetBounds(18, 129, 500, 30)
        scopeBox.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        scopeBox.Items.Clear()
        scopeBox.Items.AddRange(New Object() {"Halaman aktif", "Pindai semua halaman"})
        scopeBox.SelectedIndex = 0
        selectCard.Controls.Add(scopeBox)
        Dim autoSelect As Button = CreateButton("Select otomatis", True, 500)
        autoSelect.SetBounds(18, 171, 500, 36)
        autoSelect.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        AddHandler autoSelect.Click, AddressOf AutoSelectClicked
        selectCard.Controls.Add(autoSelect)
        Dim nextMatch As Button = CreateButton("Pilih hasil berikutnya", False, 500)
        nextMatch.SetBounds(18, 215, 500, 34)
        nextMatch.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        AddHandler nextMatch.Click, AddressOf NextMatchingPageClicked
        selectCard.Controls.Add(nextMatch)
        Dim note As New Label With {.Text = "Pindai semua halaman untuk mencari kecocokan di seluruh dokumen. Hasil dipilih satu halaman setiap kali.", .Location = New Point(18, 257), .Size = New Size(500, 34), .ForeColor = Color.FromArgb(102, 115, 132), .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right}
        selectCard.Controls.Add(note)
        root.Controls.Add(selectCard, 0, 2)

        Dim statusCard As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.FromArgb(232, 241, 250), .Padding = New Padding(12), .Margin = New Padding(0, 10, 0, 0)}
        statusLabel.Dock = DockStyle.Fill
        statusLabel.TextAlign = ContentAlignment.MiddleLeft
        statusLabel.Font = New Font("Segoe UI", 9.0F)
        statusLabel.ForeColor = Color.FromArgb(35, 71, 105)
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
        Try
            LoadPages()
        Catch ex As Exception
            ShowError("Gagal memuat ulang daftar halaman: " & ex.Message)
        End Try
    End Sub

    Private Sub CopyClicked(sender As Object, e As EventArgs)
        Dim targets As New List(Of Integer)()
        For Each obj As Object In pagesList.CheckedItems
            Dim p = DirectCast(obj, PageItem)
            If p.Index <> sourcePageIndex Then targets.Add(p.Index)
        Next
        If targets.Count = 0 Then
            MessageBox.Show("Centang setidaknya satu halaman tujuan selain halaman sumber.", "Corel Page Copier", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
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

            ' Copy berjalan langsung tanpa dialog konfirmasi. Objek sumber tidak pernah dihapus.
            corelDoc.Pages(actualSourcePage).Activate()
            selection.Copy()
            Dim succeeded As Integer = 0
            Dim failures As New List(Of String)(validationFailures)
            For targetPosition As Integer = 0 To validatedTargets.Count - 1
                Dim pageIndex As Integer = validatedTargets(targetPosition)
                statusLabel.Text = "Copy progres: " & (targetPosition + 1) & "/" & validatedTargets.Count & " — halaman tujuan " & pageIndex & "."
                Application.DoEvents()
                Try
                    Dim destination As Object = corelDoc.Pages(pageIndex)
                    destination.Activate()
                    Dim beforeCount As Integer = CInt(destination.Shapes.Count)
                    destination.ActiveLayer.Paste()
                    Dim afterCount As Integer = CInt(destination.Shapes.Count)
                    If afterCount <= beforeCount Then
                        failures.Add("Halaman " & pageIndex & ": Paste tidak menambah shape tingkat atas yang terdeteksi; periksa hasil secara manual.")
                    Else
                        succeeded += 1
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
            Dim report As String = "Copy selesai." & Microsoft.VisualBasic.Constants.vbCrLf & "Halaman tujuan valid diproses: " & validatedTargets.Count & "." & Microsoft.VisualBasic.Constants.vbCrLf & "Berhasil diverifikasi: " & succeeded & " halaman." & Microsoft.VisualBasic.Constants.vbCrLf & "Gagal/perlu diperiksa: " & failures.Count & " halaman."
            If failures.Count > 0 Then report &= Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf & String.Join(Microsoft.VisualBasic.Constants.vbCrLf, failures.ToArray())
            report &= Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf & "Catatan: verifikasi berdasarkan perubahan jumlah shape tingkat atas; periksa hasil visual, khususnya grup, layer khusus, dan dokumen dengan objek bersarang."
            statusLabel.Text = "Copy selesai: " & succeeded & "/" & validatedTargets.Count & " halaman terverifikasi; " & failures.Count & " gagal/perlu diperiksa."
            MessageBox.Show(report, "Hasil Corel Page Copier", MessageBoxButtons.OK, If(failures.Count = 0, MessageBoxIcon.Information, MessageBoxIcon.Warning))
        Catch ex As Exception
            ShowError("Proses copy gagal: " & ex.Message)
        End Try
    End Sub

    Private Sub AutoSelectClicked(sender As Object, e As EventArgs)
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
                Application.DoEvents()
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
            Dim report As String = "Pemindaian semua halaman selesai." & Microsoft.VisualBasic.Constants.vbCrLf & "Total objek cocok: " & total & Microsoft.VisualBasic.Constants.vbCrLf & "Halaman yang memiliki kecocokan: " & matchingPageIndices.Count & Microsoft.VisualBasic.Constants.vbCrLf & "Seleksi saat ini: halaman " & matchingPageIndices(0) & ". Gunakan tombol 'Pilih Halaman Cocok Berikutnya' untuk berpindah." & Microsoft.VisualBasic.Constants.vbCrLf & Microsoft.VisualBasic.Constants.vbCrLf & String.Join(Microsoft.VisualBasic.Constants.vbCrLf, details.ToArray())
            statusLabel.Text = "Pindai semua halaman: " & total & " objek; " & matchingPageIndices.Count & " halaman cocok."
            MessageBox.Show(report, "Auto Select — Semua Halaman", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            ShowError("Auto Select gagal: " & ex.Message)
        End Try
    End Sub

    Private Sub NextMatchingPageClicked(sender As Object, e As EventArgs)
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
