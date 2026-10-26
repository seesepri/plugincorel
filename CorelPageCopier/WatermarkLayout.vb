Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  WatermarkLayout.vb
'  Bagian "murni" fitur watermark: model pengaturan + matematika penempatan.
'  File ini TIDAK menyentuh COM / WinForms sehingga bisa diuji terpisah.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO

''' <summary>Jenis objek watermark yang dibuat.</summary>
Public Enum WatermarkKind
    TextOnly = 0
    LogoOnly = 1
    TextAndLogo = 2
End Enum

''' <summary>Posisi jangkar watermark pada halaman.</summary>
Public Enum WatermarkAnchor
    TopLeft = 0
    TopCenter = 1
    TopRight = 2
    MiddleLeft = 3
    Center = 4
    MiddleRight = 5
    BottomLeft = 6
    BottomCenter = 7
    BottomRight = 8
End Enum

''' <summary>Semua parameter watermark. Satu objek = satu "preset" yang bisa disimpan.</summary>
Public Class WatermarkSettings

    ' --- Objek watermark -----------------------------------------------------
    Public Kind As WatermarkKind = WatermarkKind.TextOnly
    Public Text As String = "CONTOH - WATERMARK"
    Public FontName As String = "Segoe UI"
    Public FontSize As Single = 24.0F
    Public Color As Color = Color.FromArgb(120, 120, 120)

    Public LogoPath As String = ""
    Public LogoWidthMm As Double = 35.0            ' 0 = pakai ukuran asli file

    ' --- Gaya ----------------------------------------------------------------
    Public OpacityPercent As Integer = 25          ' 1..100 (100 = pekat)
    Public RotationDeg As Double = 0.0             ' derajat, berlawanan arah jarum jam

    ' --- Penempatan ----------------------------------------------------------
    Public Anchor As WatermarkAnchor = WatermarkAnchor.Center
    Public MarginMm As Double = 8.0
    Public Tile As Boolean = False
    Public TileSpacingXMm As Double = 60.0
    Public TileSpacingYMm As Double = 45.0

    ' --- Perilaku dokumen ----------------------------------------------------
    Public AllPages As Boolean = True
    Public TargetPages As New List(Of Integer)()    ' dipakai bila AllPages = False
    Public ReplaceExisting As Boolean = True        ' hapus watermark lama dulu (idempoten)
    Public LayerName As String = WatermarkLayout.DefaultLayerName
    Public GroupWithPrefix As Boolean = True        ' grup teks+logo jadi 1 shape bergrup
    Public LockLayer As Boolean = False             ' kunci layer setelah diterapkan

    ' --- Tambahan ------------------------------------------------------------
    Public PerPageText As Boolean = False           ' tambahkan "Hal. n" di belakang teks

    ''' <summary>Label untuk ComboBox jangkar (urutan = nilai enum WatermarkAnchor).</summary>
    Public Shared Function AnchorLabels() As String()
        Return New String() {
            "Kiri atas", "Tengah atas", "Kanan atas",
            "Kiri tengah", "Tengah halaman", "Kanan tengah",
            "Kiri bawah", "Tengah bawah", "Kanan bawah"
        }
    End Function

    ''' <summary>Validasi sebelum menyentuh dokumen. Mengembalikan daftar pesan kesalahan.</summary>
    Public Function Validate() As List(Of String)
        Dim errors As New List(Of String)()

        If Kind = WatermarkKind.TextOnly OrElse Kind = WatermarkKind.TextAndLogo Then
            If Text Is Nothing OrElse Text.Trim().Length = 0 Then
                errors.Add("Teks watermark masih kosong.")
            End If
            If FontSize < 1.0F OrElse FontSize > 2000.0F Then
                errors.Add("Ukuran font harus antara 1 dan 2000 pt.")
            End If
            If FontName Is Nothing OrElse FontName.Trim().Length = 0 Then
                errors.Add("Nama font belum diisi.")
            End If
        End If

        If Kind = WatermarkKind.LogoOnly OrElse Kind = WatermarkKind.TextAndLogo Then
            If LogoPath Is Nothing OrElse LogoPath.Trim().Length = 0 Then
                errors.Add("File logo belum dipilih.")
            ElseIf Not File.Exists(LogoPath) Then
                errors.Add("File logo tidak ditemukan: " & LogoPath)
            End If
            If LogoWidthMm < 0 OrElse LogoWidthMm > 5000 Then
                errors.Add("Lebar logo harus antara 0 dan 5000 mm (0 = ukuran asli).")
            End If
        End If

        If OpacityPercent < 1 OrElse OpacityPercent > 100 Then
            errors.Add("Opasitas harus antara 1% dan 100%.")
        End If

        If MarginMm < 0 Then errors.Add("Margin tidak boleh negatif.")
        If Tile AndAlso (TileSpacingXMm <= 0 OrElse TileSpacingYMm <= 0) Then
            errors.Add("Jarak antar-watermark (tile) harus lebih besar dari 0.")
        End If

        If Not AllPages Then
            If TargetPages Is Nothing OrElse TargetPages.Count = 0 Then
                errors.Add("Tidak ada halaman target yang dipilih.")
            End If
        End If

        Return errors
    End Function

    ''' <summary>Ringkasan satu baris untuk log / laporan.</summary>
    Public Function Describe() As String
        Dim sb As New System.Text.StringBuilder()
        Select Case Kind
            Case WatermarkKind.TextOnly
                sb.Append("Teks """ & Text & """")
            Case WatermarkKind.LogoOnly
                sb.Append("Logo " & Path.GetFileName(LogoPath))
            Case Else
                sb.Append("Teks """ & Text & """ + logo " & Path.GetFileName(LogoPath))
        End Select
        sb.Append(" | opasitas " & OpacityPercent & "%")
        If Math.Abs(RotationDeg) > 0.01 Then sb.Append(" | rotasi " & RotationDeg & "°")
        If Tile Then
            sb.Append(" | tile " & TileSpacingXMm & "×" & TileSpacingYMm & " mm")
        Else
            sb.Append(" | jangkar " & AnchorLabels()(CInt(Anchor)))
        End If
        sb.Append(" | layer """ & LayerName & """")
        Return sb.ToString()
    End Function
End Class

''' <summary>Perhitungan posisi. Semua nilai dalam satuan dokumen (mm pada fitur ini).</summary>
Public Class WatermarkLayout

    ''' <summary>Awalan nama shape watermark, dipakai untuk cari/hapus/ganti.</summary>
    Public Const ShapePrefix As String = "WM_"

    ''' <summary>Nama layer default tempat watermark diletakkan.</summary>
    Public Const DefaultLayerName As String = "WATERMARK"

    ''' <summary>True bila nama shape adalah objek watermark (dipakai Auto Select &amp; hapus).</summary>
    Public Shared Function IsWatermarkObjectName(shapeName As String) As Boolean
        If shapeName Is Nothing Then Return False
        Return shapeName.StartsWith(ShapePrefix, StringComparison.Ordinal)
    End Function

    ''' <summary>True bila nama layer adalah layer watermark.</summary>
    Public Shared Function IsWatermarkLayerName(layerName As String) As Boolean
        If layerName Is Nothing Then Return False
        Return String.Equals(layerName.Trim(), DefaultLayerName, StringComparison.OrdinalIgnoreCase)
    End Function
    ''' <summary>
    ''' Identitas watermark berdasarkan nama shape ATAU nama layer-nya.
    ''' Dipakai bersama oleh Auto Select (agar watermark tidak ikut terpilih),
    ''' penghapusan watermark, dan penggantian watermark.
    ''' </summary>
    Public Shared Function IsWatermarkIdentity(shapeName As String, layerName As String) As Boolean
        Return IsWatermarkObjectName(shapeName) OrElse IsWatermarkLayerName(layerName)
    End Function


    ''' <summary>Konversi opasitas (%) menjadi nilai transparency CorelDRAW (0..100).</summary>
    Public Shared Function TransparencyFromOpacity(opacityPercent As Integer) As Integer
        Dim o As Integer = opacityPercent
        If o < 0 Then o = 0
        If o > 100 Then o = 100
        Return 100 - o
    End Function

    ''' <summary>
    ''' Koordinat kiri-bawah untuk watermark tunggal.
    ''' CorelDRAW memakai titik origin di kiri-bawah, jadi "bottom" = jarak dari bawah halaman.
    ''' </summary>
    Public Shared Function AnchorLeftBottom(pageWidth As Double, pageHeight As Double,
                                             objectWidth As Double, objectHeight As Double,
                                             anchor As WatermarkAnchor, marginMm As Double) As PointF
        Dim m As Double = If(marginMm < 0, 0, marginMm)
        Dim left As Double = m
        Dim bottom As Double = m

        Select Case anchor
            Case WatermarkAnchor.TopLeft
                left = m
                bottom = pageHeight - m - objectHeight
            Case WatermarkAnchor.TopCenter
                left = (pageWidth - objectWidth) / 2.0
                bottom = pageHeight - m - objectHeight
            Case WatermarkAnchor.TopRight
                left = pageWidth - m - objectWidth
                bottom = pageHeight - m - objectHeight
            Case WatermarkAnchor.MiddleLeft
                left = m
                bottom = (pageHeight - objectHeight) / 2.0
            Case WatermarkAnchor.Center
                left = (pageWidth - objectWidth) / 2.0
                bottom = (pageHeight - objectHeight) / 2.0
            Case WatermarkAnchor.MiddleRight
                left = pageWidth - m - objectWidth
                bottom = (pageHeight - objectHeight) / 2.0
            Case WatermarkAnchor.BottomLeft
                left = m
                bottom = m
            Case WatermarkAnchor.BottomCenter
                left = (pageWidth - objectWidth) / 2.0
                bottom = m
            Case WatermarkAnchor.BottomRight
                left = pageWidth - m - objectWidth
                bottom = m
        End Select

        Return New PointF(CSng(left), CSng(bottom))
    End Function

    ''' <summary>
    ''' Posisi kiri-bawah untuk mode tile (pola berulang) yang menutupi seluruh halaman.
    ''' Distribusi dibuat simetris sehingga margin kiri/kanan dan atas/bawah sama besar.
    ''' </summary>
    Public Shared Function TilePositions(pageWidth As Double, pageHeight As Double,
                                         objectWidth As Double, objectHeight As Double,
                                         spacingX As Double, spacingY As Double) As List(Of PointF)
        Dim result As New List(Of PointF)()

        If pageWidth <= 0 OrElse pageHeight <= 0 OrElse objectWidth <= 0 OrElse objectHeight <= 0 Then
            Return result
        End If

        ' Jarak minimum: 110% lebar / 140% tinggi objek, supaya watermark tidak menumpuk.
        Dim stepX As Double = Math.Max(spacingX, objectWidth * 1.1)
        Dim stepY As Double = Math.Max(spacingY, objectHeight * 1.4)

        Dim nx As Integer = Math.Max(1, CInt(Math.Ceiling(pageWidth / stepX)))
        Dim ny As Integer = Math.Max(1, CInt(Math.Ceiling(pageHeight / stepY)))

        Dim startX As Double = (pageWidth - ((nx - 1) * stepX) - objectWidth) / 2.0
        Dim startY As Double = (pageHeight - ((ny - 1) * stepY) - objectHeight) / 2.0

        For j As Integer = 0 To ny - 1
            For i As Integer = 0 To nx - 1
                Dim left As Double = startX + i * stepX
                Dim bottom As Double = startY + j * stepY
                ' Buang posisi yang sama sekali di luar halaman.
                If left + objectWidth > 0 AndAlso left < pageWidth AndAlso
                   bottom + objectHeight > 0 AndAlso bottom < pageHeight Then
                    result.Add(New PointF(CSng(left), CSng(bottom)))
                End If
            Next
        Next

        ' Objek lebih besar dari halaman: tetap beri satu watermark di tengah.
        If result.Count = 0 Then
            result.Add(New PointF(CSng((pageWidth - objectWidth) / 2.0),
                                  CSng((pageHeight - objectHeight) / 2.0)))
        End If

        Return result
    End Function
    ' ========================================================================
    '  VERIFIKASI GEOMETRI (dipakai CopyClicked)
    ' ========================================================================

    ''' <summary>Selisih posisi & ukuran objek hasil operasi terhadap objek sumber (mm).</summary>
    Public Structure GeometryDelta
        Public DX As Double
        Public DY As Double
        Public DW As Double
        Public DH As Double

        Public Function Matches(toleranceMm As Double) As Boolean
            Return Math.Abs(DX) <= toleranceMm AndAlso
                   Math.Abs(DY) <= toleranceMm AndAlso
                   Math.Abs(DW) <= toleranceMm AndAlso
                   Math.Abs(DH) <= toleranceMm
        End Function

        Public Function Describe() As String
            Return "Δx " & Fmt(DX) & ", Δy " & Fmt(DY) & ", Δw " & Fmt(DW) & ", Δh " & Fmt(DH)
        End Function

        Private Shared Function Fmt(value As Double) As String
            Return String.Format("{0:+0.0;-0.0;0.0} mm", value)
        End Function
    End Structure

    ''' <summary>
    ''' Bandingkan kotak pembatas objek sumber dengan objek hasil (posisi tengah + ukuran).
    ''' Nilai harus dalam satuan yang sama (mm selama operasi copy).
    ''' </summary>
    Public Shared Function CompareGeometry(srcCenterX As Double, srcCenterY As Double,
                                           srcWidth As Double, srcHeight As Double,
                                           dstCenterX As Double, dstCenterY As Double,
                                           dstWidth As Double, dstHeight As Double) As GeometryDelta
        Dim delta As New GeometryDelta()
        delta.DX = dstCenterX - srcCenterX
        delta.DY = dstCenterY - srcCenterY
        delta.DW = dstWidth - srcWidth
        delta.DH = dstHeight - srcHeight
        Return delta
    End Function
End Class

''' <summary>
''' Penjaga re-entrancy untuk operasi panjang (paste banyak halaman, pindai
''' semua halaman, batch). Kelas murni tanpa COM/WinForms sehingga bisa diuji
''' terpisah dari UI.
'''
''' Pola pakai:
'''     If Not gate.TryBegin("Salin objek") Then Return
'''     Try
'''         ' ...operasi panjang...
'''     Finally
'''         gate.Release()
'''     End Try
''' </summary>
Public Class OperationGate
    Private _busy As Boolean = False
    Private _currentOperation As String = ""

    Public ReadOnly Property IsBusy As Boolean
        Get
            Return _busy
        End Get
    End Property

    Public ReadOnly Property CurrentOperation As String
        Get
            Return _currentOperation
        End Get
    End Property

    Public ReadOnly Property BusyMessage As String
        Get
            If Not _busy Then Return ""
            Return "Masih memproses: " & _currentOperation & ". Tunggu hingga selesai."
        End Get
    End Property

    ''' <summary>
    ''' Ambil kunci operasi. False = masih ada operasi lain yang berjalan, sehingga
    ''' pemanggil wajib berhenti tanpa melakukan apa pun.
    ''' </summary>
    Public Function TryBegin(operationName As String) As Boolean
        If _busy Then Return False
        _busy = True
        _currentOperation = If(operationName, "")
        Return True
    End Function

    ''' <summary>Lepas kunci. Aman dipanggil walau tidak sedang sibuk.</summary>
    Public Sub Release()
        _busy = False
        _currentOperation = ""
    End Sub
End Class
