Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  WatermarkPlacementCheck.vb
'  Verifikasi penempatan watermark: harapan (dihitung) vs kenyataan (dibaca
'  kembali dari objek di CorelDRAW).
'
'  Kenapa perlu:
'    Kode lama menempatkan objek lalu langsung melaporkan sukses. Bila
'    CorelDRAW membalik/menggeser objek (mis. karena unit dokumen belum
'    benar-benar berubah, halaman landscape, atau objek lebih besar dari
'    halaman), hasilnya tetap dilaporkan "berhasil" tanpa bukti.
'
'  Semua fungsi di sini murni (tanpa COM / WinForms) sehingga bisa diuji:
'    - posisi jangkar 3×3 + margin, termasuk pembersihan (clamp) agar objek
'      tidak keluar halaman,
'    - bounding box objek yang diputar (rotasi),
'    - konversi satuan mm / inch / point sebelum dibandingkan,
'    - jumlah posisi tile yang diharapkan.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Globalization

''' <summary>Selisih penempatan objek terhadap perhitungan (semua dalam mm).</summary>
Public Class PlacementVerification
    Public DX As Double = 0
    Public DY As Double = 0
    Public DW As Double = 0
    Public DH As Double = 0
    ''' <summary>True bila objek yang diukur memang ada dan terbaca.</summary>
    Public Measured As Boolean = False
    Public ToleranceMm As Double = 0.5
    Public Problems As New List(Of String)()

    Public ReadOnly Property WithinTolerance As Boolean
        Get
            If Not Measured Then Return False
            If Problems.Count > 0 Then Return False
            Return Math.Abs(DX) <= ToleranceMm AndAlso Math.Abs(DY) <= ToleranceMm AndAlso
                   Math.Abs(DW) <= ToleranceMm AndAlso Math.Abs(DH) <= ToleranceMm
        End Get
    End Property

    Public Function Describe() As String
        If Not Measured Then Return "penempatan tidak dapat diukur"
        Dim text As String = "Δpusat " & Format(DX) & " / " & Format(DY) &
                             ", Δukuran " & Format(DW) & " / " & Format(DH) &
                             " (toleransi " & Format(ToleranceMm) & ")"
        If Problems.Count > 0 Then
            text = text & " — " & String.Join("; ", Problems.ToArray())
        End If
        Return text
    End Function

    Private Shared Function Format(value As Double) As String
        Return String.Format(CultureInfo.InvariantCulture, "{0:+0.00;-0.00;0.00} mm", value)
    End Function
End Class

''' <summary>Pemeriksaan penempatan watermark (murni).</summary>
Public Class WatermarkPlacementCheck

    ''' <summary>Toleransi bawaan pembacaan posisi objek dari CorelDRAW (mm).</summary>
    Public Const DefaultToleranceMm As Double = 0.5

    ''' <summary>
    ''' Bounding box sejajar sumbu dari objek berukuran w×h yang diputar `angleDeg`
    ''' derajat. CorelDRAW melaporkan SizeWidth/SizeHeight sebagai bounding box
    ''' sejajar sumbu, jadi perhitungan ini yang harus dipakai sebelum menaruh objek.
    ''' </summary>
    Public Shared Function RotatedBounds(width As Double, height As Double, angleDeg As Double) As SizeInfo
        Dim result As New SizeInfo()
        If width <= 0 OrElse height <= 0 Then Return result
        Dim radians As Double = angleDeg * Math.PI / 180.0
        Dim cosValue As Double = Math.Abs(Math.Cos(radians))
        Dim sinValue As Double = Math.Abs(Math.Sin(radians))
        result.Width = width * cosValue + height * sinValue
        result.Height = width * sinValue + height * cosValue
        Return result
    End Function

    ''' <summary>
    ''' Titik kiri-bawah yang diharapkan untuk mode jangkar. Memakai
    ''' WatermarkLayout.AnchorLeftBottom (satu sumber matematika) sehingga test
    ''' dan mesin memakai rumus yang sama.
    ''' </summary>
    Public Shared Function ExpectedLeftBottom(pageWidthMm As Double, pageHeightMm As Double,
                                              objectWidthMm As Double, objectHeightMm As Double,
                                              anchor As WatermarkAnchor, marginMm As Double) As PointF
        Return WatermarkLayout.AnchorLeftBottom(pageWidthMm, pageHeightMm,
                                                objectWidthMm, objectHeightMm, anchor, marginMm)
    End Function

    ''' <summary>Titik pusat yang diharapkan (dipakai mesin saat menempatkan objek).</summary>
    Public Shared Function ExpectedCenter(pageWidthMm As Double, pageHeightMm As Double,
                                          objectWidthMm As Double, objectHeightMm As Double,
                                          anchor As WatermarkAnchor, marginMm As Double) As PointF
        Dim leftBottom As PointF = ExpectedLeftBottom(pageWidthMm, pageHeightMm,
                                                     objectWidthMm, objectHeightMm, anchor, marginMm)
        Return New PointF(CSng(leftBottom.X + objectWidthMm / 2.0),
                          CSng(leftBottom.Y + objectHeightMm / 2.0))
    End Function

    ''' <summary>
    ''' Bandingkan pusat & ukuran objek hasil dengan harapan. Nilai harus dalam mm.
    ''' </summary>
    Public Shared Function Verify(expectedCenter As PointF, expectedWidthMm As Double, expectedHeightMm As Double,
                                  actualCenterX As Double, actualCenterY As Double,
                                  actualWidthMm As Double, actualHeightMm As Double,
                                  Optional toleranceMm As Double = DefaultToleranceMm) As PlacementVerification
        Dim verification As New PlacementVerification()
        verification.ToleranceMm = If(toleranceMm > 0, toleranceMm, DefaultToleranceMm)

        If expectedWidthMm <= 0 OrElse expectedHeightMm <= 0 Then
            verification.Problems.Add("ukuran objek yang diharapkan tidak valid (0).")
            Return verification
        End If
        If actualWidthMm <= 0 OrElse actualHeightMm <= 0 Then
            verification.Problems.Add("ukuran objek hasil tidak terbaca (0).")
            Return verification
        End If

        verification.Measured = True
        verification.DX = actualCenterX - expectedCenter.X
        verification.DY = actualCenterY - expectedCenter.Y
        verification.DW = actualWidthMm - expectedWidthMm
        verification.DH = actualHeightMm - expectedHeightMm
        Return verification
    End Function

    ''' <summary>
    ''' Apakah objek muat di halaman (tidak terpotong). Objek yang lebih besar
    ''' dari halaman dianggap tidak muat meskipun sudah dipusatkan.
    ''' </summary>
    Public Shared Function FitsOnPage(pageWidthMm As Double, pageHeightMm As Double,
                                      objectWidthMm As Double, objectHeightMm As Double) As Boolean
        If pageWidthMm <= 0 OrElse pageHeightMm <= 0 Then Return False
        Return objectWidthMm <= pageWidthMm + 0.001 AndAlso objectHeightMm <= pageHeightMm + 0.001
    End Function

    ''' <summary>Jumlah posisi tile yang diharapkan untuk halaman & jarak tertentu.</summary>
    Public Shared Function ExpectedTileCount(pageWidthMm As Double, pageHeightMm As Double,
                                             objectWidthMm As Double, objectHeightMm As Double,
                                             spacingXMm As Double, spacingYMm As Double) As Integer
        If pageWidthMm <= 0 OrElse pageHeightMm <= 0 OrElse objectWidthMm <= 0 OrElse objectHeightMm <= 0 Then
            Return 0
        End If
        Return WatermarkLayout.TilePositions(pageWidthMm, pageHeightMm, objectWidthMm, objectHeightMm,
                                             spacingXMm, spacingYMm).Count
    End Function

    ''' <summary>
    ''' Ukuran halaman dalam mm dengan konversi satuan eksplisit. Dipakai untuk
    ''' memastikan penempatan memakai ukuran halaman AKTUAL dan satuan yang
    ''' konsisten (mm / inch / point), bukan asumsi mm.
    ''' </summary>
    Public Shared Function PageSizeMillimeters(sizeWidthValue As Double, sizeHeightValue As Double,
                                               unit As Integer) As SizeInfo
        Dim result As New SizeInfo()
        result.Width = DocumentUnits.ToMillimeters(sizeWidthValue, unit)
        result.Height = DocumentUnits.ToMillimeters(sizeHeightValue, unit)
        Return result
    End Function

    ''' <summary>True bila konversi satuan mungkin (mm/inch/point) dan hasilnya masuk akal.</summary>
    Public Shared Function PageSizeUsable(pageSize As SizeInfo) As Boolean
        If Double.IsNaN(pageSize.Width) OrElse Double.IsNaN(pageSize.Height) Then Return False
        Return pageSize.Width > 0 AndAlso pageSize.Height > 0
    End Function

    ''' <summary>Deskripsi singkat orientasi halaman (untuk laporan).</summary>
    Public Shared Function OrientationLabel(pageWidthMm As Double, pageHeightMm As Double) As String
        If Math.Abs(pageWidthMm - pageHeightMm) < 0.01 Then Return "persegi"
        Return If(pageWidthMm > pageHeightMm, "landscape", "portrait")
    End Function
End Class
