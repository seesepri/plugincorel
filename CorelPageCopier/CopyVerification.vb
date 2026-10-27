Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  CopyVerification.vb
'  Logika verifikasi hasil "Salin ke halaman" yang DIPISAH dari COM.
'
'  Alasan file ini terpisah:
'    - Temuan A (audit reliabilitas): kode lama menyatakan copy "terverifikasi"
'      hanya karena jumlah shape bertambah, dan membandingkan geometri sumber
'      dengan shape TERBAWAH pada layer (indeks 1..n), bukan objek hasil paste
'      yang berada di ATAS tumpukan (z-order).
'    - Semua perhitungan di sini murni (tanpa COM / WinForms) sehingga bisa
'      diuji unit tanpa CorelDRAW.
'
'  Status yang dipakai:
'    Verified : geometri (dan struktur bila tersedia) cocok dalam toleransi.
'    Partial  : paste terjadi, tetapi pengukuran/struktur tidak lengkap
'               -> TIDAK boleh dilaporkan sebagai sukses penuh.
'    Failed   : tidak ada objek hasil, atau geometri jelas berbeda.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text

''' <summary>Status verifikasi satu halaman hasil copy.</summary>
Public Enum CopyVerifyStatus
    Verified = 0
    ' 'Partial' adalah kata kunci VB, jadi dinamai PartiallyVerified.
    PartiallyVerified = 1
    Failed = 2
End Enum

''' <summary>Kotak pembatas (bounding box) dalam milimeter, sumbu Y ke atas.</summary>
Public Structure BoundingBox
    Public MinX As Double
    Public MinY As Double
    Public MaxX As Double
    Public MaxY As Double
    Private _valid As Boolean

    ''' <summary>Kotak kosong (belum valid).</summary>
    Public Shared Function Empty() As BoundingBox
        Return New BoundingBox()
    End Function

    ''' <summary>Bangun kotak dari titik tengah + ukuran (sama seperti properti COM CenterX/SizeWidth).</summary>
    Public Shared Function FromCenter(centerX As Double, centerY As Double,
                                      width As Double, height As Double) As BoundingBox
        Dim box As New BoundingBox()
        box.MinX = centerX - width / 2.0
        box.MaxX = centerX + width / 2.0
        box.MinY = centerY - height / 2.0
        box.MaxY = centerY + height / 2.0
        box._valid = True
        Return box
    End Function

    ''' <summary>Bangun kotak dari batas kiri/bawah/kanan/atas.</summary>
    Public Shared Function FromBounds(minX As Double, minY As Double, maxX As Double, maxY As Double) As BoundingBox
        Dim box As New BoundingBox()
        box.MinX = Math.Min(minX, maxX)
        box.MinY = Math.Min(minY, maxY)
        box.MaxX = Math.Max(minX, maxX)
        box.MaxY = Math.Max(minY, maxY)
        box._valid = True
        Return box
    End Function

    Public ReadOnly Property IsValid As Boolean
        Get
            Return _valid
        End Get
    End Property

    Public ReadOnly Property CenterX As Double
        Get
            Return (MinX + MaxX) / 2.0
        End Get
    End Property

    Public ReadOnly Property CenterY As Double
        Get
            Return (MinY + MaxY) / 2.0
        End Get
    End Property

    Public ReadOnly Property Width As Double
        Get
            Return MaxX - MinX
        End Get
    End Property

    Public ReadOnly Property Height As Double
        Get
            Return MaxY - MinY
        End Get
    End Property

    ''' <summary>Gabungkan dua kotak. Kotak tidak valid diabaikan.</summary>
    Public Function Union(other As BoundingBox) As BoundingBox
        If Not _valid Then Return other
        If Not other._valid Then Return Me
        Dim box As BoundingBox = Me
        box.MinX = Math.Min(Me.MinX, other.MinX)
        box.MinY = Math.Min(Me.MinY, other.MinY)
        box.MaxX = Math.Max(Me.MaxX, other.MaxX)
        box.MaxY = Math.Max(Me.MaxY, other.MaxY)
        box._valid = True
        Return box
    End Function

    ''' <summary>Skala semua nilai (dipakai untuk konversi satuan eksplisit ke mm).</summary>
    Public Function Scale(factor As Double) As BoundingBox
        If Not _valid Then Return Me
        Dim box As New BoundingBox()
        box.MinX = Me.MinX * factor
        box.MaxX = Me.MaxX * factor
        box.MinY = Me.MinY * factor
        box.MaxY = Me.MaxY * factor
        box._valid = True
        Return box
    End Function

    Public Function Describe() As String
        If Not _valid Then Return "(tidak terukur)"
        Return "pusat (" & Fmt(CenterX) & ", " & Fmt(CenterY) & ") ukuran " &
               Fmt(Width) & " × " & Fmt(Height)
    End Function

    Private Shared Function Fmt(value As Double) As String
        Return value.ToString("0.00", CultureInfo.InvariantCulture)
    End Function
End Structure

''' <summary>
''' Sidik jari struktur hasil copy: jumlah objek + sebaran kode tipe shape.
''' Dipakai untuk memastikan jumlah/tipe objek hasil paste sesuai sumber
''' (mis. grup bersarang tetap 1 grup, multi-selection tetap 3 objek).
''' </summary>
Public Class ShapeSignature
    Private ReadOnly _typeCounts As New Dictionary(Of Integer, Integer)()
    Private _count As Integer = 0
    Private _readable As Boolean = True

    ''' <summary>Jumlah objek yang berhasil dibaca.</summary>
    Public ReadOnly Property Count As Integer
        Get
            Return _count
        End Get
    End Property

    ''' <summary>False bila ada objek yang tipenya tidak dapat dibaca.</summary>
    Public ReadOnly Property IsReadable As Boolean
        Get
            Return _readable AndAlso _count > 0
        End Get
    End Property

    ''' <summary>True bila pembandingan struktur mungkin dilakukan.</summary>
    Public ReadOnly Property IsComparable As Boolean
        Get
            Return IsReadable
        End Get
    End Property

    Public Sub AddShapeType(typeCode As Integer)
        _count += 1
        Dim current As Integer = 0
        If _typeCounts.TryGetValue(typeCode, current) Then
            _typeCounts(typeCode) = current + 1
        Else
            _typeCounts(typeCode) = 1
        End If
    End Sub

    ''' <summary>Objek yang tipenya gagal dibaca tetap dihitung, tetapi tanda "tidak terbaca" diset.</summary>
    Public Sub MarkUnreadable()
        _readable = False
    End Sub

    Public Shared Function FromTypeCodes(typeCodes As Integer()) As ShapeSignature
        Dim signature As New ShapeSignature()
        If typeCodes Is Nothing Then Return signature
        For Each code As Integer In typeCodes
            signature.AddShapeType(code)
        Next
        Return signature
    End Function

    ''' <summary>Bandingkan jumlah objek dan sebaran tipenya.</summary>
    Public Function Matches(other As ShapeSignature) As Boolean
        If other Is Nothing Then Return False
        If Not IsComparable OrElse Not other.IsComparable Then Return False
        If _count <> other._count Then Return False
        If _typeCounts.Count <> other._typeCounts.Count Then Return False
        For Each pair As KeyValuePair(Of Integer, Integer) In _typeCounts
            Dim otherCount As Integer = 0
            If Not other._typeCounts.TryGetValue(pair.Key, otherCount) Then Return False
            If otherCount <> pair.Value Then Return False
        Next
        Return True
    End Function

    Public Function Describe() As String
        If _count = 0 Then Return "(tidak ada objek)"
        Dim parts As New List(Of String)()
        For Each pair As KeyValuePair(Of Integer, Integer) In _typeCounts
            parts.Add("tipe " & pair.Key.ToString(CultureInfo.InvariantCulture) & "×" &
                      pair.Value.ToString(CultureInfo.InvariantCulture))
        Next
        parts.Sort(StringComparer.Ordinal)
        Return _count.ToString(CultureInfo.InvariantCulture) & " objek (" &
               String.Join(", ", parts.ToArray()) & ")"
    End Function
End Class

''' <summary>Hasil satu pemeriksaan verifikasi.</summary>
Public Class CopyVerificationResult
    Public Status As CopyVerifyStatus = CopyVerifyStatus.Failed
    Public Reason As String = ""
    Public HasGeometry As Boolean = False
    Public HasStructure As Boolean = False
    Public Delta As WatermarkLayout.GeometryDelta

    Public ReadOnly Property IsVerified As Boolean
        Get
            Return Status = CopyVerifyStatus.Verified
        End Get
    End Property

    Public Function Describe() As String
        Dim label As String
        Select Case Status
            Case CopyVerifyStatus.Verified
                label = "Verified"
            Case CopyVerifyStatus.PartiallyVerified
                label = "Partial"
            Case Else
                label = "Failed"
        End Select
        If String.IsNullOrEmpty(Reason) Then Return label
        Return label & " — " & Reason
    End Function
End Class

''' <summary>
''' Aturan verifikasi copy. Semua keputusan status ada di sini supaya bisa diuji
''' unit tanpa CorelDRAW (Temuan A).
''' </summary>
Public Class CopyVerification

    ''' <summary>
    ''' Toleransi geometri terdokumentasi: 0,5 mm per sumbu untuk selisih titik
    ''' tengah MAUPUN ukuran. Nilai ini dipilih karena:
    '''   - 0,5 mm ≈ 1,4 pt, di bawah ambang yang terlihat mata pada halaman A4;
    '''   - menutupi pembulatan internal CorelDRAW saat paste antarhalaman.
    ''' Semua nilai dibandingkan DALAM MM (lihat DocumentUnits.ToMillimeters).
    ''' </summary>
    Public Const ToleranceMm As Double = 0.5

    ''' <summary>
    ''' Klasifikasi satu halaman.
    ''' addedCount        : jumlah shape yang bertambah pada layer tujuan.
    ''' geometryAvailable : kotak sumber DAN hasil paste berhasil diukur.
    ''' delta             : selisih pusat/ukuran (mm) bila geometryAvailable.
    ''' structureAvailable: jumlah + tipe objek sumber dan hasil bisa dibaca.
    ''' structureMatches  : sidik jari struktur sama.
    ''' </summary>
    Public Shared Function Classify(addedCount As Integer,
                                    geometryAvailable As Boolean,
                                    delta As WatermarkLayout.GeometryDelta,
                                    structureAvailable As Boolean,
                                    structureMatches As Boolean,
                                    Optional toleranceMm As Double = ToleranceMm) As CopyVerificationResult
        Dim result As New CopyVerificationResult()
        result.HasGeometry = geometryAvailable
        result.HasStructure = structureAvailable
        result.Delta = delta

        ' 1. Tidak ada objek hasil paste -> gagal, titik.
        If addedCount <= 0 Then
            result.Status = CopyVerifyStatus.Failed
            result.Reason = "paste tidak menambah objek pada layer tujuan"
            Return result
        End If

        ' 2. Geometri terukur tetapi di luar toleransi -> gagal (posisi/ukuran salah).
        If geometryAvailable Then
            If Not delta.Matches(toleranceMm) Then
                result.Status = CopyVerifyStatus.Failed
                result.Reason = "posisi/ukuran objek hasil berbeda dari sumber (" & delta.Describe() & ")"
                Return result
            End If

            If structureAvailable AndAlso Not structureMatches Then
                result.Status = CopyVerifyStatus.PartiallyVerified
                result.Reason = "geometri cocok, tetapi struktur objek berbeda dari sumber"
                Return result
            End If

            result.Status = CopyVerifyStatus.Verified
            If structureAvailable Then
                result.Reason = "geometri & struktur cocok (" & delta.Describe() & ")"
            Else
                result.Reason = "geometri cocok (" & delta.Describe() & "); struktur objek tidak dapat dibandingkan"
            End If
            Return result
        End If

        ' 3. Geometri tidak tersedia -> TIDAK PERNAH Verified.
        If structureAvailable AndAlso structureMatches Then
            result.Status = CopyVerifyStatus.PartiallyVerified
            result.Reason = "jumlah/tipe objek cocok, tetapi geometri tidak dapat diukur"
            Return result
        End If

        result.Status = CopyVerifyStatus.PartiallyVerified
        result.Reason = "objek bertambah, tetapi geometri dan struktur tidak dapat diverifikasi"
        Return result
    End Function
    ''' <summary>
    ''' Verifikasi satu halaman hasil paste: konversi satuan eksplisit ke mm,
    ''' bandingkan geometri, bandingkan struktur, lalu klasifikasikan.
    '''
    ''' Inilah komposisi yang dipakai langsung oleh CopyClickedCore, sehingga unit
    ''' test menguji jalur yang sama dengan aplikasi (bukan salinan logika).
    '''
    ''' unitFactorMm = faktor konversi satuan pengukuran ke mm
    '''                (1,0 bila Document.Unit berhasil di-set ke mm).
    ''' </summary>
    Public Shared Function VerifyPasted(sourceSet As MeasuredShapeSet,
                                        pastedSet As MeasuredShapeSet,
                                        addedCount As Integer,
                                        unitFactorMm As Double,
                                        Optional toleranceMm As Double = ToleranceMm) As CopyVerificationResult
        Dim factor As Double = If(unitFactorMm > 0.0, unitFactorMm, 0.0)
        Dim sourceMm As MeasuredShapeSet = sourceSet
        Dim pastedMm As MeasuredShapeSet = pastedSet
        If factor > 0.0 Then
            sourceMm = sourceSet.Scaled(factor)
            pastedMm = pastedSet.Scaled(factor)
        End If

        Dim geometryAvailable As Boolean = sourceMm.HasGeometry AndAlso pastedMm.HasGeometry
        Dim delta As New WatermarkLayout.GeometryDelta()
        If geometryAvailable Then
            delta = WatermarkLayout.CompareGeometry(
                sourceMm.Bounds.CenterX, sourceMm.Bounds.CenterY,
                sourceMm.Bounds.Width, sourceMm.Bounds.Height,
                pastedMm.Bounds.CenterX, pastedMm.Bounds.CenterY,
                pastedMm.Bounds.Width, pastedMm.Bounds.Height)
        End If

        Dim structureAvailable As Boolean = sourceSet.HasStructure AndAlso pastedSet.HasStructure
        Dim structureMatches As Boolean = structureAvailable AndAlso sourceSet.Signature.Matches(pastedSet.Signature)

        Dim result As CopyVerificationResult = Classify(addedCount, geometryAvailable, delta,
                                                        structureAvailable, structureMatches, toleranceMm)

        ' Sebab teknis pengukuran yang gagal ikut dilaporkan (transparansi).
        If Not geometryAvailable AndAlso pastedSet IsNot Nothing AndAlso
           Not String.IsNullOrEmpty(pastedSet.Reason) Then
            result.Reason = result.Reason & " (" & pastedSet.Reason & ")"
        End If
        Return result
    End Function

End Class

''' <summary>Hasil satu halaman dalam satu operasi copy.</summary>
Public Class CopyPageOutcome
    Public PageIndex As Integer = 0
    Public Status As CopyVerifyStatus = CopyVerifyStatus.Failed
    Public Reason As String = ""

    Public Sub New()
    End Sub

    Public Sub New(pageIndex As Integer, status As CopyVerifyStatus, reason As String)
        Me.PageIndex = pageIndex
        Me.Status = status
        Me.Reason = If(reason, "")
    End Sub

    Public Function StatusLabel() As String
        Select Case Status
            Case CopyVerifyStatus.Verified
                Return "Verified"
            Case CopyVerifyStatus.PartiallyVerified
                Return "Partial"
            Case Else
                Return "Failed"
        End Select
    End Function
End Class

''' <summary>
''' Ringkasan satu operasi copy + pembentuk teks laporan.
''' Satu sumber status untuk MessageBox DAN status bar (Temuan A), sehingga
''' halaman Partial tidak pernah masuk hitungan sukses.
''' </summary>
Public Class CopyRunSummary
    Public Pages As New List(Of CopyPageOutcome)()
    Public Warnings As New List(Of String)()

    Public ReadOnly Property TotalCount As Integer
        Get
            Return Pages.Count
        End Get
    End Property

    Public ReadOnly Property VerifiedCount As Integer
        Get
            Return CountOf(CopyVerifyStatus.Verified)
        End Get
    End Property

    Public ReadOnly Property PartialCount As Integer
        Get
            Return CountOf(CopyVerifyStatus.PartiallyVerified)
        End Get
    End Property

    Public ReadOnly Property FailedCount As Integer
        Get
            Return CountOf(CopyVerifyStatus.Failed)
        End Get
    End Property

    ''' <summary>True hanya bila SEMUA halaman Verified (dipakai untuk ikon/judul laporan).</summary>
    Public ReadOnly Property AllVerified As Boolean
        Get
            Return Pages.Count > 0 AndAlso VerifiedCount = Pages.Count
        End Get
    End Property

    Public Function CountOf(status As CopyVerifyStatus) As Integer
        Dim count As Integer = 0
        For Each page As CopyPageOutcome In Pages
            If page.Status = status Then count += 1
        Next
        Return count
    End Function

    Public Sub Add(pageIndex As Integer, status As CopyVerifyStatus, reason As String)
        Pages.Add(New CopyPageOutcome(pageIndex, status, reason))
    End Sub

    ''' <summary>Satu baris ringkas untuk status bar.</summary>
    Public Function Headline(durationSeconds As Double) As String
        Return "Copy selesai: " & VerifiedCount & " verified, " & PartialCount & " partial, " &
               FailedCount & " gagal dari " & TotalCount & " halaman (" &
               durationSeconds.ToString("0.0", CultureInfo.InvariantCulture) & " detik)."
    End Function

    ''' <summary>
    ''' Teks laporan lengkap. documentIdentity/unitNote/undoNote diisi pemanggil
    ''' supaya laporan tidak pernah menyembunyikan kegagalan pemulihan state.
    ''' </summary>
    Public Function BuildReport(documentIdentity As String, unitNote As String,
                                undoNote As String, durationSeconds As Double) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("Copy selesai.")
        If Not String.IsNullOrEmpty(documentIdentity) Then sb.AppendLine("Dokumen        : " & documentIdentity)
        sb.AppendLine("Halaman diproses: " & TotalCount)
        sb.AppendLine("  Verified      : " & VerifiedCount & " halaman (geometri cocok dalam toleransi)")
        sb.AppendLine("  Partial       : " & PartialCount & " halaman (paste terjadi, verifikasi tidak lengkap)")
        sb.AppendLine("  Gagal         : " & FailedCount & " halaman")
        If Not String.IsNullOrEmpty(unitNote) Then sb.AppendLine()
        If Not String.IsNullOrEmpty(unitNote) Then sb.AppendLine(unitNote)
        If Not String.IsNullOrEmpty(undoNote) Then sb.AppendLine(undoNote)

        Dim problemPages As New List(Of String)()
        For Each page As CopyPageOutcome In Pages
            If page.Status <> CopyVerifyStatus.Verified Then
                problemPages.Add("  • Halaman " & page.PageIndex & " [" & page.StatusLabel() & "]: " & page.Reason)
            End If
        Next
        If problemPages.Count > 0 Then
            sb.AppendLine()
            sb.AppendLine("Perlu diperiksa:")
            For Each line As String In problemPages
                sb.AppendLine(line)
            Next
        End If

        If Warnings.Count > 0 Then
            sb.AppendLine()
            sb.AppendLine("Peringatan:")
            For Each warning As String In Warnings
                sb.AppendLine("  • " & warning)
            Next
        End If

        sb.AppendLine()
        sb.AppendLine("Durasi: " & durationSeconds.ToString("0.0", CultureInfo.InvariantCulture) & " detik.")
        Return sb.ToString()
    End Function
End Class
