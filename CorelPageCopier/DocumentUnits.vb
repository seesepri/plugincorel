Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  DocumentUnits.vb
'  Normalisasi satuan dokumen (Temuan B).
'
'  Masalah pada kode lama:
'    - `previousUnit = CInt(corelDoc.Unit)` dibungkus Try/Catch kosong, lalu
'      `corelDoc.Unit = mm` juga dibungkus Try/Catch kosong. Jika salah satunya
'      gagal, verifikasi TETAP berjalan dengan asumsi milimeter -> hasil salah
'      tetapi dilaporkan rapi.
'    - Kegagalan memulihkan satuan ditelan diam-diam.
'
'  Aturan baru:
'    1. Satuan asli WAJIB terbaca. Bila tidak, operasi dihentikan sebelum
'       dokumen diubah (tidak ada paste), dengan pesan eksplisit.
'    2. Jika penyetelan ke mm gagal, dipakai KONVERSI EKSPLISIT dari satuan
'       asli ke mm (faktor di tabel MillimetersPerUnit).
'    3. Jika satuan asli tidak dikenal (termasuk pixel yang bergantung resolusi),
'       operasi dihentikan - tidak ada asumsi milimeter.
'    4. Pemulihan dilakukan di Finally dan kegagalannya DILAPORKAN.
'
'  Seluruh kelas di file ini murni (tanpa COM) sehingga bisa diuji unit.
' ============================================================================

Imports System
Imports System.Globalization

''' <summary>
''' Tabel satuan CorelDRAW (cdrUnit) + konversi eksplisit ke milimeter.
''' Nilai enum cdrUnit yang stabil dan didokumentasikan Corel:
'''   1 = inch, 2 = centimeter, 3 = millimeter, 4 = point,
'''   5 = pica, 6 = didot, 7 = cicero, 8 = pixel.
''' </summary>
Public Class DocumentUnits

    Public Const Unknown As Integer = -1
    Public Const Inch As Integer = 1
    Public Const Centimeter As Integer = 2
    Public Const Millimeter As Integer = 3
    Public Const Point As Integer = 4
    Public Const Pica As Integer = 5
    Public Const Didot As Integer = 6
    Public Const Cicero As Integer = 7
    Public Const Pixel As Integer = 8

    ''' <summary>
    ''' Faktor mm per 1 satuan dokumen. Nilai 0 = tidak dapat dikonversi
    ''' (mis. pixel, karena bergantung pada resolusi dokumen).
    ''' </summary>
    Public Shared Function MillimetersPerUnit(unit As Integer) As Double
        Select Case unit
            Case Millimeter
                Return 1.0
            Case Centimeter
                Return 10.0
            Case Inch
                Return 25.4
            Case Point
                Return 25.4 / 72.0          ' 0.352777... mm
            Case Pica
                Return 25.4 / 6.0           ' 12 pt = 4.2333... mm
            Case Didot
                Return 0.376                   ' 1 didot ≈ 0.376 mm
            Case Cicero
                Return 4.5115                  ' 12 didot ≈ 4.5115 mm
            Case Else
                Return 0.0                     ' pixel / tidak dikenal
        End Select
    End Function

    ''' <summary>True bila nilai dalam satuan ini bisa dikonversi eksplisit ke mm.</summary>
    Public Shared Function IsConvertible(unit As Integer) As Boolean
        Return MillimetersPerUnit(unit) > 0.0
    End Function

    ''' <summary>
    ''' Konversi eksplisit ke mm. Mengembalikan Double.NaN bila satuan tidak
    ''' dapat dikonversi - pemanggil wajib memeriksa (jangan asumsikan mm).
    ''' </summary>
    Public Shared Function ToMillimeters(value As Double, unit As Integer) As Double
        Dim factor As Double = MillimetersPerUnit(unit)
        If factor <= 0.0 Then Return Double.NaN
        Return value * factor
    End Function

    Public Shared Function ToUnit(valueMm As Double, unit As Integer) As Double
        Dim factor As Double = MillimetersPerUnit(unit)
        If factor <= 0.0 Then Return Double.NaN
        Return valueMm / factor
    End Function

    ''' <summary>Nama pendek satuan untuk laporan.</summary>
    Public Shared Function UnitName(unit As Integer) As String
        Select Case unit
            Case Millimeter
                Return "mm"
            Case Centimeter
                Return "cm"
            Case Inch
                Return "inch"
            Case Point
                Return "pt"
            Case Pica
                Return "pica"
            Case Didot
                Return "didot"
            Case Cicero
                Return "cicero"
            Case Pixel
                Return "pixel"
            Case Else
                Return "tidak dikenal (" & unit.ToString(CultureInfo.InvariantCulture) & ")"
        End Select
    End Function
End Class

''' <summary>
''' State machine normalisasi satuan untuk satu operasi.
''' Pemanggil (Program.vb / WatermarkEngine.vb) melakukan panggilan COM, kelas
''' ini mencatat hasilnya dan memutuskan apakah verifikasi masih boleh dipercaya.
''' </summary>
Public Class DocumentUnitScope

    Private _originalUnit As Integer = DocumentUnits.Unknown
    Private _originalKnown As Boolean = False
    Private _originalError As String = ""
    Private _changeApplied As Boolean = False
    Private _changeError As String = ""
    Private _restoreAttempted As Boolean = False
    Private _restored As Boolean = False
    Private _restoreError As String = ""

    Public Sub New(Optional targetUnit As Integer = DocumentUnits.Millimeter)
        Me.TargetUnit = targetUnit
    End Sub

    ''' <summary>Satuan yang ingin dipakai selama operasi (default mm).</summary>
    Public ReadOnly Property TargetUnit As Integer

    Public ReadOnly Property OriginalUnit As Integer
        Get
            Return _originalUnit
        End Get
    End Property

    Public ReadOnly Property OriginalUnitKnown As Boolean
        Get
            Return _originalKnown
        End Get
    End Property

    ''' <summary>Pesan kegagalan membaca Document.Unit (kosong bila berhasil).</summary>
    Public ReadOnly Property OriginalError As String
        Get
            Return _originalError
        End Get
    End Property

    ''' <summary>True bila Document.Unit berhasil diubah ke TargetUnit.</summary>
    Public ReadOnly Property ChangeApplied As Boolean
        Get
            Return _changeApplied
        End Get
    End Property

    Public ReadOnly Property ChangeError As String
        Get
            Return _changeError
        End Get
    End Property

    ''' <summary>
    ''' Satuan yang benar-benar dipakai saat mengukur shape.
    ''' TargetUnit bila perubahan berhasil, selain itu satuan asli.
    ''' </summary>
    Public ReadOnly Property MeasurementUnit As Integer
        Get
            If Not _originalKnown Then Return DocumentUnits.Unknown
            If _changeApplied Then Return Me.TargetUnit
            Return _originalUnit
        End Get
    End Property

    ''' <summary>Faktor mm per satuan pengukuran (0 = tidak dapat dikonversi).</summary>
    Public ReadOnly Property MeasurementFactorMm As Double
        Get
            Return DocumentUnits.MillimetersPerUnit(MeasurementUnit)
        End Get
    End Property

    ''' <summary>True bila hasil pengukuran bisa dibandingkan dalam mm.</summary>
    Public ReadOnly Property MeasurementsComparable As Boolean
        Get
            Return MeasurementFactorMm > 0.0
        End Get
    End Property

    ''' <summary>True bila satuan dokumen perlu dikembalikan di Finally.</summary>
    Public ReadOnly Property RestoreRequired As Boolean
        Get
            Return _changeApplied AndAlso _originalKnown
        End Get
    End Property

    Public ReadOnly Property RestoreAttempted As Boolean
        Get
            Return _restoreAttempted
        End Get
    End Property

    Public ReadOnly Property RestoreSucceeded As Boolean
        Get
            Return _restored
        End Get
    End Property

    ''' <summary>True bila pemulihan satuan GAGAL (wajib dilaporkan ke pengguna).</summary>
    Public ReadOnly Property RestoreFailed As Boolean
        Get
            Return _restoreAttempted AndAlso Not _restored
        End Get
    End Property

    Public ReadOnly Property RestoreError As String
        Get
            Return _restoreError
        End Get
    End Property

    ''' <summary>
    ''' True bila operasi boleh dilanjutkan: satuan asli terbaca dan hasil ukur
    ''' dapat dikonversi ke mm. False = verifikasi TIDAK boleh mengasumsikan mm.
    ''' </summary>
    Public ReadOnly Property CanVerifyInMillimeters As Boolean
        Get
            Return _originalKnown AndAlso MeasurementsComparable
        End Get
    End Property

    ''' <summary>Alasan operasi harus dihentikan (kosong = aman dilanjutkan).</summary>
    Public ReadOnly Property BlockingReason As String
        Get
            If Not _originalKnown Then
                Return "Satuan dokumen (Document.Unit) tidak dapat dibaca" &
                       If(String.IsNullOrEmpty(_originalError), "", ": " & _originalError) &
                       ". Operasi dibatalkan agar verifikasi tidak memakai asumsi milimeter."
            End If
            If Not MeasurementsComparable Then
                Return "Satuan dokumen """ & DocumentUnits.UnitName(_originalUnit) &
                       """ tidak dapat dikonversi ke milimeter secara eksplisit" &
                       If(String.IsNullOrEmpty(_changeError), "", " (penyetelan mm juga gagal: " & _changeError & ")") &
                       ". Operasi dibatalkan agar verifikasi tidak memakai asumsi milimeter."
            End If
            Return ""
        End Get
    End Property

    ' --- pencatatan hasil panggilan COM -------------------------------------

    Public Sub MarkOriginalKnown(unit As Integer)
        _originalUnit = unit
        _originalKnown = True
        _originalError = ""
    End Sub

    Public Sub MarkOriginalUnknown(errorMessage As String)
        _originalUnit = DocumentUnits.Unknown
        _originalKnown = False
        _originalError = If(errorMessage, "")
    End Sub

    Public Sub MarkChangeApplied()
        _changeApplied = True
        _changeError = ""
    End Sub

    Public Sub MarkChangeFailed(errorMessage As String)
        _changeApplied = False
        _changeError = If(errorMessage, "")
    End Sub

    Public Sub MarkRestored()
        _restoreAttempted = True
        _restored = True
        _restoreError = ""
    End Sub

    Public Sub MarkRestoreFailed(errorMessage As String)
        _restoreAttempted = True
        _restored = False
        _restoreError = If(errorMessage, "")
    End Sub

    ''' <summary>Catatan satuan untuk laporan (selalu menyebut cara pengukuran).</summary>
    Public Function Describe() As String
        If Not _originalKnown Then
            Return "Satuan dokumen TIDAK dapat dibaca — verifikasi geometri dibatalkan."
        End If

        Dim factorText As String = Me.MeasurementFactorMm.ToString("0.####", CultureInfo.InvariantCulture)
        If _changeApplied Then
            If _restored Then
                Return "Satuan dokumen dinormalkan ke mm selama operasi; satuan asli (" &
                       DocumentUnits.UnitName(_originalUnit) & ") berhasil dipulihkan."
            ElseIf _restoreAttempted Then
                Return "Satuan dokumen dinormalkan ke mm, tetapi PEMULIHAN ke " &
                       DocumentUnits.UnitName(_originalUnit) & " GAGAL: " & _restoreError &
                       ". Ubah satuan dokumen kembali secara manual!"
            Else
                Return "Satuan dokumen dinormalkan ke mm selama operasi; pemulihan belum dijalankan."
            End If
        End If

        If MeasurementsComparable Then
            Return "Document.Unit tidak dapat diubah ke mm (" & _changeError &
                   "); pengukuran dilakukan dalam " & DocumentUnits.UnitName(_originalUnit) &
                   " dan dikonversi eksplisit ke mm (1 " & DocumentUnits.UnitName(_originalUnit) &
                   " = " & factorText & " mm)."
        End If

        Return "Satuan dokumen """ & DocumentUnits.UnitName(_originalUnit) &
               """ tidak dapat dikonversi ke mm — verifikasi geometri tidak dapat dipercaya."
    End Function
End Class
