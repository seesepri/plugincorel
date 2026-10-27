Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  CopyInputKindTests.vb
'  Review lanjutan PR #2 — pengukuran objek sumber.
'
'  Fokus: membedakan
'    (a) Shape tunggal yang MEMANG tidak punya Count,
'    (b) ShapeRange,
'    (c) koleksi COM lain yang punya Count,
'    (d) objek yang Count-nya GAGAL dibaca.
'
'  Kesalahan yang dijaga: kegagalan membaca Count tidak boleh dianggap sebagai
'  bukti bahwa input adalah satu Shape — kalau itu terjadi, geometri "kebetulan"
'  terbaca dan verifikasi bisa mengklaim Verified padahal jumlah objek tidak
'  diketahui. Dalam kondisi itu status wajib Partial.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Runtime.InteropServices

''' <summary>Shape tunggal palsu: TIDAK memiliki anggota Count (mirip CorelDRAW.Shape).</summary>
Public Class FakeSingleShapeInput
    Public Property Type As Integer = 6
    Public Property CenterX As Double = 0
    Public Property CenterY As Double = 0
    Public Property SizeWidth As Double = 0
    Public Property SizeHeight As Double = 0

    Public Sub New(typeCode As Integer, cx As Double, cy As Double, w As Double, h As Double)
        Me.Type = typeCode
        Me.CenterX = cx
        Me.CenterY = cy
        Me.SizeWidth = w
        Me.SizeHeight = h
    End Sub
End Class

''' <summary>ShapeRange palsu: punya Count + indexer 1-based.</summary>
Public Class FakeShapeRangeInput
    Private ReadOnly _items As New List(Of FakeSingleShapeInput)()

    Public ReadOnly Property Count As Integer
        Get
            Return _items.Count
        End Get
    End Property

    Public Default ReadOnly Property Item(oneBasedIndex As Integer) As FakeSingleShapeInput
        Get
            Return _items(oneBasedIndex - 1)
        End Get
    End Property

    Public Sub Add(shape As FakeSingleShapeInput)
        _items.Add(shape)
    End Sub
End Class

''' <summary>Koleksi COM generik (bukan ShapeRange) yang Count-nya valid.</summary>
Public Class FakeCollectionInput
    Private ReadOnly _items As New List(Of FakeSingleShapeInput)()

    Public ReadOnly Property Count As Integer
        Get
            Return _items.Count
        End Get
    End Property

    Public Default ReadOnly Property Item(oneBasedIndex As Integer) As FakeSingleShapeInput
        Get
            Return _items(oneBasedIndex - 1)
        End Get
    End Property

    Public Sub Add(shape As FakeSingleShapeInput)
        _items.Add(shape)
    End Sub
End Class

''' <summary>
''' Koleksi yang Count-nya GAGAL dibaca (mis. RPC sibuk / COM bermasalah),
''' tetapi properti geometri tetap terbaca. Inilah objek yang dulu salah
''' dianggap \"satu Shape\" sehingga geometrinya dipakai sebagai bukti.
''' </summary>
Public Class FakeBrokenCountInput
    Public Property CenterX As Double = 500
    Public Property CenterY As Double = 500
    Public Property SizeWidth As Double = 40
    Public Property SizeHeight As Double = 20
    Public Property Type As Integer = 3

    Public ReadOnly Property Count As Integer
        Get
            Throw New COMException("simulasi: RPC gagal saat membaca Count", &H8001010E)
        End Get
    End Property
End Class

''' <summary>
''' Objek yang Count-nya gagal dibaca DAN geometrinya tidak terbaca:
''' tidak ada satu pun fakta yang bisa dipercaya.
''' </summary>
Public Class FakeBrokenEverythingInput
    Public ReadOnly Property Count As Integer
        Get
            Throw New COMException("simulasi: COM sudah tidak valid", &H80010108)
        End Get
    End Property

    Public ReadOnly Property CenterX As Double
        Get
            Throw New InvalidOperationException("simulasi: geometri tidak terbaca")
        End Get
    End Property

    Public ReadOnly Property Type As Integer
        Get
            Throw New InvalidOperationException("simulasi: tipe tidak terbaca")
        End Get
    End Property
End Class

Public Module CopyInputKindTests

    Private Function NewMeasuredSource(kindInput As Object) As MeasuredShapeSet
        Return CopyMeasurement.MeasureSelection(kindInput)
    End Function

    ''' <summary>Objek hasil paste palsu dengan geometri tertentu.</summary>
    Private Function NewPasted(cx As Double, cy As Double, w As Double, h As Double,
                               typeCode As Integer) As MeasuredShapeSet
        Dim items As New List(Of FakeShape)()
        items.Add(New FakeShape(typeCode, cx, cy, w, h))
        Return CopyMeasurement.MeasureTopShapes(New FakeShapeCollection(items), 1)
    End Function

    Public Sub Register(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("Pengukuran input: Shape vs ShapeRange vs koleksi vs gagal")

        ' ------------------------------------------------------------------
        ' 1. TryReadCount membedakan \"tidak punya Count\" dari \"gagal dibaca\"
        ' ------------------------------------------------------------------
        suite.Add("TryReadCount: Shape tunggal -> MemberMissing (bukan Failed)", Sub()
            Dim shape As New FakeSingleShapeInput(6, 10, 10, 4, 4)
            Dim read As CountReadResult = CopyMeasurement.TryReadCount(shape)
            Check.AreEqual(CountReadStatus.MemberMissing, read.Status, "status MemberMissing")
            Check.IsTrue(read.IsMemberMissing, "objek memang tidak punya Count")
            Check.IsFalse(read.IsOk, "bukan hasil bacaan sukses")
            Check.Contains(read.Describe(), "tidak memiliki anggota Count", "penjelasan")
        End Sub)

        suite.Add("TryReadCount: Count gagal dibaca -> Failed (jumlah tidak diketahui)", Sub()
            Dim broken As New FakeBrokenCountInput()
            Dim read As CountReadResult = CopyMeasurement.TryReadCount(broken)
            Check.AreEqual(CountReadStatus.Failed, read.Status, "status Failed")
            Check.IsFalse(read.IsMemberMissing, "BUKAN dianggap Shape tunggal")
            Check.Contains(read.Describe(), "tidak dapat dibaca", "penjelasan")
            Check.Contains(read.ErrorMessage, "RPC gagal", "pesan asli dilaporkan")
        End Sub)

        suite.Add("TryReadCount: exception bersarang tetap dikenali sebagai member hilang", Sub()
            Dim nested As New System.Reflection.TargetInvocationException(
                New MissingMemberException("anggota Count tidak ada"))
            Check.IsTrue(CopyMeasurement.IsMemberMissing(nested), "dibuka sampai InnerException")
        End Sub)

        ' ------------------------------------------------------------------
        ' 2. Klasifikasi jenis input
        ' ------------------------------------------------------------------
        suite.Add("ClassifyInput: Shape tunggal, ShapeRange, koleksi COM dibedakan", Sub()
            Dim shape As New FakeSingleShapeInput(6, 1, 1, 2, 2)
            Check.AreEqual(MeasurementInputKind.SingleShape, CopyMeasurement.ClassifyInput(shape),
                           "Shape tunggal")

            Dim rangeInput As New FakeShapeRangeInput()
            rangeInput.Add(New FakeSingleShapeInput(6, 1, 1, 2, 2))
            Check.AreEqual(MeasurementInputKind.ShapeRange, CopyMeasurement.ClassifyInput(rangeInput),
                           "ShapeRange (punya Count, nama tipe ShapeRange)")

            Dim collection As New FakeCollectionInput()
            collection.Add(New FakeSingleShapeInput(6, 1, 1, 2, 2))
            Check.AreEqual(MeasurementInputKind.ShapeCollection, CopyMeasurement.ClassifyInput(collection),
                           "koleksi COM biasa")

            Check.AreEqual(MeasurementInputKind.Empty, CopyMeasurement.ClassifyInput(Nothing), "kosong")
            Check.AreEqual(MeasurementInputKind.Unknown,
                           CopyMeasurement.ClassifyInput(New FakeBrokenCountInput()), "Count gagal dibaca")
        End Sub)

        ' ------------------------------------------------------------------
        ' 3. Jumlah objek diketahui -> semua objek diukur
        ' ------------------------------------------------------------------
        suite.Add("ShapeRange: semua objek diukur & jenisnya dicatat", Sub()
            Dim rangeInput As New FakeShapeRangeInput()
            rangeInput.Add(New FakeSingleShapeInput(6, 0, 0, 10, 10))
            rangeInput.Add(New FakeSingleShapeInput(5, 20, 0, 10, 10))

            Dim measured As MeasuredShapeSet = NewMeasuredSource(rangeInput)
            Check.IsTrue(measured.HasGeometry, "geometri terukur")
            Check.AreEqual(2, measured.Signature.Count, "dua objek diukur (bukan satu)")
            Check.Near(10, measured.Bounds.CenterX, 0.0001, "pusat gabungan")
            Check.Near(30, measured.Bounds.Width, 0.0001, "lebar gabungan")
            Check.AreEqual(MeasurementInputKind.ShapeRange, measured.InputKind, "jenis input dicatat")
            Check.Contains(measured.Source, "ShapeRange", "sumber disebut di laporan")
        End Sub)

        suite.Add("koleksi COM: semua objek diukur & jenisnya koleksi", Sub()
            Dim collection As New FakeCollectionInput()
            collection.Add(New FakeSingleShapeInput(3, 5, 5, 10, 10))
            collection.Add(New FakeSingleShapeInput(3, 15, 5, 10, 10))
            collection.Add(New FakeSingleShapeInput(3, 25, 5, 10, 10))

            Dim measured As MeasuredShapeSet = NewMeasuredSource(collection)
            Check.AreEqual(3, measured.Signature.Count, "tiga objek")
            Check.AreEqual(MeasurementInputKind.ShapeCollection, measured.InputKind, "jenis koleksi")
        End Sub)

        suite.Add("Shape tunggal tetap diukur seperti sebelumnya (perilaku lama dijaga)", Sub()
            Dim shape As New FakeSingleShapeInput(6, 5, 5, 4, 4)
            Dim measured As MeasuredShapeSet = NewMeasuredSource(shape)
            Check.IsTrue(measured.HasGeometry, "geometri terukur")
            Check.Near(5, measured.Bounds.CenterX, 0.0001, "centerX")
            Check.AreEqual(1, measured.Signature.Count, "satu objek")
            Check.AreEqual(MeasurementInputKind.SingleShape, measured.InputKind, "jenis input")
        End Sub)

        ' ------------------------------------------------------------------
        ' 4. Count gagal dibaca -> TIDAK dianggap satu Shape -> Partial
        ' ------------------------------------------------------------------
        suite.Add("Count gagal dibaca: geometri yang 'kebetulan' terbaca TIDAK dipakai", Sub()
            Dim broken As New FakeBrokenCountInput()
            Dim measured As MeasuredShapeSet = NewMeasuredSource(broken)

            Check.IsFalse(measured.HasGeometry, "tidak ada geometri yang diklaim")
            Check.IsFalse(measured.HasStructure, "struktur ditandai tidak terbaca")
            Check.AreEqual(MeasurementInputKind.Unknown, measured.InputKind, "jenis tidak dikenali")
            Check.Contains(measured.Reason, "jumlah objek seleksi tidak dapat diverifikasi", "alasan eksplisit")
        End Sub)

        suite.Add("Count gagal dibaca -> status verifikasi Partial (BUKAN Verified)", Sub()
            ' Objek sumber bermasalah, tetapi geometrinya sama persis dengan objek
            ' hasil paste. Versi lama akan mengukur 'satu shape' dan menyatakan
            ' Verified; sekarang harus Partial.
            Dim broken As New FakeBrokenCountInput()
            Dim source As MeasuredShapeSet = NewMeasuredSource(broken)
            Dim pasted As MeasuredShapeSet = NewPasted(500, 500, 40, 20, 3)

            Dim result As CopyVerificationResult = CopyVerification.VerifyPasted(source, pasted, 1, 1.0)
            Check.AreEqual(CopyVerifyStatus.PartiallyVerified, result.Status, "Partial, bukan Verified")
            Check.Contains(result.Reason, "tidak dapat", "alasan menyebut pengukuran tidak lengkap")
        End Sub)

        suite.Add("Count & geometri gagal -> Partial, tidak ada klaim sukses", Sub()
            Dim source As MeasuredShapeSet = NewMeasuredSource(New FakeBrokenEverythingInput())
            Dim pasted As MeasuredShapeSet = NewPasted(0, 0, 10, 10, 6)

            Dim result As CopyVerificationResult = CopyVerification.VerifyPasted(source, pasted, 1, 1.0)
            Check.AreEqual(CopyVerifyStatus.PartiallyVerified, result.Status, "Partial")
            Check.IsFalse(result.HasGeometry, "tidak ada geometri")
        End Sub)

        suite.Add("Shape tunggal tanpa tipe & geometri -> Partial (bukan Verified)", Sub()
            Dim unreadable As New FakeBrokenEverythingInput()   ' tanpa Count juga
            Dim source As MeasuredShapeSet = CopyMeasurement.MeasureSingleShape(unreadable, "uji")
            Dim pasted As MeasuredShapeSet = NewPasted(0, 0, 10, 10, 6)

            Dim result As CopyVerificationResult = CopyVerification.VerifyPasted(source, pasted, 1, 1.0)
            Check.AreEqual(CopyVerifyStatus.PartiallyVerified, result.Status, "Partial")
            Check.IsFalse(source.HasGeometry, "geometri tidak diklaim")
            Check.Contains(source.Reason, "tidak dapat dibaca", "alasan menyebut bagian yang gagal")
        End Sub)

        ' ------------------------------------------------------------------
        ' 5. Kasus koleksi kosong & Count layer tak terbaca
        ' ------------------------------------------------------------------
        suite.Add("koleksi kosong -> alasan 'selection kosong' (bukan diukur)", Sub()
            Dim empty As New FakeCollectionInput()
            Dim measured As MeasuredShapeSet = NewMeasuredSource(empty)
            Check.AreEqual("selection kosong", measured.Reason, "alasan")
            Check.IsFalse(measured.HasGeometry, "tidak ada geometri")
        End Sub)

        suite.Add("MeasureTopShapes: jumlah objek layer tak terbaca -> tidak mengarang hasil", Sub()
            Dim measured As MeasuredShapeSet = CopyMeasurement.MeasureTopShapes(New FakeBrokenCountInput(), 2)
            Check.IsFalse(measured.HasGeometry, "tidak ada geometri")
            Check.IsFalse(measured.HasStructure, "struktur ditandai tidak terbaca")
            Check.Contains(measured.Reason, "jumlah objek pada layer tidak dapat dibaca", "alasan eksplisit")
        End Sub)

        suite.Add("SafeCount tetap -1 untuk kedua jenis kegagalan (API lama kompatibel)", Sub()
            Check.AreEqual(-1, CopyMeasurement.SafeCount(New FakeBrokenCountInput()), "Count gagal dibaca")
            Check.AreEqual(-1, CopyMeasurement.SafeCount(New FakeSingleShapeInput(6, 0, 0, 1, 1)),
                           "objek tanpa Count")
            Dim twoShapes As New FakeShapeRangeInput()
            twoShapes.Add(New FakeSingleShapeInput(6, 0, 0, 1, 1))
            twoShapes.Add(New FakeSingleShapeInput(6, 5, 0, 1, 1))
            Check.AreEqual(2, CopyMeasurement.SafeCount(twoShapes), "Count terbaca")
        End Sub)

        ' ------------------------------------------------------------------
        ' 6. Semua jalur lama: copy sumber ShapeRange yang cocok -> Verified
        ' ------------------------------------------------------------------
        suite.Add("regresi: ShapeRange sumber yang cocok tetap Verified", Sub()
            Dim rangeInput As New FakeShapeRangeInput()
            rangeInput.Add(New FakeSingleShapeInput(5, 100, 100, 40, 20))

            Dim source As MeasuredShapeSet = NewMeasuredSource(rangeInput)
            Dim pasted As MeasuredShapeSet = NewPasted(100, 100, 40, 20, 5)

            Dim result As CopyVerificationResult = CopyVerification.VerifyPasted(source, pasted, 1, 1.0)
            Check.AreEqual(CopyVerifyStatus.Verified, result.Status, "masih Verified bila semua fakta ada")
        End Sub)
    End Sub
End Module
