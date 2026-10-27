Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  CopyInputEvidenceTests.vb
'  Review follow-up KEDUA pada PR #2 — bukti jenis input pengukuran.
'
'  Aturan yang dijaga (butir 1 review):
'    1. Shape tunggal, ShapeRange, dan koleksi COM dibedakan.
'    2. Kegagalan membaca Count TIDAK PERNAH menjadi bukti bahwa input adalah
'       satu Shape.
'    3. "Tidak punya Count" saja juga belum cukup: bentuk objek harus terbukti
'       (kode tipe terbaca). Kalau tidak, jumlah objek tidak dapat dipastikan.
'    4. Pengukuran yang terpotong tidak boleh meninggalkan sidik jari struktur
'       yang tampak utuh.
'    5. Semua kondisi di atas berujung pada status Partial, bukan Verified.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Runtime.InteropServices

''' <summary>
''' Objek yang TIDAK punya Count dan TIDAK punya kode tipe terbaca, tetapi
''' geometrinya kebetulan terbaca. Versi yang salah akan menganggapnya "satu
''' Shape" lalu memakai geometri itu sebagai bukti -> Verified palsu.
''' </summary>
Public Class FakeNoCountNoTypeInput
    Public ReadOnly Property CenterX As Double
        Get
            Return 500
        End Get
    End Property

    Public ReadOnly Property CenterY As Double
        Get
            Return 500
        End Get
    End Property

    Public ReadOnly Property SizeWidth As Double
        Get
            Return 40
        End Get
    End Property

    Public ReadOnly Property SizeHeight As Double
        Get
            Return 20
        End Get
    End Property

    Public ReadOnly Property Type As Integer
        Get
            Throw New InvalidOperationException("simulasi: Shape.Type tidak terbaca")
        End Get
    End Property
End Class

''' <summary>
''' Koleksi yang Count-nya ADA tetapi tidak didukung (NotSupportedException).
''' Ini bukan "anggota hilang", jadi jumlah objek tetap tidak diketahui.
''' </summary>
Public Class FakeNotSupportedCountInput
    Public Property CenterX As Double = 700
    Public Property CenterY As Double = 700
    Public Property SizeWidth As Double = 10
    Public Property SizeHeight As Double = 10
    Public Property Type As Integer = 3

    Public ReadOnly Property Count As Integer
        Get
            Throw New NotSupportedException("simulasi: Count tidak didukung objek ini")
        End Get
    End Property
End Class

''' <summary>
''' Koleksi yang BERBOHONG: Count = 3, tetapi objek ketiga tidak dapat diakses.
''' Pengukuran harus berhenti dan menandai struktur tidak terbaca.
''' </summary>
Public Class FakeLyingCountInput
    Public ReadOnly Property Count As Integer
        Get
            Return 3
        End Get
    End Property

    Public Default ReadOnly Property Item(oneBasedIndex As Integer) As FakeSingleShapeInput
        Get
            If oneBasedIndex = 3 Then
                Throw New InvalidOperationException("simulasi: objek ke-3 tidak dapat diakses")
            End If
            Return New FakeSingleShapeInput(6, 10 * oneBasedIndex, 10, 5, 5)
        End Get
    End Property
End Class

Public Module CopyInputEvidenceTests

    ''' <summary>Objek hasil paste palsu dengan geometri & tipe tertentu.</summary>
    Private Function NewPasted(cx As Double, cy As Double, w As Double, h As Double,
                               typeCode As Integer) As MeasuredShapeSet
        Dim items As New List(Of FakeShape)()
        items.Add(New FakeShape(typeCode, cx, cy, w, h))
        Return CopyMeasurement.MeasureTopShapes(New FakeShapeCollection(items), 1)
    End Function

    Public Sub Register(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("Bukti jenis input pengukuran (review kedua)")

        ' ------------------------------------------------------------------
        ' 1. Tanpa Count DAN tanpa bentuk yang terbukti -> Unknown
        ' ------------------------------------------------------------------
        suite.Add("tanpa Count & tipe tak terbaca: BUKAN Shape tunggal, BUKAN Verified", Sub()
            Dim input As New FakeNoCountNoTypeInput()
            Dim read As CountReadResult = CopyMeasurement.TryReadCount(input)
            Check.AreEqual(CountReadStatus.MemberMissing, read.Status, "Count memang tidak ada")

            Check.IsFalse(CopyMeasurement.LooksLikeShape(input), "bentuk Shape TIDAK terbukti")
            Check.AreEqual(MeasurementInputKind.Unknown, CopyMeasurement.ClassifyInput(input),
                           "jenis input = Unknown (bukan SingleShape)")

            Dim measured As MeasuredShapeSet = CopyMeasurement.MeasureSelection(input)
            Check.IsFalse(measured.HasGeometry, "geometri yang kebetulan terbaca TIDAK dipakai")
            Check.IsFalse(measured.HasStructure, "struktur ditandai tidak terbaca")
            Check.Contains(measured.Reason, "jumlah objek seleksi tidak dapat diverifikasi", "alasan eksplisit")
            Check.Contains(measured.Reason, "tidak dapat dipastikan", "alasan menyebut bentuk objek tak terbukti")

            ' Geometri hasil paste dibuat IDENTIK dengan geometri yang "kebetulan"
            ' terbaca dari input; tetap tidak boleh Verified.
            Dim pasted As MeasuredShapeSet = NewPasted(500, 500, 40, 20, 3)
            Dim result As CopyVerificationResult = CopyVerification.VerifyPasted(measured, pasted, 1, 1.0)
            Check.AreEqual(CopyVerifyStatus.PartiallyVerified, result.Status, "Partial, bukan Verified")
        End Sub)

        suite.Add("tanpa Count & geometri tak terbaca: tetap Partial (tidak ada klaim)", Sub()
            Dim measured As MeasuredShapeSet = CopyMeasurement.MeasureSelection(New FakeBrokenEverythingInput())
            Check.AreEqual(MeasurementInputKind.Unknown, measured.InputKind, "jenis input Unknown")
            Check.IsFalse(measured.HasGeometry, "tanpa geometri")
            Dim result As CopyVerificationResult =
                CopyVerification.VerifyPasted(measured, NewPasted(0, 0, 10, 10, 6), 1, 1.0)
            Check.AreEqual(CopyVerifyStatus.PartiallyVerified, result.Status, "Partial")
        End Sub)

        ' ------------------------------------------------------------------
        ' 2. Count ada tetapi tidak didukung -> Failed (bukan MemberMissing)
        ' ------------------------------------------------------------------
        suite.Add("Count NotSupportedException: bukan 'anggota hilang' -> Partial", Sub()
            Dim input As New FakeNotSupportedCountInput()
            Dim read As CountReadResult = CopyMeasurement.TryReadCount(input)
            Check.AreEqual(CountReadStatus.Failed, read.Status, "status Failed")
            Check.IsFalse(read.IsMemberMissing, "NotSupportedException bukan anggota hilang")
            Check.AreEqual(MeasurementInputKind.Unknown, CopyMeasurement.ClassifyInput(input), "jenis Unknown")

            Dim measured As MeasuredShapeSet = CopyMeasurement.MeasureSelection(input)
            Check.IsFalse(measured.HasGeometry, "tidak mengukur apa pun")
            Dim result As CopyVerificationResult =
                CopyVerification.VerifyPasted(measured, NewPasted(700, 700, 10, 10, 3), 1, 1.0)
            Check.AreEqual(CopyVerifyStatus.PartiallyVerified, result.Status, "Partial, bukan Verified")
        End Sub)

        ' ------------------------------------------------------------------
        ' 3. Shape tunggal sah tetap diperlakukan sebagai Shape tunggal
        ' ------------------------------------------------------------------
        suite.Add("Shape tunggal sah (tipe terbaca) tetap SingleShape", Sub()
            Dim shape As New FakeSingleShapeInput(6, 5, 5, 4, 4)
            Check.IsTrue(CopyMeasurement.LooksLikeShape(shape), "bentuk Shape terbukti")
            Check.AreEqual(MeasurementInputKind.SingleShape, CopyMeasurement.ClassifyInput(shape),
                           "jenis input SingleShape")
            Dim measured As MeasuredShapeSet = CopyMeasurement.MeasureSelection(shape)
            Check.IsTrue(measured.HasGeometry, "geometri terukur")
            Dim result As CopyVerificationResult =
                CopyVerification.VerifyPasted(measured, NewPasted(5, 5, 4, 4, 6), 1, 1.0)
            Check.AreEqual(CopyVerifyStatus.Verified, result.Status, "jalur normal tetap Verified")
        End Sub)

        ' ------------------------------------------------------------------
        ' 4. Pengukuran terpotong: struktur tidak boleh tampak utuh
        ' ------------------------------------------------------------------
        suite.Add("koleksi berbohong (Count=3, objek ke-3 hilang): struktur ditandai tidak terbaca", Sub()
            Dim input As New FakeLyingCountInput()
            Check.AreEqual(CountReadStatus.Ok, CopyMeasurement.TryReadCount(input).Status, "Count terbaca = 3")

            Dim measured As MeasuredShapeSet = CopyMeasurement.MeasureSelection(input)
            Check.IsFalse(measured.HasGeometry, "geometri tidak diklaim")
            Check.IsFalse(measured.HasStructure, "sidik jari struktur TIDAK dianggap bisa dibandingkan")
            ' Dua objek pertama memang terbaca, tetapi karena pengukuran
            ' terpotong, sidik jari itu TIDAK boleh dianggap lengkap.
            Check.AreEqual(2, measured.Signature.Count, "hanya objek yang benar-benar terbaca dicatat")
            Check.IsFalse(measured.Signature.IsComparable,
                          "sidik jari terpotong tidak boleh dibandingkan sebagai utuh")
            Check.Contains(measured.Reason, "geometri tidak dapat diukur", "alasan teknis")

            ' Meskipun hasil paste kebetulan terbaca lengkap, status tidak boleh Verified.
            Dim pasted As MeasuredShapeSet = NewPasted(10, 10, 5, 5, 6)
            Dim result As CopyVerificationResult = CopyVerification.VerifyPasted(measured, pasted, 3, 1.0)
            Check.IsFalse(result.IsVerified, "bukan Verified")
        End Sub)

        ' ------------------------------------------------------------------
        ' 5. MeasurePasted: fallback selection hanya dengan Count yang TERBACA
        ' ------------------------------------------------------------------
        suite.Add("MeasurePasted: Count seleksi gagal -> fallback TIDAK dipakai, Partial", Sub()
            Dim emptyLayer As New FakeShapeCollection(New List(Of FakeShape)())
            Dim brokenSelection As New FakeBrokenCountInput()   ' geometri terbaca, Count gagal

            Dim measured As MeasuredShapeSet = CopyMeasurement.MeasurePasted(brokenSelection, emptyLayer, 1)
            Check.IsFalse(measured.HasGeometry, "fallback tidak dipakai")
            Check.Contains(measured.Reason, "fallback selection tidak dipakai", "alasan transparan")

            Dim result As CopyVerificationResult =
                CopyVerification.VerifyPasted(NewPasted(500, 500, 40, 20, 3), measured, 1, 1.0)
            Check.AreEqual(CopyVerifyStatus.PartiallyVerified, result.Status, "Partial")
        End Sub)

        suite.Add("MeasurePasted: Count seleksi salah jumlah -> fallback tidak dipakai", Sub()
            Dim emptyLayer As New FakeShapeCollection(New List(Of FakeShape)())
            Dim selectionItems As New List(Of FakeShape)()
            selectionItems.Add(New FakeShape(6, 50, 50, 10, 10))
            selectionItems.Add(New FakeShape(6, 70, 50, 10, 10))
            Dim selection As New FakeShapeCollection(selectionItems)

            ' addedCount = 1 tetapi seleksi berisi 2 objek -> bukan bukti hasil paste.
            Dim measured As MeasuredShapeSet = CopyMeasurement.MeasurePasted(selection, emptyLayer, 1)
            Check.IsFalse(measured.HasGeometry, "seleksi dengan jumlah berbeda tidak dipakai")
        End Sub)

        suite.Add("MeasurePasted: Count seleksi cocok -> fallback dipakai (perilaku lama dijaga)", Sub()
            Dim emptyLayer As New FakeShapeCollection(New List(Of FakeShape)())
            Dim selectionItems As New List(Of FakeShape)()
            selectionItems.Add(New FakeShape(6, 50, 50, 10, 10))
            Dim measured As MeasuredShapeSet =
                CopyMeasurement.MeasurePasted(New FakeShapeCollection(selectionItems), emptyLayer, 1)
            Check.IsTrue(measured.HasGeometry, "geometri dari selection")
            Check.Near(50, measured.Bounds.CenterX, 0.0001, "pusat dari selection")
        End Sub)

        ' ------------------------------------------------------------------
        ' 6. Kompatibilitas API lama
        ' ------------------------------------------------------------------
        suite.Add("SafeCount tetap ada & tetap -1 untuk kegagalan (kompatibilitas)", Sub()
            Check.AreEqual(-1, CopyMeasurement.SafeCount(New FakeNotSupportedCountInput()), "NotSupported")
            Check.AreEqual(-1, CopyMeasurement.SafeCount(New FakeBrokenCountInput()), "COM gagal")
            Check.AreEqual(-1, CopyMeasurement.SafeCount(New FakeSingleShapeInput(6, 0, 0, 1, 1)), "tanpa Count")
            Dim collection As New FakeCollectionInput()
            collection.Add(New FakeSingleShapeInput(3, 0, 0, 1, 1))
            Check.AreEqual(1, CopyMeasurement.SafeCount(collection), "Count terbaca")
        End Sub)
    End Sub
End Module
