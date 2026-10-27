Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  CopyPipelineTests.vb
'  Menjalankan URUTAN yang sama dengan CopyClickedCore untuk satu halaman:
'     ukur objek sumber  ->  "paste"  ->  ukur objek hasil  ->  VerifyPasted
'  Memakai objek palsu, jadi tidak butuh CorelDRAW, tetapi kode verifikasi yang
'  diuji adalah kode produksi yang sama (CopyMeasurement + CopyVerification).
'
'  Skenario yang diminta audit: objek tunggal, multi-selection, grup bersarang,
'  bitmap, teks, dan objek dengan ukuran campuran.
' ============================================================================

Imports System
Imports System.Collections.Generic

Public Module CopyPipelineTests

    ''' <summary>Kode shape CorelDRAW: 3 kurva, 5 bitmap, 6 teks, 7 grup.</summary>
    Private Const TYPE_CURVE As Integer = 3
    Private Const TYPE_BITMAP As Integer = 5
    Private Const TYPE_TEXT As Integer = 6
    Private Const TYPE_GROUP As Integer = 7

    ''' <summary>
    ''' Simulasi satu halaman tujuan: layer sudah berisi objek lama, lalu
    ''' "paste" menambahkan salinan objek sumber di ATAS tumpukan.
    ''' </summary>
    Private Function SimulatePage(source As List(Of FakeShape),
                                  Optional offsetX As Double = 0.0,
                                  Optional offsetY As Double = 0.0,
                                  Optional breakGeometry As Boolean = False) As FakeShapeCollection
        Dim layerItems As New List(Of FakeShape)()
        ' objek lama pada halaman tujuan (harus TIDAK ikut diukur)
        layerItems.Add(New FakeShape(TYPE_CURVE, 5, 5, 3, 3))
        layerItems.Add(New FakeShape(TYPE_TEXT, 5, 40, 3, 3))

        For Each shape As FakeShape In source
            Dim copy As New FakeShape(shape.Type, shape.CenterX + offsetX, shape.CenterY + offsetY,
                                      shape.SizeWidth, shape.SizeHeight)
            copy.FailGeometryRead = breakGeometry
            layerItems.Add(copy)
        Next
        Return New FakeShapeCollection(layerItems)
    End Function

    Private Function RunPipeline(source As List(Of FakeShape), layerShapes As FakeShapeCollection,
                                 Optional unitFactor As Double = 1.0) As CopyVerificationResult
        Dim sourceSet As MeasuredShapeSet = CopyMeasurement.MeasureSelection(New FakeShapeCollection(source))
        Dim added As Integer = layerShapes.Count - 2       ' 2 objek lama pada halaman tujuan
        Dim pastedSet As MeasuredShapeSet = CopyMeasurement.MeasurePasted(Nothing, layerShapes, added)
        Return CopyVerification.VerifyPasted(sourceSet, pastedSet, added, unitFactor)
    End Function

    Public Sub Register(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("Alur copy per halaman (Temuan A)")

        suite.Add("objek tunggal (kurva) ditempel tepat -> Verified", Sub()
            Dim source As New List(Of FakeShape)()
            source.Add(New FakeShape(TYPE_CURVE, 100, 100, 50, 30))
            Dim result As CopyVerificationResult = RunPipeline(source, SimulatePage(source))
            Check.AreEqual(CopyVerifyStatus.Verified, result.Status, "status & alasan: " & result.Reason)
        End Sub)

        suite.Add("multi-selection 3 objek -> Verified", Sub()
            Dim source As New List(Of FakeShape)()
            source.Add(New FakeShape(TYPE_TEXT, 50, 50, 40, 10))
            source.Add(New FakeShape(TYPE_BITMAP, 120, 60, 30, 30))
            source.Add(New FakeShape(TYPE_CURVE, 80, 130, 60, 20))
            Dim result As CopyVerificationResult = RunPipeline(source, SimulatePage(source))
            Check.AreEqual(CopyVerifyStatus.Verified, result.Status, "status & alasan: " & result.Reason)
        End Sub)

        suite.Add("grup bersarang (1 shape grup) -> Verified", Sub()
            Dim group As New FakeShape(TYPE_GROUP, 100, 100, 80, 60)
            group.Children.Add(New FakeShape(TYPE_CURVE, 90, 90, 10, 10))
            group.Children.Add(New FakeShape(TYPE_TEXT, 110, 110, 20, 10))
            Dim source As New List(Of FakeShape)()
            source.Add(group)
            Dim result As CopyVerificationResult = RunPipeline(source, SimulatePage(source))
            Check.AreEqual(CopyVerifyStatus.Verified, result.Status, "status & alasan: " & result.Reason)
        End Sub)

        suite.Add("bitmap + teks ukuran campuran -> Verified", Sub()
            Dim source As New List(Of FakeShape)()
            source.Add(New FakeShape(TYPE_BITMAP, 60, 60, 200, 150))
            source.Add(New FakeShape(TYPE_TEXT, 60, 200, 12, 6))
            Dim result As CopyVerificationResult = RunPipeline(source, SimulatePage(source))
            Check.AreEqual(CopyVerifyStatus.Verified, result.Status, "status & alasan: " & result.Reason)
        End Sub)

        suite.Add("hasil paste bergeser 3 mm -> Failed", Sub()
            Dim source As New List(Of FakeShape)()
            source.Add(New FakeShape(TYPE_TEXT, 100, 100, 40, 10))
            Dim result As CopyVerificationResult = RunPipeline(source, SimulatePage(source, 3, 0))
            Check.AreEqual(CopyVerifyStatus.Failed, result.Status, "status")
            Check.Contains(result.Reason, "berbeda", "alasan menyebut perbedaan")
        End Sub)

        suite.Add("hasil paste lebih kecil 2 mm -> Failed", Sub()
            Dim source As New List(Of FakeShape)()
            source.Add(New FakeShape(TYPE_BITMAP, 100, 100, 40, 40))
            Dim layer As FakeShapeCollection = SimulatePage(source)
            layer.All(2).SizeWidth = 38
            Dim result As CopyVerificationResult = RunPipeline(source, layer)
            Check.AreEqual(CopyVerifyStatus.Failed, result.Status, "status")
        End Sub)

        suite.Add("objek hasil paste tidak terukur -> Partial (bukan Verified)", Sub()
            Dim source As New List(Of FakeShape)()
            source.Add(New FakeShape(TYPE_TEXT, 100, 100, 40, 10))
            Dim result As CopyVerificationResult = RunPipeline(source, SimulatePage(source, 0, 0, True))
            Check.AreEqual(CopyVerifyStatus.PartiallyVerified, result.Status, "status")
            Check.IsFalse(result.IsVerified, "tidak boleh Verified")
        End Sub)

        suite.Add("paste tidak menambah objek -> Failed", Sub()
            Dim source As New List(Of FakeShape)()
            source.Add(New FakeShape(TYPE_TEXT, 100, 100, 40, 10))
            Dim layer As New List(Of FakeShape)()
            layer.Add(New FakeShape(TYPE_CURVE, 5, 5, 3, 3))
            Dim layerCollection As New FakeShapeCollection(layer)
            Dim sourceSet As MeasuredShapeSet = CopyMeasurement.MeasureSelection(New FakeShapeCollection(source))
            Dim pastedSet As MeasuredShapeSet = CopyMeasurement.MeasurePasted(Nothing, layerCollection, 0)
            Dim result As CopyVerificationResult = CopyVerification.VerifyPasted(sourceSet, pastedSet, 0, 1.0)
            Check.AreEqual(CopyVerifyStatus.Failed, result.Status, "status")
        End Sub)

        suite.Add("dokumen inch: geometri dikonversi eksplisit -> Verified", Sub()
            ' Sumber & hasil sama-sama dalam inch (dokumen tidak bisa di-set ke mm),
            ' faktor 25.4 dipakai; selisih 0,01 inch (0,254 mm) masih dalam toleransi.
            Dim source As New List(Of FakeShape)()
            source.Add(New FakeShape(TYPE_TEXT, 4, 4, 1, 0.5))
            Dim result As CopyVerificationResult = RunPipeline(source, SimulatePage(source, 0.01, 0), 25.4)
            Check.AreEqual(CopyVerifyStatus.Verified, result.Status, "status & alasan: " & result.Reason)
        End Sub)

        suite.Add("dokumen inch: selisih 0,05 inch (1,27 mm) -> Failed", Sub()
            Dim source As New List(Of FakeShape)()
            source.Add(New FakeShape(TYPE_TEXT, 4, 4, 1, 0.5))
            Dim result As CopyVerificationResult = RunPipeline(source, SimulatePage(source, 0.05, 0), 25.4)
            Check.AreEqual(CopyVerifyStatus.Failed, result.Status, "status di luar toleransi mm")
        End Sub)

        suite.Add("struktur berubah (2 objek jadi 1) -> Partial", Sub()
            Dim source As New List(Of FakeShape)()
            source.Add(New FakeShape(TYPE_TEXT, 100, 100, 40, 10))
            source.Add(New FakeShape(TYPE_BITMAP, 120, 100, 40, 10))
            Dim layerItems As New List(Of FakeShape)()
            layerItems.Add(New FakeShape(TYPE_CURVE, 5, 5, 3, 3))
            ' hasil paste berupa satu grup dengan kotak sama
            layerItems.Add(New FakeShape(TYPE_GROUP, 110, 100, 60, 10))
            Dim sourceSet As MeasuredShapeSet = CopyMeasurement.MeasureSelection(New FakeShapeCollection(source))
            Dim pastedSet As MeasuredShapeSet = CopyMeasurement.MeasurePasted(Nothing, New FakeShapeCollection(layerItems), 1)
            Dim result As CopyVerificationResult = CopyVerification.VerifyPasted(sourceSet, pastedSet, 1, 1.0)
            Check.AreEqual(CopyVerifyStatus.PartiallyVerified, result.Status, "status & alasan: " & result.Reason)
        End Sub)
    End Sub
End Module
