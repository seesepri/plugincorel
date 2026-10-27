Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  CopyMeasurementTests.vb
'  Regression test inti Temuan A: objek hasil paste harus diukur dari
'  z-order ATAS (bukan shape terbawah pada layer), dan objek sumber tidak
'  boleh dibandingkan dengan objek yang salah.
' ============================================================================

Imports System
Imports System.Collections.Generic

Public Module CopyMeasurementTests

    ''' <summary>Buat koleksi shape dengan urutan z-order: index 1 = paling bawah.</summary>
    Private Function BuildLayerShapes() As FakeShapeCollection
        Dim items As New List(Of FakeShape)()
        ' 3 shape lama di bawah
        items.Add(New FakeShape(3, 10, 10, 2, 2))
        items.Add(New FakeShape(3, 20, 20, 2, 2))
        items.Add(New FakeShape(3, 30, 30, 2, 2))
        ' 2 shape hasil paste di atas
        items.Add(New FakeShape(6, 100, 100, 40, 20))
        items.Add(New FakeShape(5, 160, 100, 40, 20))
        Return New FakeShapeCollection(items)
    End Function

    Public Sub Register(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("Pengukuran objek (Temuan A)")

        suite.Add("MeasureTopShapes mengukur objek TERATAS, bukan terbawah", Sub()
            Dim shapes As FakeShapeCollection = BuildLayerShapes()
            Dim measured As MeasuredShapeSet = CopyMeasurement.MeasureTopShapes(shapes, 2)

            Check.IsTrue(measured.HasGeometry, "geometri terukur")
            ' Gabungan dua objek teratas: (100,100,40x20) dan (160,100,40x20)
            ' -> x 80..180 (pusat 130), y 90..110. Objek terbawah ada di x<=31.
            Check.Near(130, measured.Bounds.CenterX, 0.0001, "centerX gabungan")
            Check.Near(100, measured.Bounds.CenterY, 0.0001, "centerY gabungan")
            Check.Near(100, measured.Bounds.Width, 0.0001, "lebar gabungan")
            Check.Near(20, measured.Bounds.Height, 0.0001, "tinggi gabungan")
            Check.AreEqual(2, measured.Signature.Count, "dua objek diukur")
        End Sub)

        suite.Add("pengukuran lama (index 1..n) akan salah — regresi dijaga", Sub()
            ' Meniru cara lama: 2 shape pertama (terbawah) -> geometri sama sekali lain.
            Dim shapes As FakeShapeCollection = BuildLayerShapes()
            Dim oldWay As MeasuredShapeSet = CopyMeasurement.MeasureSlice(shapes, 1, 2, "cara lama")
            Dim newWay As MeasuredShapeSet = CopyMeasurement.MeasureTopShapes(shapes, 2)
            Check.IsTrue(oldWay.Bounds.CenterX < 30, "cara lama mengukur area bawah (pusat x < 30)")
            Check.Near(130, newWay.Bounds.CenterX, 0.0001, "cara baru mengukur area hasil paste")
        End Sub)

        suite.Add("sidik jari struktur membedakan teks dan bitmap", Sub()
            Dim shapes As FakeShapeCollection = BuildLayerShapes()
            Dim measured As MeasuredShapeSet = CopyMeasurement.MeasureTopShapes(shapes, 2)
            Dim source As ShapeSignature = ShapeSignature.FromTypeCodes(New Integer() {6, 5})
            Check.IsTrue(measured.Signature.Matches(source), "tipe [teks, bitmap] cocok")
            Check.IsFalse(measured.Signature.Matches(ShapeSignature.FromTypeCodes(New Integer() {6, 6})),
                          "jumlah/tipe berbeda terdeteksi")
        End Sub)

        suite.Add("koleksi lebih pendek dari jumlah objek baru -> tanpa geometri", Sub()
            Dim shapes As FakeShapeCollection = BuildLayerShapes()
            Dim measured As MeasuredShapeSet = CopyMeasurement.MeasureTopShapes(shapes, 50)
            Check.IsFalse(measured.HasGeometry, "tidak menebak geometri")
            Check.Contains(measured.Reason, "kurang dari", "alasan jelas")
        End Sub)

        suite.Add("objek yang gagal dibaca -> HasGeometry False (Partial)", Sub()
            Dim items As New List(Of FakeShape)()
            Dim broken As New FakeShape(6, 0, 0, 10, 10)
            broken.FailGeometryRead = True
            items.Add(broken)
            Dim measured As MeasuredShapeSet = CopyMeasurement.MeasureTopShapes(New FakeShapeCollection(items), 1)
            Check.IsFalse(measured.HasGeometry, "geometri tidak diklaim")
            Check.Contains(measured.Reason, "geometri tidak dapat diukur", "alasan")
        End Sub)

        suite.Add("MeasurePasted memakai z-order lebih dulu", Sub()
            Dim shapes As FakeShapeCollection = BuildLayerShapes()
            Dim measured As MeasuredShapeSet = CopyMeasurement.MeasurePasted(Nothing, shapes, 1)
            Check.IsTrue(measured.HasGeometry, "geometri terukur")
            Check.Contains(measured.Source, "z-order", "sumber pengukuran z-order")
            ' Objek teratas: pusat (160,100) ukuran 40x20
            Check.Near(160, measured.Bounds.CenterX, 0.0001, "centerX objek teratas")
        End Sub)

        suite.Add("MeasurePasted jatuh ke selection bila z-order gagal", Sub()
            ' Layer kosong + selection berisi 1 objek -> pakai selection.
            Dim emptyLayer As New FakeShapeCollection(New List(Of FakeShape)())
            Dim selectionItems As New List(Of FakeShape)()
            selectionItems.Add(New FakeShape(6, 50, 50, 10, 10))
            Dim selection As New FakeShapeCollection(selectionItems)
            Dim measured As MeasuredShapeSet = CopyMeasurement.MeasurePasted(selection, emptyLayer, 1)
            Check.IsTrue(measured.HasGeometry, "geometri dari selection")
            Check.Contains(measured.Source, "selection", "sumber pengukuran")
            Check.Near(50, measured.Bounds.CenterX, 0.0001, "centerX dari selection")
        End Sub)

        suite.Add("MeasureSelection mengukur gabungan objek sumber", Sub()
            Dim items As New List(Of FakeShape)()
            items.Add(New FakeShape(6, 0, 0, 10, 10))
            items.Add(New FakeShape(5, 20, 0, 10, 10))
            Dim measured As MeasuredShapeSet = CopyMeasurement.MeasureSelection(New FakeShapeCollection(items))
            Check.Near(10, measured.Bounds.CenterX, 0.0001, "pusat gabungan")
            Check.Near(30, measured.Bounds.Width, 0.0001, "lebar gabungan")
            Check.AreEqual(2, measured.Signature.Count, "dua objek")
        End Sub)

        suite.Add("MeasureSelection pada objek tunggal (tanpa Count)", Sub()
            Dim shape As New FakeShape(6, 5, 5, 4, 4)
            Dim measured As MeasuredShapeSet = CopyMeasurement.MeasureSelection(shape)
            Check.IsTrue(measured.HasGeometry, "geometri terukur")
            Check.Near(5, measured.Bounds.CenterX, 0.0001, "centerX")
        End Sub)

        suite.Add("Scaled mengonversi inch ke mm", Sub()
            Dim items As New List(Of FakeShape)()
            items.Add(New FakeShape(6, 1, 1, 1, 1))
            Dim measured As MeasuredShapeSet = CopyMeasurement.MeasureTopShapes(New FakeShapeCollection(items), 1)
            Dim mm As MeasuredShapeSet = measured.Scaled(25.4)
            Check.Near(25.4, mm.Bounds.Width, 0.0001, "1 inch = 25.4 mm")
            Check.Near(25.4, mm.Bounds.CenterX, 0.0001, "pusat juga dikonversi")
        End Sub)
    End Sub
End Module
