Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  AutoSelectClassifierTests.vb
'  Regression test untuk ShapeClassifier (Temuan Auto Select #1–#4, #8).
' ============================================================================

Imports System

Public Module AutoSelectClassifierTests

    Private Function Facts(typeCode As Integer, Optional name As String = "", Optional layer As String = "") As ShapeFacts
        Return ShapeFacts.FromValues(typeCode, name, layer)
    End Function

    Public Sub Register(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("Auto Select — klasifikasi objek")

        ' --- Teks: artistic & paragraph -------------------------------------
        suite.Add("teks artistik -> kategori Teks", Sub()
            Dim facts_ As ShapeFacts = Facts(ShapeKind.TextShape)
            facts_.TextKind = ShapeKind.TextArtistic
            Dim result As ClassificationResult = ShapeClassifier.Classify(facts_, ShapeCategory.Text)
            Check.IsTrue(result.Matches, "cocok")
            Check.AreEqual("Teks artistik", result.KindLabel, "label")
        End Sub)

        suite.Add("teks paragraf -> kategori Teks", Sub()
            Dim facts_ As ShapeFacts = Facts(ShapeKind.TextShape)
            facts_.TextKind = ShapeKind.TextParagraph
            Dim result As ClassificationResult = ShapeClassifier.Classify(facts_, ShapeCategory.Text)
            Check.IsTrue(result.Matches, "cocok")
            Check.AreEqual("Teks paragraf", result.KindLabel, "label")
        End Sub)

        suite.Add("teks tanpa jenis terbaca tetap masuk Teks + catatan", Sub()
            Dim facts_ As ShapeFacts = Facts(ShapeKind.TextShape)
            facts_.TextKindError = "simulasi: Text.Type tidak terbaca"
            Dim result As ClassificationResult = ShapeClassifier.Classify(facts_, ShapeCategory.Text)
            Check.IsTrue(result.Matches, "tetap cocok (tipe shape sudah pasti teks)")
            Check.Contains(result.Note, "jenis teks tidak terbaca", "catatan keterbacaan")
        End Sub)

        suite.Add("bitmap bukan teks", Sub()
            Check.IsFalse(ShapeClassifier.Classify(Facts(ShapeKind.BitmapShape), ShapeCategory.Text).Matches, "bitmap != teks")
        End Sub)

        ' --- Bitmap ----------------------------------------------------------
        suite.Add("bitmap -> kategori Bitmap", Sub()
            Dim result As ClassificationResult = ShapeClassifier.Classify(Facts(ShapeKind.BitmapShape), ShapeCategory.Bitmap)
            Check.IsTrue(result.Matches, "cocok")
            Check.AreEqual("Bitmap", result.KindLabel, "label")
        End Sub)

        ' --- Kurva terbuka: status aktual, bukan tebakan --------------------
        suite.Add("kurva tertutup -> BUKAN kurva terbuka", Sub()
            Dim facts_ As ShapeFacts = Facts(ShapeKind.CurveShape)
            facts_.CurveClosed = True
            facts_.CurveClosedKnown = True
            Check.IsFalse(ShapeClassifier.Classify(facts_, ShapeCategory.OpenCurve).Matches, "kurva tertutup bukan kurva terbuka")
        End Sub)

        suite.Add("kurva terbuka -> kategori Kurva terbuka", Sub()
            Dim facts_ As ShapeFacts = Facts(ShapeKind.CurveShape)
            facts_.CurveClosed = False
            facts_.CurveClosedKnown = True
            Dim result As ClassificationResult = ShapeClassifier.Classify(facts_, ShapeCategory.OpenCurve)
            Check.IsTrue(result.Matches, "cocok")
            Check.AreEqual("Kurva terbuka", result.KindLabel, "label")
        End Sub)

        suite.Add("Curve.Closed tak terbaca -> tidak dianggap kurva terbuka + dilaporkan", Sub()
            Dim facts_ As ShapeFacts = Facts(ShapeKind.CurveShape)
            facts_.CurveClosedKnown = False
            facts_.CurveError = "simulasi: gagal baca"
            Dim result As ClassificationResult = ShapeClassifier.Classify(facts_, ShapeCategory.OpenCurve)
            Check.IsFalse(result.Matches, "tidak menebak")
            Check.Contains(result.Note, "Curve.Closed tidak dapat dibaca", "kegagalan dicatat")
        End Sub)

        ' --- Vektor: definisi eksplisit --------------------------------------
        suite.Add("vektor mencakup rectangle, ellipse, polygon, kurva", Sub()
            Check.IsTrue(ShapeClassifier.Classify(Facts(ShapeKind.RectangleShape), ShapeCategory.Vector).Matches, "rectangle")
            Check.IsTrue(ShapeClassifier.Classify(Facts(ShapeKind.EllipseShape), ShapeCategory.Vector).Matches, "ellipse")
            Check.IsTrue(ShapeClassifier.Classify(Facts(ShapeKind.PolygonShape), ShapeCategory.Vector).Matches, "polygon")

            Dim openCurve As ShapeFacts = Facts(ShapeKind.CurveShape)
            openCurve.CurveClosed = False
            openCurve.CurveClosedKnown = True
            Dim closedCurve As ShapeFacts = Facts(ShapeKind.CurveShape)
            closedCurve.CurveClosed = True
            closedCurve.CurveClosedKnown = True
            Check.IsTrue(ShapeClassifier.Classify(openCurve, ShapeCategory.Vector).Matches, "kurva terbuka termasuk vektor")
            Check.IsTrue(ShapeClassifier.Classify(closedCurve, ShapeCategory.Vector).Matches, "kurva tertutup termasuk vektor")
        End Sub)

        suite.Add("vektor TIDAK mencakup guideline / OLE / teks / bitmap / grup", Sub()
            Check.IsFalse(ShapeClassifier.Classify(Facts(ShapeKind.GuidelineShape), ShapeCategory.Vector).Matches, "guideline")
            Check.IsFalse(ShapeClassifier.Classify(Facts(ShapeKind.OleObjectShape), ShapeCategory.Vector).Matches, "OLE")
            Check.IsFalse(ShapeClassifier.Classify(Facts(ShapeKind.TextShape), ShapeCategory.Vector).Matches, "teks")
            Check.IsFalse(ShapeClassifier.Classify(Facts(ShapeKind.BitmapShape), ShapeCategory.Vector).Matches, "bitmap")
            Check.IsFalse(ShapeClassifier.Classify(Facts(ShapeKind.GroupShape), ShapeCategory.Vector).Matches, "grup")
            Check.IsFalse(ShapeClassifier.Classify(Facts(ShapeKind.SelectionShape), ShapeCategory.Vector).Matches, "bentuk seleksi")
        End Sub)

        suite.Add("kode tipe tak dikenal -> vektor khusus + catatan verifikasi", Sub()
            Dim result As ClassificationResult = ShapeClassifier.Classify(Facts(15), ShapeCategory.Vector)
            Check.IsTrue(result.Matches, "diperlakukan sebagai vektor khusus")
            Check.Contains(result.Note, "verifikasi manual", "catatan eksplisit")
        End Sub)

        suite.Add("kurva dengan status tertutup tak terbaca tetap vektor + catatan", Sub()
            Dim facts_ As ShapeFacts = Facts(ShapeKind.CurveShape)
            facts_.CurveClosedKnown = False
            Dim result As ClassificationResult = ShapeClassifier.Classify(facts_, ShapeCategory.Vector)
            Check.IsTrue(result.Matches, "tetap objek vektor")
            Check.Contains(result.KindLabel, "status tertutup tidak terbaca", "label jujur")
        End Sub)

        ' --- Semua objek ------------------------------------------------------
        suite.Add("Semua objek: grup dipilih sebagai satu kesatuan", Sub()
            Dim result As ClassificationResult = ShapeClassifier.Classify(Facts(ShapeKind.GroupShape), ShapeCategory.AllObjects)
            Check.IsTrue(result.Matches, "grup cocok")
            Check.AreEqual("Grup", result.KindLabel, "label grup")
        End Sub)

        ' --- Watermark --------------------------------------------------------
        suite.Add("watermark WM_ dikecualikan di semua kategori", Sub()
            Dim facts_ As ShapeFacts = Facts(ShapeKind.TextShape, "WM_TEXT")
            facts_.TextKind = ShapeKind.TextArtistic
            For Each category As ShapeCategory In New ShapeCategory() {
                    ShapeCategory.Text, ShapeCategory.Bitmap, ShapeCategory.OpenCurve,
                    ShapeCategory.Vector, ShapeCategory.AllObjects}
                Check.IsFalse(ShapeClassifier.Classify(facts_, category).Matches,
                              "watermark tidak boleh cocok pada " & category.ToString())
            Next
        End Sub)

        suite.Add("objek pada layer WATERMARK dikecualikan", Sub()
            Dim facts_ As ShapeFacts = Facts(ShapeKind.BitmapShape, "logo", "WATERMARK")
            Check.IsFalse(ShapeClassifier.Classify(facts_, ShapeCategory.Bitmap).Matches, "bitmap di layer WATERMARK")
            Check.IsFalse(ShapeClassifier.Classify(facts_, ShapeCategory.AllObjects).Matches, "semua objek")
        End Sub)

        suite.Add("layer watermark case-insensitive", Sub()
            Dim facts_ As ShapeFacts = Facts(ShapeKind.BitmapShape, "logo", "watermark")
            Check.IsFalse(ShapeClassifier.Classify(facts_, ShapeCategory.Bitmap).Matches, "huruf kecil")
        End Sub)

        ' --- Tipe tidak terbaca ------------------------------------------------
        suite.Add("tipe tidak terbaca -> tidak cocok + catatan", Sub()
            Dim facts_ As ShapeFacts = New ShapeFacts()
            facts_.TypeError = "simulasi: shape.Type gagal"
            Dim result As ClassificationResult = ShapeClassifier.Classify(facts_, ShapeCategory.AllObjects)
            Check.IsFalse(result.Matches, "tidak menebak")
            Check.Contains(result.Note, "tidak dapat dibaca", "catatan")
        End Sub)

        ' --- Layer terkunci ----------------------------------------------------
        suite.Add("objek pada layer terkunci diberi catatan", Sub()
            Dim facts_ As ShapeFacts = Facts(ShapeKind.RectangleShape, "kotak", "Layer kunci")
            facts_.LayerEditable = False
            facts_.LayerEditableKnown = True
            Dim result As ClassificationResult = ShapeClassifier.Classify(facts_, ShapeCategory.Vector)
            Check.IsTrue(result.Matches, "tetap diklasifikasi")
            Check.Contains(result.Note, "terkunci", "catatan layer terkunci")
        End Sub)

        ' --- Kontrak UI --------------------------------------------------------
        suite.Add("label kategori sesuai ComboBox lama (kompatibilitas UI)", Sub()
            Dim labels As String() = ShapeClassifier.CategoryLabels()
            Check.AreEqual(5, labels.Length, "5 kategori")
            Check.AreEqual("Teks (Artistik + Paragraf)", labels(0), "kategori 0")
            Check.AreEqual("Gambar / Bitmap", labels(1), "kategori 1")
            Check.AreEqual("Garis / Kurva terbuka", labels(2), "kategori 2")
            Check.Contains(labels(3), "Objek vektor", "kategori 3")
            Check.AreEqual("Semua objek", labels(4), "kategori 4")
        End Sub)

        suite.Add("CategoryFromIndex aman untuk indeks tidak valid", Sub()
            Check.AreEqual(ShapeCategory.Text, ShapeClassifier.CategoryFromIndex(-1), "negatif")
            Check.AreEqual(ShapeCategory.AllObjects, ShapeClassifier.CategoryFromIndex(4), "batas atas")
            Check.AreEqual(ShapeCategory.Text, ShapeClassifier.CategoryFromIndex(99), "di luar batas")
        End Sub)

        suite.Add("CategoryHelp menjelaskan definisi vektor", Sub()
            Check.Contains(ShapeClassifier.CategoryHelp(ShapeCategory.Vector), "rectangle", "menyebut rectangle")
            Check.Contains(ShapeClassifier.CategoryHelp(ShapeCategory.Vector), "Tidak termasuk", "menyebut pengecualian")
            Check.Contains(ShapeClassifier.CategoryHelp(ShapeCategory.OpenCurve), "Curve.Closed", "menyebut sumber status")
        End Sub)
    End Sub
End Module
