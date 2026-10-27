Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  ShapeScannerTests.vb
'  Regression test untuk ShapeScanner (Temuan Auto Select #2, #5, #6, #7, #8, #10).
'
'  Dataset uji mengikuti struktur dokumen yang didokumentasikan di
'  PANDUAN-UJI.md §11 (dataset "AS-1" .. "AS-6") supaya angka yang diharapkan
'  di UI sama dengan angka di test.
' ============================================================================

Imports System
Imports System.Collections.Generic

Public Module ShapeScannerTests

    ''' <summary>Bangun dataset campuran: teks, bitmap, kurva, vektor, watermark, grup, PowerClip.</summary>
    Private Function MixedDataset() As List(Of IShapeNode)
        Dim nodes As New List(Of IShapeNode)()

        ' 1. teks artistik
        nodes.Add(FakeShapeNode.Text(ShapeKind.TextArtistic, "Judul"))
        ' 2. teks paragraf
        nodes.Add(FakeShapeNode.Text(ShapeKind.TextParagraph, "Isi"))
        ' 3. bitmap level atas
        nodes.Add(FakeShapeNode.Bitmap("Foto"))
        ' 4. kurva terbuka
        nodes.Add(FakeShapeNode.Curve(10, 5, False, "Garis terbuka"))
        ' 5. kurva tertutup
        nodes.Add(FakeShapeNode.Curve(20, 5, True, "Kurva tertutup"))
        ' 6. rectangle
        nodes.Add(FakeShapeNode.Rectangle("Kotak"))
        ' 7. ellipse
        nodes.Add(FakeShapeNode.Ellipse("Bulat"))
        ' 8. polygon
        nodes.Add(FakeShapeNode.Polygon("Segi"))
        ' 9. garis bantu (tidak boleh masuk vektor)
        nodes.Add(FakeShapeNode.Guideline())
        ' 10. objek OLE (tidak boleh masuk vektor)
        nodes.Add(FakeShapeNode.OleObject())
        ' 11. watermark (tidak boleh masuk kategori apa pun)
        nodes.Add(FakeShapeNode.WatermarkShape())
        ' 12. bitmap di layer WATERMARK
        nodes.Add(FakeShapeNode.WatermarkLayerBitmap())
        ' 13. grup berisi 2 bitmap + 1 teks
        nodes.Add(FakeShapeNode.Group(New List(Of IShapeNode)() From {
            FakeShapeNode.Bitmap("Foto dalam grup 1"),
            FakeShapeNode.Bitmap("Foto dalam grup 2"),
            FakeShapeNode.Text(ShapeKind.TextParagraph, "Teks dalam grup")
        }, "Grup isi"))
        ' 14. PowerClip berisi 1 bitmap
        nodes.Add(FakeShapeNode.PowerClip(New List(Of IShapeNode)() From {
            FakeShapeNode.Bitmap("Foto dalam PowerClip")
        }, "Container"))
        ' 15. grup bersarang 2 level berisi kurva terbuka
        nodes.Add(FakeShapeNode.Group(New List(Of IShapeNode)() From {
            FakeShapeNode.Group(New List(Of IShapeNode)() From {
                FakeShapeNode.Curve(30, 4, False, "Kurva bersarang")
            }, "Grup dalam")
        }, "Grup luar"))
        ' 16. objek yang tipenya tak terbaca
        nodes.Add(FakeShapeNode.UnreadableShape())

        Return nodes
    End Function

    Private Function Scan(category As ShapeCategory) As ShapeScanResult
        Return ShapeScanner.Scan(MixedDataset(), 1, ShapeScanOptions.ForCategory(category))
    End Function

    Public Sub Register(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("Auto Select — pemindaian objek")

        ' --- Teks (AS-1) -----------------------------------------------------
        suite.Add("teks: 2 di level atas + 1 di dalam grup", Sub()
            Dim result As ShapeScanResult = Scan(ShapeCategory.Text)
            Check.AreEqual(3, result.MatchCount, "jumlah teks")
        End Sub)

        ' --- Bitmap (AS-2) ---------------------------------------------------
        suite.Add("bitmap: level atas + dalam grup + dalam PowerClip", Sub()
            Dim result As ShapeScanResult = Scan(ShapeCategory.Bitmap)
            ' Foto, 2 bitmap dalam grup, 1 bitmap dalam PowerClip = 4.
            Check.AreEqual(4, result.MatchCount, "jumlah bitmap")
            Dim insidePowerClip As Integer = 0
            For Each hit As ShapeHit In result.Hits
                If hit.InsidePowerClip Then insidePowerClip += 1
            Next
            Check.AreEqual(1, insidePowerClip, "satu bitmap terdeteksi di dalam PowerClip")
        End Sub)

        suite.Add("bitmap di dalam PowerClip memakai CONTAINER sebagai target seleksi", Sub()
            Dim result As ShapeScanResult = Scan(ShapeCategory.Bitmap)
            For Each hit As ShapeHit In result.Hits
                If hit.InsidePowerClip Then
                    Check.IsTrue(TypeOf hit.SelectionTarget Is FakeShapeNode, "target bertipe node")
                    Dim targetNode As FakeShapeNode = DirectCast(hit.SelectionTarget, FakeShapeNode)
                    Check.AreEqual("Container", targetNode.Facts.ShapeName, "target = container PowerClip")
                    Return
                End If
            Next
            Check.Fail("tidak ada hit di dalam PowerClip")
        End Sub)

        ' --- Kurva terbuka (AS-3) -------------------------------------------
        suite.Add("kurva terbuka: level atas + bersarang, kurva tertutup dikecualikan", Sub()
            Dim result As ShapeScanResult = Scan(ShapeCategory.OpenCurve)
            Check.AreEqual(2, result.MatchCount, "2 kurva terbuka")
        End Sub)

        ' --- Vektor (AS-4) ---------------------------------------------------
        suite.Add("vektor: rectangle + ellipse + polygon + 2 kurva (guideline/OLE dikecualikan)", Sub()
            Dim result As ShapeScanResult = Scan(ShapeCategory.Vector)
            ' rectangle, ellipse, polygon, kurva terbuka, kurva tertutup, kurva bersarang = 6
            Check.AreEqual(6, result.MatchCount, "jumlah objek vektor")
            For Each hit As ShapeHit In result.Hits
                Check.DoesNotContain(hit.KindLabel, "guideline", "garis bantu tidak boleh masuk")
                Check.DoesNotContain(hit.KindLabel, "OLE", "objek OLE tidak boleh masuk")
            Next
        End Sub)

        ' --- Semua objek (AS-5) ---------------------------------------------
        suite.Add("semua objek: grup utuh, watermark & grup ber-watermark dilewati", Sub()
            Dim result As ShapeScanResult = Scan(ShapeCategory.AllObjects)
            ' Dataset: 16 node akar.
            '   - 2 watermark (WM_TEXT + bitmap layer WATERMARK) dilewati  -> -2
            '   - 1 objek yang tipenya tak terbaca tidak ditebak            -> -1
            '   = 13 objek. Tidak ada grup ber-watermark di dataset ini.
            Check.AreEqual(13, result.MatchCount, "jumlah objek level atas")
            Check.AreEqual(1, result.UnreadableShapes, "objek tak terbaca dilaporkan terpisah")
            For Each hit As ShapeHit In result.Hits
                Check.DoesNotContain(hit.ShapeName, "WM_", "watermark tidak boleh terpilih")
            Next
            Dim groupHits As Integer = 0
            For Each hit As ShapeHit In result.Hits
                If hit.KindLabel = "Grup" Then groupHits += 1
            Next
            Check.AreEqual(3, groupHits, "3 grup dipilih sebagai satu kesatuan")
        End Sub)

        suite.Add("semua objek: grup yang memuat watermark dilewati seluruhnya", Sub()
            Dim nodes As New List(Of IShapeNode)()
            nodes.Add(FakeShapeNode.Group(New List(Of IShapeNode)() From {
                FakeShapeNode.Bitmap("Foto sah"),
                FakeShapeNode.WatermarkShape()
            }, "Grup campur"))
            nodes.Add(FakeShapeNode.Bitmap("Foto biasa"))

            Dim result As ShapeScanResult = ShapeScanner.Scan(nodes, 1, ShapeScanOptions.ForCategory(ShapeCategory.AllObjects))
            Check.AreEqual(1, result.MatchCount, "hanya objek non-grup yang dipilih")
            Check.AreEqual(1, result.WatermarksSkipped, "grup ber-watermark dihitung dilewati")
        End Sub)

        suite.Add("kategori Bitmap: watermark di dalam grup dilewati, saudaranya tetap terpilih", Sub()
            Dim nodes As New List(Of IShapeNode)()
            nodes.Add(FakeShapeNode.Group(New List(Of IShapeNode)() From {
                FakeShapeNode.Bitmap("Foto sah"),
                FakeShapeNode.WatermarkLayerBitmap()
            }, "Grup campur"))

            Dim result As ShapeScanResult = ShapeScanner.Scan(nodes, 1, ShapeScanOptions.ForCategory(ShapeCategory.Bitmap))
            Check.AreEqual(1, result.MatchCount, "hanya bitmap sah yang terpilih")
            Check.AreEqual(1, result.WatermarksSkipped, "watermark dilewati")
        End Sub)

        ' --- Watermark global -------------------------------------------------
        suite.Add("watermark tidak masuk kategori apa pun", Sub()
            For Each category As ShapeCategory In New ShapeCategory() {
                    ShapeCategory.Text, ShapeCategory.Bitmap, ShapeCategory.OpenCurve,
                    ShapeCategory.Vector, ShapeCategory.AllObjects}
                Dim result As ShapeScanResult = Scan(category)
                For Each hit As ShapeHit In result.Hits
                    Check.DoesNotContain(hit.ShapeName, "WM_", "watermark pada " & category.ToString())
                Next
                Check.IsTrue(result.WatermarksSkipped >= 2, "dua watermark terdeteksi pada " & category.ToString())
            Next
        End Sub)

        ' --- Grup bersarang & kedalaman --------------------------------------
        suite.Add("grup bersarang 2 level: objek terdalam ditemukan", Sub()
            Dim nodes As New List(Of IShapeNode)()
            nodes.Add(FakeShapeNode.Group(New List(Of IShapeNode)() From {
                FakeShapeNode.Group(New List(Of IShapeNode)() From {
                    FakeShapeNode.Group(New List(Of IShapeNode)() From {
                        FakeShapeNode.Bitmap("Bitmap paling dalam")
                    }, "L3")
                }, "L2")
            }, "L1"))

            Dim result As ShapeScanResult = ShapeScanner.Scan(nodes, 1, ShapeScanOptions.ForCategory(ShapeCategory.Bitmap))
            Check.AreEqual(1, result.MatchCount, "ditemukan di kedalaman 3")
            Check.AreEqual(3, result.Hits(0).Depth, "kedalaman tercatat")
        End Sub)

        suite.Add("batas kedalaman: pemotongan dicatat, bukan hilang diam-diam", Sub()
            Dim nodes As New List(Of IShapeNode)()
            Dim deepest As IShapeNode = FakeShapeNode.Bitmap("Terlalu dalam")
            For level As Integer = 1 To 6
                deepest = FakeShapeNode.Group(New List(Of IShapeNode)() From {deepest}, "G" & level.ToString())
            Next
            nodes.Add(deepest)

            Dim options As ShapeScanOptions = ShapeScanOptions.ForCategory(ShapeCategory.Bitmap)
            options.MaxDepth = 2
            Dim result As ShapeScanResult = ShapeScanner.Scan(nodes, 1, options)

            Check.AreEqual(0, result.MatchCount, "tidak ada hasil di luar batas")
            Check.IsTrue(result.DepthTruncated > 0, "pemotongan tercatat")
            Check.Contains(String.Join(" ", result.Diagnostics.ToArray()), "melebihi batas", "diagnostik menyebut batas")
        End Sub)

        suite.Add("grup yang isinya tidak terbaca: dicatat sebagai kegagalan", Sub()
            Dim nodes As New List(Of IShapeNode)()
            nodes.Add(FakeShapeNode.GroupWithUnreadableChildren())

            Dim result As ShapeScanResult = ShapeScanner.Scan(nodes, 1, ShapeScanOptions.ForCategory(ShapeCategory.Bitmap))
            Check.AreEqual(0, result.MatchCount, "tidak ada hasil")
            Check.IsTrue(result.ChildrenFailures > 0, "kegagalan baca dicatat")
            Check.Contains(String.Join(" ", result.Diagnostics.ToArray()), "tidak dapat dibaca", "diagnostik")
        End Sub)

        suite.Add("fakta objek melempar exception: dicatat, pemindaian lanjut", Sub()
            Dim nodes As New List(Of IShapeNode)()
            Dim broken As New FakeShapeNode()
            broken.ThrowOnGetFacts = True
            nodes.Add(broken)
            nodes.Add(FakeShapeNode.Bitmap("Foto sah"))

            Dim result As ShapeScanResult = ShapeScanner.Scan(nodes, 1, ShapeScanOptions.ForCategory(ShapeCategory.Bitmap))
            Check.AreEqual(1, result.MatchCount, "objek lain tetap dipindai")
            Check.AreEqual(1, result.UnreadableShapes, "kegagalan dicatat")
        End Sub)

        ' --- PowerClip ---------------------------------------------------------
        suite.Add("PowerClip: isi tidak terbaca -> dilaporkan (jumlah bisa kurang)", Sub()
            Dim nodes As New List(Of IShapeNode)()
            nodes.Add(FakeShapeNode.PowerClipWithUnreadableContents())

            Dim result As ShapeScanResult = ShapeScanner.Scan(nodes, 1, ShapeScanOptions.ForCategory(ShapeCategory.Bitmap))
            Check.AreEqual(0, result.MatchCount, "tidak ada hasil")
            Check.AreEqual(1, result.PowerClipsScanned, "container terdeteksi")
            Check.AreEqual(1, result.PowerClipFailures, "kegagalan dicatat")
            Check.Contains(String.Join(" ", result.Diagnostics.ToArray()), "TIDAK dipindai", "peringatan eksplisit")
        End Sub)

        suite.Add("PowerClip: beberapa isi cocok -> target di-dedupe", Sub()
            Dim nodes As New List(Of IShapeNode)()
            nodes.Add(FakeShapeNode.PowerClip(New List(Of IShapeNode)() From {
                FakeShapeNode.Bitmap("Foto A"),
                FakeShapeNode.Bitmap("Foto B"),
                FakeShapeNode.Bitmap("Foto C")
            }, "Container"))

            Dim result As ShapeScanResult = ShapeScanner.Scan(nodes, 1, ShapeScanOptions.ForCategory(ShapeCategory.Bitmap))
            Check.AreEqual(3, result.MatchCount, "3 objek cocok")
            Check.AreEqual(1, result.DistinctTargetCount, "hanya 1 target seleksi (container)")
        End Sub)

        suite.Add("PowerClip tidak dipindai bila opsi dimatikan (dicatat)", Sub()
            Dim nodes As New List(Of IShapeNode)()
            nodes.Add(FakeShapeNode.PowerClip(New List(Of IShapeNode)() From {
                FakeShapeNode.Bitmap("Foto dalam")
            }))

            Dim options As ShapeScanOptions = ShapeScanOptions.ForCategory(ShapeCategory.Bitmap)
            options.ScanPowerClipContents = False
            Dim result As ShapeScanResult = ShapeScanner.Scan(nodes, 1, options)
            Check.AreEqual(0, result.MatchCount, "isi PowerClip dilewati")
            Check.AreEqual(0, result.PowerClipsScanned, "tidak ada container dipindai")
        End Sub)

        ' --- Layer terkunci ----------------------------------------------------
        suite.Add("layer terkunci: objek tidak diklaim terpilih + dicatat", Sub()
            Dim nodes As New List(Of IShapeNode)()
            nodes.Add(FakeShapeNode.LockedLayerRectangle())
            nodes.Add(FakeShapeNode.Rectangle("Kotak bebas"))

            Dim result As ShapeScanResult = ShapeScanner.Scan(nodes, 1, ShapeScanOptions.ForCategory(ShapeCategory.Vector))
            Check.AreEqual(2, result.MatchCount, "dua objek cocok")
            Check.AreEqual(1, result.DistinctTargetCount, "hanya satu yang bisa dipilih")
            Check.AreEqual(1, result.LockedLayerSkipped, "objek terkunci dicatat")
        End Sub)

        ' --- Kategori kosong ---------------------------------------------------
        suite.Add("tidak ada hasil: target kosong, tidak ada operasi destruktif", Sub()
            Dim nodes As New List(Of IShapeNode)()
            nodes.Add(FakeShapeNode.Rectangle("Kotak"))
            Dim result As ShapeScanResult = ShapeScanner.Scan(nodes, 1, ShapeScanOptions.ForCategory(ShapeCategory.Text))
            Check.AreEqual(0, result.MatchCount, "tidak ada teks")
            Check.AreEqual(0, result.DistinctTargetCount, "tidak ada target")
            Check.AreEqual("(tidak ada objek cocok)", result.DescribeHits(), "laporan jujur")
        End Sub)

        suite.Add("daftar objek null: diagnostik, bukan exception", Sub()
            Dim result As ShapeScanResult = ShapeScanner.Scan(Nothing, 1, ShapeScanOptions.ForCategory(ShapeCategory.Text))
            Check.AreEqual(0, result.MatchCount, "tidak ada hasil")
            Check.IsTrue(result.Diagnostics.Count > 0, "diagnostik ada")
        End Sub)

        ' --- Statistik ---------------------------------------------------------
        suite.Add("statistik pemindaian tercatat (scanned, grup, watermark)", Sub()
            Dim result As ShapeScanResult = Scan(ShapeCategory.Bitmap)
            Check.IsTrue(result.ScannedShapes > 0, "jumlah objek dipindai")
            Check.IsTrue(result.GroupsVisited >= 4, "grup dikunjungi: " & result.GroupsVisited)
            Check.IsTrue(result.WatermarksSkipped >= 2, "watermark dilewati")
            Check.IsTrue(result.UnreadableShapes >= 1, "objek tak terbaca dicatat")
        End Sub)
    End Sub
End Module
