Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  AutoSelectServiceTests.vb
'  Regression test untuk pemindahan orkestrasi Auto Select dari CopierForm ke
'  AutoSelectService: pemindaian semua halaman, snapshot navigasi, penerapan
'  seleksi + verifikasi jumlah.
'
'  Memakai FakeCorel (halaman/layer/shape palsu) dan FakeShapeNode untuk pohon
'  objek — bukan COM.
' ============================================================================

Imports System
Imports System.Collections.Generic

Public Module AutoSelectServiceTests

    Private Function NewConnection(doc As FakeDocument, app As FakeCorelApplication) As CorelConnectionService
        app.ActiveDocumentInstance = doc
        Dim service As New CorelConnectionService(New FakeCorelPrompt(),
            Function(progId As String) DirectCast(app, Object),
            Function(progId As String) DirectCast(Nothing, Object),
            New String() {"CorelDRAW.Application.24"})
        service.Connect()
        service.RememberLoadedIdentity()
        Return service
    End Function

    Private Function NewDocument(pageCount As Integer) As FakeDocument
        Dim doc As New FakeDocument("C:\uji\auto.cdr")
        For i As Integer = 1 To pageCount
            Dim page As New FakePage(i)
            page.EnsureActiveLayer()
            doc.Pages.Add(page)
        Next
        doc.ActivePage = doc.Pages(1)
        Return doc
    End Function

    Private Function AddShape(doc As FakeDocument, pageIndex As Integer, typeCode As Integer,
                              name As String, width As Double, height As Double) As FakeShape
        Dim layer As FakeLayer = doc.Pages(pageIndex).EnsureActiveLayer()
        Dim shape As New FakeShape(typeCode, 50, 50, width, height)
        shape.Name = name
        shape.Layer = layer
        layer.Items.Add(shape)
        Return shape
    End Function

    Public Sub Register(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("AutoSelectService (arsitektur)")

        suite.Add("pemindaian semua halaman: hitungan per halaman + snapshot navigasi", Sub()
            Dim doc As FakeDocument = NewDocument(3)
            AddShape(doc, 1, 3, "Kotak 1", 20, 20)     ' rectangle -> vektor
            AddShape(doc, 2, 6, "Teks 2", 40, 8)       ' teks
            AddShape(doc, 3, 5, "Foto 3", 60, 40)      ' bitmap

            Dim svc As New AutoSelectService(NewConnection(doc, New FakeCorelApplication()), New NullProgressSink())
            Dim report As AutoSelectScanReport = svc.ScanAllPages(ShapeCategory.AllObjects)

            Check.AreEqual(3, report.PagesScanned, "tiga halaman dipindai")
            Check.AreEqual(3, report.TotalMatches, "tiga objek cocok")
            Check.AreEqual(3, report.PagesWithMatches.Count, "tiga halaman berisi hasil")
            Check.IsTrue(svc.Navigator.HasResults, "snapshot navigasi dibuat")
            Check.AreEqual(3, svc.Navigator.Count, "navigator memuat 3 halaman hasil")
            Check.AreEqual(3, svc.Scans.Count, "hasil per halaman disimpan")
        End Sub)

        suite.Add("progres pemindaian dilaporkan per halaman (teks status dipertahankan)", Sub()
            Dim doc As FakeDocument = NewDocument(2)
            AddShape(doc, 1, 3, "Kotak", 20, 20)
            Dim sink As New RecordingProgressSink()
            Dim svc As New AutoSelectService(NewConnection(doc, New FakeCorelApplication()), sink)

            svc.ScanAllPages(ShapeCategory.AllObjects)

            Check.AreEqual(2, sink.Messages.Count, "satu pesan per halaman")
            Check.Contains(sink.Messages(0), "Memindai progres: halaman 1/2", "format status")
            Check.Contains(sink.Messages(1), "halaman 2/2", "halaman kedua")
            Check.AreEqual(2, sink.PumpCount, "DoEvents dipanggil per halaman")
        End Sub)

        suite.Add("kategori spesifik hanya menghitung objek yang cocok", Sub()
            Dim doc As FakeDocument = NewDocument(2)
            AddShape(doc, 1, 6, "Teks A", 40, 8)
            AddShape(doc, 2, 3, "Kotak B", 20, 20)

            Dim svc As New AutoSelectService(NewConnection(doc, New FakeCorelApplication()), New NullProgressSink())
            Dim report As AutoSelectScanReport = svc.ScanAllPages(ShapeCategory.Text)

            Check.AreEqual(1, report.TotalMatches, "hanya satu teks")
            Check.AreEqual(1, report.PagesWithMatches(0), "halaman 1")
            Check.AreEqual(0, report.PerPage(2).MatchCount, "halaman 2 kosong")
        End Sub)

        suite.Add("tanpa hasil -> navigator tidak dimulai (tidak ada sesi palsu)", Sub()
            Dim doc As FakeDocument = NewDocument(2)
            AddShape(doc, 1, 3, "Kotak", 20, 20)

            Dim svc As New AutoSelectService(NewConnection(doc, New FakeCorelApplication()), New NullProgressSink())
            Dim report As AutoSelectScanReport = svc.ScanAllPages(ShapeCategory.Text)

            Check.IsFalse(report.HasMatches, "tidak ada hasil")
            Check.IsFalse(svc.Navigator.HasResults, "navigator kosong")
        End Sub)

        suite.Add("pemindaian satu halaman (cakupan halaman aktif)", Sub()
            Dim doc As FakeDocument = NewDocument(2)
            AddShape(doc, 1, 3, "Kotak", 20, 20)
            AddShape(doc, 2, 3, "Kotak 2", 20, 20)

            Dim svc As New AutoSelectService(NewConnection(doc, New FakeCorelApplication()), New NullProgressSink())
            Dim scan As ShapeScanResult = svc.ScanSinglePage(1, ShapeCategory.Vector)

            Check.IsNotNull(scan, "hasil ada")
            Check.AreEqual(1, scan.MatchCount, "satu objek")
            Check.IsTrue(svc.Navigator.HasResults, "navigasi disiapkan")
            Check.AreEqual(1, svc.Navigator.Count, "satu halaman hasil")
        End Sub)

        suite.Add("ApplySelection: tidak ada target -> seleksi lama TIDAK dibersihkan", Sub()
            Dim doc As FakeDocument = NewDocument(1)
            Dim existing As FakeShape = AddShape(doc, 1, 3, "Kotak lama", 20, 20)
            doc.Selection.Add(existing)

            Dim scan As New ShapeScanResult()
            scan.PageIndex = 1
            scan.Category = ShapeCategory.Text

            Dim svc As New AutoSelectService(NewConnection(doc, New FakeCorelApplication()), New NullProgressSink())
            Dim outcome As SelectionOutcome = svc.ApplySelection(scan, 1)

            Check.AreEqual(0, doc.ClearSelectionCalls, "ClearSelection tidak dipanggil")
            Check.IsFalse(outcome.WasAttempted, "tidak ada operasi")
            Check.AreEqual(1, doc.Selection.Count, "seleksi lama tetap ada")
        End Sub)

        suite.Add("ApplySelection: target dipilih & jumlah diverifikasi", Sub()
            Dim doc As FakeDocument = NewDocument(1)
            Dim target As FakeShape = AddShape(doc, 1, 3, "Kotak", 20, 20)

            Dim hit As New ShapeHit()
            hit.SelectionTarget = target
            hit.ShapeName = "Kotak"
            Dim scan As New ShapeScanResult()
            scan.PageIndex = 1
            scan.Category = ShapeCategory.Vector
            scan.Hits.Add(hit)

            Dim svc As New AutoSelectService(NewConnection(doc, New FakeCorelApplication()), New NullProgressSink())
            Dim outcome As SelectionOutcome = svc.ApplySelection(scan, 1)

            Check.AreEqual(1, doc.ClearSelectionCalls, "seleksi lama dibersihkan (ada target)")
            Check.IsTrue(outcome.WasAttempted, "operasi dijalankan")
            Check.AreEqual(1, outcome.ExpectedCount, "satu target")
            Check.IsTrue(outcome.ActualCountKnown, "jumlah terbaca dari CorelDRAW")
            Check.IsTrue(outcome.SelectionVerified, "terverifikasi")
            Check.IsTrue(target.Selected, "objek benar-benar terpilih")
        End Sub)

        suite.Add("ApplySelection: AddToSelection tidak berefek -> TIDAK diverifikasi", Sub()
            Dim doc As FakeDocument = NewDocument(1)
            Dim target As FakeShape = AddShape(doc, 1, 3, "Kotak", 20, 20)
            target.FailAddToSelection = True

            Dim hit As New ShapeHit()
            hit.SelectionTarget = target
            Dim scan As New ShapeScanResult()
            scan.PageIndex = 1
            scan.Hits.Add(hit)

            Dim svc As New AutoSelectService(NewConnection(doc, New FakeCorelApplication()), New NullProgressSink())
            Dim outcome As SelectionOutcome = svc.ApplySelection(scan, 1)

            Check.IsFalse(outcome.SelectionVerified, "tidak diklaim terverifikasi")
            Check.IsTrue(outcome.Failures.Count > 0, "kegagalan dicatat")
        End Sub)

        suite.Add("ApplySelection: ClearSelection gagal -> exception, tidak ada fallback Delete", Sub()
            Dim doc As FakeDocument = NewDocument(1)
            Dim target As FakeShape = AddShape(doc, 1, 3, "Kotak", 20, 20)
            doc.FailClearSelection = True

            Dim hit As New ShapeHit()
            hit.SelectionTarget = target
            Dim scan As New ShapeScanResult()
            scan.PageIndex = 1
            scan.Hits.Add(hit)

            Dim svc As New AutoSelectService(NewConnection(doc, New FakeCorelApplication()), New NullProgressSink())
            Dim message As String = ""
            Try
                svc.ApplySelection(scan, 1)
            Catch ex As Exception
                message = ex.Message
            End Try

            Check.Contains(message, "Tidak ada objek yang dihapus", "pesan keselamatan")
            Check.IsFalse(target.Deleted, "objek tidak dihapus")
        End Sub)

        suite.Add("navigasi maju tanpa lompat/ulang, berhenti di hasil terakhir", Sub()
            Dim doc As FakeDocument = NewDocument(3)
            AddShape(doc, 1, 3, "Kotak 1", 20, 20)
            AddShape(doc, 3, 3, "Kotak 3", 20, 20)

            Dim svc As New AutoSelectService(NewConnection(doc, New FakeCorelApplication()), New NullProgressSink())
            svc.ScanAllPages(ShapeCategory.Vector)

            Check.AreEqual(1, svc.Navigator.CurrentPage, "mulai dari hasil pertama")
            Check.AreEqual(NavigationStep.Moved, svc.MoveNext(ShapeCategory.Vector, True), "maju ke hasil 2")
            Check.AreEqual(3, svc.Navigator.CurrentPage, "halaman 3 (tidak melewati)")
            Check.AreEqual(NavigationStep.AtEnd, svc.MoveNext(ShapeCategory.Vector, True), "berhenti di hasil terakhir")
            Check.AreEqual(3, svc.Navigator.CurrentPage, "tidak kembali ke awal")
        End Sub)

        suite.Add("kategori diubah setelah pemindaian -> StaleSnapshot", Sub()
            Dim doc As FakeDocument = NewDocument(2)
            AddShape(doc, 1, 3, "Kotak", 20, 20)

            Dim svc As New AutoSelectService(NewConnection(doc, New FakeCorelApplication()), New NullProgressSink())
            svc.ScanAllPages(ShapeCategory.Vector)

            Check.AreEqual(NavigationStep.StaleSnapshot, svc.MoveNext(ShapeCategory.Text, True),
                           "navigasi ditolak, bukan memakai aturan baru")
        End Sub)

        suite.Add("Reset membatalkan sesi (dipakai saat pemindaian gagal)", Sub()
            Dim doc As FakeDocument = NewDocument(2)
            AddShape(doc, 1, 3, "Kotak", 20, 20)

            Dim svc As New AutoSelectService(NewConnection(doc, New FakeCorelApplication()), New NullProgressSink())
            svc.ScanAllPages(ShapeCategory.Vector)
            svc.Reset()

            Check.IsFalse(svc.Navigator.HasResults, "navigator kosong")
            Check.AreEqual(0, svc.Scans.Count, "hasil pemindaian dibuang")
        End Sub)
    End Sub
End Module
