Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  BatchReliabilityTests.vb
'  Fokus A (batch aman: hentikan, exception, save gagal, dokumen ditutup,
'  pembukuan file) & Fokus C (PDF proof tidak dianggap sukses tanpa validasi)
'  & Fokus D (status final sama untuk UI/ringkasan/CSV).
'
'  Memakai FakeCorel (bukan COM). Hasilnya BUKAN integration test CorelDRAW.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO

Public Module BatchReliabilityTests

    Private Function NewSettings() As WatermarkSettings
        Dim s As New WatermarkSettings()
        s.Kind = WatermarkKind.TextOnly
        s.Text = "UJI"
        s.FontName = "Segoe UI"
        s.FontSize = 12.0F
        s.Color = Color.Black
        s.AllPages = True
        s.ReplaceExisting = True
        s.LayerName = "WATERMARK"
        s.Anchor = WatermarkAnchor.Center
        s.OpacityPercent = 50
        Return s
    End Function

    Private Function NewFolder(tag As String, ParamArray names As String()) As String
        Dim folderPath As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                                                          "cpb-" & tag & "-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(folderPath)
        For Each name As String In names
            File.WriteAllText(System.IO.Path.Combine(folderPath, name), "isi asli " & name)
        Next
        Return folderPath
    End Function

    Private Sub Cleanup(folder As String)
        Try
            If Directory.Exists(folder) Then Directory.Delete(folder, True)
        Catch
        End Try
    End Sub

    ''' <summary>PDF A4 dengan jumlah halaman tertentu (dipakai FakeDocument.PublishToPDF).</summary>
    Private Function A4Pdf(pageCount As Integer) As Byte()
        Dim widths As New List(Of Double)()
        Dim heights As New List(Of Double)()
        For i As Integer = 1 To pageCount
            widths.Add(595.28)
            heights.Add(841.89)
        Next
        Return PdfFixture.BuildPdf(widths, heights)
    End Function

    Private Function CsvLines(path As String) As String()
        If Not File.Exists(path) Then Return New String() {}
        Dim lines As New List(Of String)()
        For Each line As String In File.ReadAllLines(path)
            If line.Trim().Length > 0 Then lines.Add(line)
        Next
        Return lines.ToArray()
    End Function

    Public Sub Register(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("Watermark — reliabilitas batch & PDF")

        ' ---------------------------------------------------------------- A ---
        suite.Add("file yang gagal dibuka -> GAGAL, file lain tetap OK", Sub()
            Dim folder As String = NewFolder("openfail", "a.cdr", "b.cdr", "c.cdr")
            Try
                Dim app As New FakeCorelApplication()
                app.FailOpenFor.Add("b.cdr")
                Dim engine As New WatermarkEngine()
                Dim result As WatermarkResult = engine.ApplyToFolder(app, folder, NewSettings(), False, False)

                Check.AreEqual(2, result.FilesOk, "2 file OK")
                Check.AreEqual(1, result.FilesFailed, "1 file GAGAL")
                Check.AreEqual(BatchFileStatus.Failed, result.FileOutcomes(1).Status, "b.cdr GAGAL")
                Check.Contains(result.FileOutcomes(1).MessageText(), "sedang dipakai", "alasan dicatat")
            Finally
                Cleanup(folder)
            End Try
        End Sub)

        suite.Add("batch dihentikan setelah file pertama -> sisanya DIBATALKAN", Sub()
            Dim folder As String = NewFolder("stop", "a.cdr", "b.cdr", "c.cdr")
            Try
                Dim app As New FakeCorelApplication()
                Dim engine As New WatermarkEngine()
                Dim settings As WatermarkSettings = NewSettings()
                settings.WriteCsvLog = True
                settings.CsvLogPath = Path.Combine(folder, "log.csv")

                ' Hentikan setelah file pertama selesai.
                AddHandler engine.Progress, Sub(message As String)
                                                If message.StartsWith("Status ") Then engine.RequestStop()
                                            End Sub

                Dim result As WatermarkResult = engine.ApplyToFolder(app, folder, settings, False, False)

                Check.AreEqual(1, result.FilesOk, "1 file selesai sebelum dihentikan")
                Check.AreEqual(2, result.FilesCancelled, "2 file DIBATALKAN")
                Check.AreEqual(3, result.FileOutcomes.Count, "semua file tercatat (tidak ada yang hilang)")
                Check.AreEqual(BatchFileStatus.Cancelled, result.FileOutcomes(1).Status, "status DIBATALKAN")
                Check.Contains(result.FileOutcomes(1).StatusLabel(), "DIBATALKAN", "label status")
                Check.AreEqual(1, app.Opened.Count, "hanya 1 file dibuka")

                Dim lines As String() = CsvLines(Path.Combine(folder, "log.csv"))
                Check.AreEqual(5, lines.Length, "header + 3 file + ringkasan")
                Dim summary As String = lines(lines.Length - 1)
                Check.Contains(summary, "RINGKASAN", "baris ringkasan ada")
                Check.Contains(summary, "OK 1", "ringkasan OK 1")
                Check.Contains(summary, "DIBATALKAN 2", "ringkasan DIBATALKAN 2")
            Finally
                Cleanup(folder)
            End Try
        End Sub)

        suite.Add("exception di tengah batch -> sisa file DIBATALKAN, tidak hilang dari laporan", Sub()
            Dim folder As String = NewFolder("boom", "a.cdr", "b.cdr", "c.cdr")
            Try
                Dim app As New FakeCorelApplication()
                Dim engine As New WatermarkEngine()
                Dim triggers As Integer = 0
                AddHandler engine.Progress, Sub(message As String)
                                                If message.StartsWith("Batch ") Then
                                                    triggers += 1
                                                    ' Lempar pada file ke-2 (meniru handler UI yang gagal).
                                                    If triggers = 2 Then Throw New InvalidOperationException("simulasi handler UI gagal")
                                                End If
                                            End Sub
                Dim result As WatermarkResult = engine.ApplyToFolder(app, folder, NewSettings(), False, False)

                Check.IsTrue(result.Errors.Count > 0, "error batch dicatat")
                Check.Contains(String.Join(" ", result.Errors.ToArray()), "tak terduga", "alasan exception")
                Check.AreEqual(3, result.FileOutcomes.Count, "semua file tetap punya status")
                Check.AreEqual(1, result.FilesOk, "file pertama selesai")
                Check.AreEqual(2, result.FilesCancelled, "sisa file DIBATALKAN")
            Finally
                Cleanup(folder)
            End Try
        End Sub)

        suite.Add("dokumen ditutup setelah diproses (tidak ada kebocoran)", Sub()
            Dim folder As String = NewFolder("close", "a.cdr", "b.cdr")
            Try
                Dim app As New FakeCorelApplication()
                Dim engine As New WatermarkEngine()
                Dim result As WatermarkResult = engine.ApplyToFolder(app, folder, NewSettings(), False, False)

                For Each doc As FakeDocument In app.Opened
                    Check.IsTrue(doc.Closed, "dokumen " & doc.Name & " ditutup")
                Next
                Check.AreEqual(0, result.DocumentsLeftOpen, "tidak ada dokumen tertinggal")
                Check.AreEqual(2, result.FilesOk, "2 file OK")
            Finally
                Cleanup(folder)
            End Try
        End Sub)

        suite.Add("dokumen gagal ditutup -> peringatan + jumlah dokumen terbuka dilaporkan", Sub()
            Dim folder As String = NewFolder("closefail", "a.cdr")
            Try
                Dim app As New FakeCorelApplication()
                app.Configure = Sub(doc As FakeDocument)
                                      doc.CloseFails = True
                                  End Sub
                Dim engine As New WatermarkEngine()
                Dim result As WatermarkResult = engine.ApplyToFolder(app, folder, NewSettings(), False, False)

                Check.AreEqual(1, result.DocumentsLeftOpen, "1 dokumen masih terbuka")
                Check.Contains(String.Join(" ", result.Warnings.ToArray()), "masih terbuka", "peringatan")
                Check.AreEqual(1, result.FilesOk, "output tetap tersimpan (bukan kegagalan simpan)")
            Finally
                Cleanup(folder)
            End Try
        End Sub)

        suite.Add("file sumber tidak pernah berubah isinya (mode file baru)", Sub()
            Dim folder As String = NewFolder("source", "a.cdr")
            Try
                Dim app As New FakeCorelApplication()
                Dim engine As New WatermarkEngine()
                engine.ApplyToFolder(app, folder, NewSettings(), False, False)

                Check.AreEqual("isi asli a.cdr", File.ReadAllText(Path.Combine(folder, "a.cdr")), "sumber utuh")
                Check.IsTrue(File.Exists(Path.Combine(folder, "a_wm.cdr")), "output baru dibuat")
            Finally
                Cleanup(folder)
            End Try
        End Sub)

        ' ---------------------------------------------------------------- C ---
        suite.Add("PDF proof valid (A4, jumlah halaman cocok) -> file OK & validasi dicatat", Sub()
            Dim folder As String = NewFolder("pdfok", "a.cdr")
            Try
                Dim app As New FakeCorelApplication()
                app.Configure = Sub(doc As FakeDocument)
                                      doc.PdfBytes = A4Pdf(doc.Pages.Count)
                                  End Sub
                Dim settings As WatermarkSettings = NewSettings()
                settings.ExportProofPdf = True
                Dim engine As New WatermarkEngine()
                Dim result As WatermarkResult = engine.ApplyToFolder(app, folder, settings, False, False)

                Check.AreEqual(1, result.FilesOk, "1 file OK")
                Check.IsFalse(result.ProofPdfInvalid, "PDF dianggap valid")
                Check.AreEqual(1, result.ProofPdfFiles.Count, "satu PDF proof")
                Check.AreEqual(0, result.PdfProofProblems.Count, "tanpa masalah PDF")
            Finally
                Cleanup(folder)
            End Try
        End Sub)

        suite.Add("PDF ada tetapi rusak -> SEBAGIAN, bukan OK", Sub()
            Dim folder As String = NewFolder("pdfbad", "a.cdr")
            Try
                Dim app As New FakeCorelApplication()
                ' PdfBytes = Nothing -> FakeDocument menulis "bukan pdf".
                Dim settings As WatermarkSettings = NewSettings()
                settings.ExportProofPdf = True
                Dim engine As New WatermarkEngine()
                Dim result As WatermarkResult = engine.ApplyToFolder(app, folder, settings, False, False)

                Check.IsTrue(result.ProofPdfInvalid, "PDF ditandai tidak valid")
                Check.AreEqual(1, result.FilesPartial, "status SEBAGIAN")
                Check.AreEqual(0, result.FilesOk, "tidak ada file OK")
                Check.IsTrue(result.PdfProofProblems.Count > 0, "masalah PDF dicatat")
                Check.Contains(result.FileOutcomes(0).MessageText(), "PDF proof TIDAK VALID", "pesan file")
            Finally
                Cleanup(folder)
            End Try
        End Sub)

        suite.Add("PublishToPDF tidak menulis file -> TIDAK dianggap sukses", Sub()
            Dim folder As String = NewFolder("pdfmissing", "a.cdr")
            Try
                Dim app As New FakeCorelApplication()
                app.Configure = Sub(doc As FakeDocument)
                                      doc.PublishWritesFile = False
                                  End Sub
                Dim settings As WatermarkSettings = NewSettings()
                settings.ExportProofPdf = True
                Dim engine As New WatermarkEngine()
                Dim result As WatermarkResult = engine.ApplyToFolder(app, folder, settings, False, False)

                Check.IsTrue(result.ProofPdfInvalid, "PDF ditandai tidak valid")
                Check.AreEqual(0, result.ProofPdfFiles.Count, "tidak ada PDF yang diklaim")
                Check.Contains(String.Join(" ", result.PdfProofProblems.ToArray()), "TIDAK dibuat",
                               "alasan eksplisit")
                Check.AreEqual(1, result.FilesPartial, "file SEBAGIAN")
            Finally
                Cleanup(folder)
            End Try
        End Sub)

        suite.Add("PDF ukuran halaman salah -> SEBAGIAN + masalah ukuran", Sub()
            Dim folder As String = NewFolder("pdfsize", "a.cdr")
            Try
                Dim app As New FakeCorelApplication()
                app.Configure = Sub(doc As FakeDocument)
                                      ' A5 padahal dokumen A4
                                      doc.PdfBytes = PdfFixture.BuildPdf(New Double() {419.53}, New Double() {595.28})
                                  End Sub
                Dim settings As WatermarkSettings = NewSettings()
                settings.ExportProofPdf = True
                Dim engine As New WatermarkEngine()
                Dim result As WatermarkResult = engine.ApplyToFolder(app, folder, settings, False, False)

                Check.IsTrue(result.ProofPdfInvalid, "tidak valid")
                Check.Contains(String.Join(" ", result.PdfProofProblems.ToArray()), "ukuran PDF", "alasan ukuran")
                Check.AreEqual(1, result.FilesPartial, "SEBAGIAN")
            Finally
                Cleanup(folder)
            End Try
        End Sub)

        suite.Add("PublishToPDF gagal & ExportEx gagal -> kedua pesan dilaporkan", Sub()
            Dim folder As String = NewFolder("pdfthrow", "a.cdr")
            Try
                Dim app As New FakeCorelApplication()
                app.Configure = Sub(doc As FakeDocument)
                                      doc.PublishFails = True
                                  End Sub
                Dim settings As WatermarkSettings = NewSettings()
                settings.ExportProofPdf = True
                Dim engine As New WatermarkEngine()
                Dim result As WatermarkResult = engine.ApplyToFolder(app, folder, settings, False, False)

                Check.IsTrue(result.ProofPdfInvalid, "tidak valid")
                Check.Contains(String.Join(" ", result.PdfProofProblems.ToArray()), "PublishToPDF", "pesan pertama")
                Check.Contains(String.Join(" ", result.PdfProofProblems.ToArray()), "ExportEx", "pesan fallback")
                Check.AreEqual(1, result.FilesPartial, "SEBAGIAN (CDR tetap tersimpan)")
            Finally
                Cleanup(folder)
            End Try
        End Sub)

        ' ---------------------------------------------------------------- D ---
        suite.Add("CSV: satu baris per file + ringkasan, status sama dengan UI", Sub()
            Dim folder As String = NewFolder("csv", "a.cdr", "b.cdr")
            Try
                Dim app As New FakeCorelApplication()
                app.FailOpenFor.Add("b.cdr")
                Dim settings As WatermarkSettings = NewSettings()
                settings.WriteCsvLog = True
                settings.CsvLogPath = Path.Combine(folder, "log.csv")
                Dim engine As New WatermarkEngine()
                Dim result As WatermarkResult = engine.ApplyToFolder(app, folder, settings, False, False)

                Dim lines As String() = CsvLines(Path.Combine(folder, "log.csv"))
                Check.AreEqual(4, lines.Length, "header + 2 file + ringkasan")
                Check.Contains(lines(0), "Waktu;File;Halaman;Shape;Status;Pesan;Durasi_detik", "header")
                Check.Contains(lines(1), ";OK;", "baris OK")
                Check.Contains(lines(2), ";GAGAL;", "baris GAGAL")
                Check.Contains(lines(3), "RINGKASAN", "ringkasan")
                Check.Contains(lines(3), "OK 1", "ringkasan OK 1")
                Check.Contains(lines(3), "GAGAL 1", "ringkasan GAGAL 1")
                ' UI & ringkasan memakai angka yang sama.
                Check.AreEqual(result.FilesOk, 1, "UI OK 1")
                Check.AreEqual(result.FilesFailed, 1, "UI GAGAL 1")
            Finally
                Cleanup(folder)
            End Try
        End Sub)

        suite.Add("CSV tidak dapat ditulis -> peringatan, batch tetap jalan", Sub()
            Dim folder As String = NewFolder("csvfail", "a.cdr")
            Try
                Dim app As New FakeCorelApplication()
                Dim settings As WatermarkSettings = NewSettings()
                settings.WriteCsvLog = True
                ' Path CSV diarahkan ke sebuah FOLDER: penulisan file pasti gagal
                ' di OS mana pun (deterministik, bukan bergantung karakter ilegal).
                Dim blockedPath As String = Path.Combine(folder, "log-terblokir.csv")
                Directory.CreateDirectory(blockedPath)
                settings.CsvLogPath = blockedPath
                Dim engine As New WatermarkEngine()
                Dim result As WatermarkResult = engine.ApplyToFolder(app, folder, settings, False, False)

                Check.AreEqual(1, result.FilesOk, "file tetap diproses")
                Check.Contains(String.Join(" ", result.Warnings.ToArray()), "TIDAK", "peringatan log")
            Finally
                Cleanup(folder)
            End Try
        End Sub)
    End Sub
End Module
