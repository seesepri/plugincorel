Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  EngineTests.vb
'  Test WatermarkEngine memakai objek CorelDRAW PALSU (FakeCorel.vb).
'  Ini bukan integration test CorelDRAW sebenarnya (COM tidak dipakai), tetapi
'  menguji jalur nyata: normalisasi satuan, undo group, status batch per file,
'  penyimpanan output, dan filter output lama.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO

Public Module EngineTests

    Private Function NewTextSettings() As WatermarkSettings
        Dim settings As New WatermarkSettings()
        settings.Kind = WatermarkKind.TextOnly
        settings.Text = "UJI"
        settings.FontName = "Segoe UI"
        settings.FontSize = 12.0F
        settings.Color = Color.Black
        settings.AllPages = True
        settings.ReplaceExisting = True
        settings.LayerName = "WATERMARK"
        settings.Anchor = WatermarkAnchor.Center
        settings.OpacityPercent = 50
        Return settings
    End Function

    Private Function NewTempFolder(tag As String) As String
        Dim folderPath As String = Path.Combine(Path.GetTempPath(), "cpc-tests-" & tag & "-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(folderPath)
        Return folderPath
    End Function

    Private Sub CleanupFolder(folder As String)
        Try
            If Directory.Exists(folder) Then Directory.Delete(folder, True)
        Catch
        End Try
    End Sub

    Private Function NewDocument(pageCount As Integer) As FakeDocument
        Dim doc As New FakeDocument("C:\uji\dokumen.cdr")
        For i As Integer = 1 To pageCount
            Dim page As New FakePage(i)
            page.EnsureActiveLayer()
            doc.Pages.Add(page)
        Next
        doc.ActivePage = doc.Pages(1)
        Return doc
    End Function

    Public Sub Register(runner As TestRunner)
        RegisterSingleDocument(runner)
        RegisterFolderBatch(runner)
    End Sub

    ' =====================================================================
    '  SATU DOKUMEN
    ' =====================================================================
    Private Sub RegisterSingleDocument(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("Watermark — satu dokumen")

        suite.Add("watermark teks sukses: shape dibuat, undo group ditutup", Sub()
            Dim doc As FakeDocument = NewDocument(2)
            Dim engine As New WatermarkEngine()
            Dim result As WatermarkResult = engine.ApplyToDocument(doc, NewTextSettings())

            Check.AreEqual(0, result.Errors.Count, "tanpa error")
            Check.AreEqual(0, result.Failures.Count, "tanpa kegagalan")
            Check.AreEqual(2, result.PagesIntended, "2 halaman target")
            Check.AreEqual(2, result.PagesPassed, "2 halaman diproses")
            Check.IsTrue(result.ShapesCreated >= 2, "shape watermark dibuat")
            Check.AreEqual(1, doc.CommandGroupsBegun.Count, "satu grup undo dibuka")
            Check.IsFalse(doc.CommandGroupOpen, "grup undo ditutup kembali")
        End Sub)

        suite.Add("satuan dipulihkan ke nilai asli (inch)", Sub()
            Dim doc As FakeDocument = NewDocument(1)
            doc.Unit = DocumentUnits.Inch
            doc.UnitSetAttempts.Clear()      ' jangan hitung penyiapan fixture
            Dim engine As New WatermarkEngine()
            engine.ApplyToDocument(doc, NewTextSettings())
            Check.AreEqual(DocumentUnits.Inch, doc.Unit, "satuan asli dipulihkan")
            Check.AreEqual(2, doc.UnitSetAttempts.Count, "dua perubahan: ke mm lalu kembali")
            Check.AreEqual(DocumentUnits.Millimeter, doc.UnitSetAttempts(0), "pertama: mm")
            Check.AreEqual(DocumentUnits.Inch, doc.UnitSetAttempts(1), "kedua: inch")
        End Sub)

        suite.Add("Document.Unit tidak terbaca -> batal tanpa mengubah dokumen", Sub()
            Dim doc As FakeDocument = NewDocument(1)
            doc.FailUnitRead = True
            Dim engine As New WatermarkEngine()
            Dim result As WatermarkResult = engine.ApplyToDocument(doc, NewTextSettings())

            Check.IsTrue(result.Errors.Count > 0, "error dilaporkan")
            Check.Contains(String.Join(" ", result.Errors.ToArray()), "tidak dapat dibaca", "alasan eksplisit")
            Check.AreEqual(0, result.ShapesCreated, "tidak ada shape dibuat")
            Check.AreEqual(0, doc.CommandGroupsBegun.Count, "tidak ada grup undo")
            Check.AreEqual(0, doc.Pages(1).ActiveLayer.Shapes.Count, "layer tidak tersentuh")
        End Sub)

        suite.Add("penyetelan mm gagal -> batal (tata letak mm tidak boleh diasumsikan)", Sub()
            Dim doc As FakeDocument = NewDocument(1)
            doc.Unit = DocumentUnits.Inch
            doc.FailUnitSet = True
            Dim engine As New WatermarkEngine()
            Dim result As WatermarkResult = engine.ApplyToDocument(doc, NewTextSettings())

            Check.IsTrue(result.Errors.Count > 0, "error dilaporkan")
            Check.AreEqual(0, result.ShapesCreated, "tidak membuat shape dengan satuan salah")
            Check.AreEqual(DocumentUnits.Inch, doc.Unit, "satuan tidak berubah")
        End Sub)

        suite.Add("pemulihan satuan gagal -> dilaporkan sebagai error", Sub()
            Dim doc As FakeDocument = NewDocument(1)
            doc.Unit = DocumentUnits.Inch
            doc.FailUnitRestore = True
            Dim engine As New WatermarkEngine()
            Dim result As WatermarkResult = engine.ApplyToDocument(doc, NewTextSettings())
            Dim joined As String = String.Join(" ", result.Errors.ToArray())
            Check.Contains(joined, "Pemulihan satuan", "error pemulihan dilaporkan")
            Check.IsFalse(result.Ok, "hasil tidak dianggap OK")
        End Sub)

        suite.Add("BeginCommandGroup gagal -> dilaporkan sebagai peringatan", Sub()
            Dim doc As FakeDocument = NewDocument(1)
            doc.FailBeginCommandGroup = True
            Dim engine As New WatermarkEngine()
            Dim result As WatermarkResult = engine.ApplyToDocument(doc, NewTextSettings())

            Check.IsTrue(result.Warnings.Count > 0, "peringatan undo dicatat")
            Check.AreEqual(0, doc.CommandGroupsBegun.Count, "tidak ada grup terbuka")
            Check.IsTrue(result.ShapesCreated > 0, "watermark tetap diterapkan")
        End Sub)

        suite.Add("EndCommandGroup gagal -> dilaporkan, tidak ditelan", Sub()
            Dim doc As FakeDocument = NewDocument(1)
            doc.FailEndCommandGroup = True
            Dim engine As New WatermarkEngine()
            Dim result As WatermarkResult = engine.ApplyToDocument(doc, NewTextSettings())

            Check.Contains(String.Join(" ", result.Errors.ToArray()), "EndCommandGroup gagal", "error End dilaporkan")
            Check.IsFalse(result.Ok, "hasil tidak dianggap OK")
        End Sub)

        suite.Add("halaman gagal -> PagesFailed dihitung (status SEBAGIAN)", Sub()
            Dim doc As FakeDocument = NewDocument(2)
            doc.Pages(2).ActiveLayer.FailCreateText = True
            doc.Pages(2).FailCreateText = True
            Dim engine As New WatermarkEngine()
            Dim result As WatermarkResult = engine.ApplyToDocument(doc, NewTextSettings())

            Check.AreEqual(2, result.PagesIntended, "2 target")
            Check.AreEqual(1, result.PagesPassed, "1 berhasil")
            Check.AreEqual(1, result.PagesFailed, "1 gagal")
            Dim status As BatchFileStatus = BatchStatusRules.EvaluateApplyOnly(
                True, result.PagesIntended, result.PagesPassed, result.PagesFailed, result.Errors.Count > 0)
            Check.AreEqual(BatchFileStatus.PartialSuccess, status, "status SEBAGIAN")
        End Sub)

        suite.Add("target halaman tidak valid -> error, bukan sukses palsu", Sub()
            Dim doc As FakeDocument = NewDocument(1)
            Dim settings As WatermarkSettings = NewTextSettings()
            settings.AllPages = False
            settings.TargetPages.Add(5)
            Dim engine As New WatermarkEngine()
            Dim result As WatermarkResult = engine.ApplyToDocument(doc, settings)
            Check.IsTrue(result.Errors.Count > 0, "error dilaporkan")
            Check.AreEqual(0, result.PagesPassed, "tidak ada halaman dianggap sukses")
        End Sub)

        suite.Add("watermark berulang tidak menumpuk (ReplaceExisting)", Sub()
            Dim doc As FakeDocument = NewDocument(1)
            Dim engine As New WatermarkEngine()
            engine.ApplyToDocument(doc, NewTextSettings())
            Dim afterFirst As Integer = doc.Pages(1).ActiveLayer.Shapes.Count
            engine.ApplyToDocument(doc, NewTextSettings())
            Dim afterSecond As Integer = doc.Pages(1).ActiveLayer.Shapes.Count
            Check.AreEqual(afterFirst, afterSecond, "jumlah shape tetap (idempoten)")
        End Sub)
    End Sub

    ' =====================================================================
    '  BATCH FOLDER
    ' =====================================================================
    Private Sub RegisterFolderBatch(runner As TestRunner)

        Dim suite As TestSuite = runner.AddSuite("Watermark — batch folder")

        suite.Add("output lama (_wm) tidak diproses ulang", Sub()
            Dim folder As String = NewTempFolder("reprocess")
            Try
                File.WriteAllText(Path.Combine(folder, "a.cdr"), "x")
                File.WriteAllText(Path.Combine(folder, "b_wm.cdr"), "x")

                Dim app As New FakeCorelApplication()
                Dim settings As WatermarkSettings = NewTextSettings()
                Dim engine As New WatermarkEngine()
                Dim result As WatermarkResult = engine.ApplyToFolder(app, folder, settings, False, False)

                Check.AreEqual(1, app.Opened.Count, "hanya 1 file dibuka")
                Check.AreEqual(Path.Combine(folder, "a.cdr"), app.Opened(0).OriginalPath, "hanya a.cdr")
                Check.AreEqual(1, result.SkippedFiles.Count, "b_wm.cdr dilewati")
                Check.Contains(result.SkippedFiles(0).Reason, "output batch sebelumnya", "alasan dicatat")
                Check.IsTrue(File.Exists(Path.Combine(folder, "a_wm.cdr")), "output a_wm.cdr dibuat")
                Check.IsFalse(File.Exists(Path.Combine(folder, "a_wm_wm.cdr")), "tidak ada rantai _wm_wm")
                Check.AreEqual(1, result.FilesOk, "1 file OK")
                Check.AreEqual(0, result.FilesFailed, "0 gagal")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)

        suite.Add("semua file adalah output lama -> error jelas, tidak memproses apa pun", Sub()
            Dim folder As String = NewTempFolder("alloutputs")
            Try
                File.WriteAllText(Path.Combine(folder, "a_wm.cdr"), "x")
                Dim app As New FakeCorelApplication()
                Dim engine As New WatermarkEngine()
                Dim result As WatermarkResult = engine.ApplyToFolder(app, folder, NewTextSettings(), False, False)
                Check.AreEqual(0, app.Opened.Count, "tidak ada file dibuka")
                Check.IsTrue(result.Errors.Count > 0, "error dilaporkan")
                Check.Contains(String.Join(" ", result.Errors.ToArray()), "dilewati", "menyebut file dilewati")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)

        suite.Add("SaveAs gagal -> status GAGAL, bukan OK", Sub()
            Dim folder As String = NewTempFolder("savefail")
            Try
                File.WriteAllText(Path.Combine(folder, "a.cdr"), "x")
                Dim app As New FakeCorelApplication()
                app.Configure = Sub(doc As FakeDocument)
                                      doc.SaveFails = True
                                  End Sub
                Dim settings As WatermarkSettings = NewTextSettings()
                settings.WriteCsvLog = True
                settings.CsvLogPath = Path.Combine(folder, "log.csv")

                Dim engine As New WatermarkEngine()
                Dim result As WatermarkResult = engine.ApplyToFolder(app, folder, settings, False, False)

                Check.AreEqual(0, result.FilesOk, "0 file OK")
                Check.AreEqual(1, result.FilesFailed, "1 file GAGAL")
                Check.IsFalse(result.Ok, "hasil keseluruhan tidak OK")
                Check.IsTrue(result.Failures.Count > 0, "kegagalan save dicatat")
                Check.Contains(String.Join(" ", result.Failures.ToArray()), "gagal menyimpan", "pesan kegagalan")
                Dim csv As String = File.ReadAllText(settings.CsvLogPath)
                Check.Contains(csv, "GAGAL", "CSV memakai status yang sama")
                Check.DoesNotContain(csv, ";OK;", "tidak ada status OK palsu")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)

        suite.Add("SaveAs tidak menulis file -> verifikasi save menolak OK", Sub()
            Dim folder As String = NewTempFolder("nosave")
            Try
                File.WriteAllText(Path.Combine(folder, "a.cdr"), "x")
                Dim app As New FakeCorelApplication()
                app.Configure = Sub(doc As FakeDocument)
                                      doc.SaveWritesFile = False
                                  End Sub
                Dim engine As New WatermarkEngine()
                Dim result As WatermarkResult = engine.ApplyToFolder(app, folder, NewTextSettings(), False, False)

                Check.AreEqual(1, result.FilesFailed, "file dianggap gagal")
                Check.Contains(String.Join(" ", result.Failures.ToArray()), "tidak ditemukan", "verifikasi file menyatakan gagal")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)

        suite.Add("watermark hanya sebagian halaman -> status SEBAGIAN, bukan OK", Sub()
            Dim folder As String = NewTempFolder("partial")
            Try
                File.WriteAllText(Path.Combine(folder, "a.cdr"), "x")
                Dim app As New FakeCorelApplication()
                app.Configure = Sub(doc As FakeDocument)
                                      ' Halaman 2: semua layer (termasuk layer watermark
                                      ' yang baru dibuat) gagal membuat teks.
                                      Dim page2 As New FakePage(2)
                                      page2.FailCreateText = True
                                      page2.EnsureActiveLayer().FailCreateText = True
                                      doc.Pages.Add(page2)
                                  End Sub
                Dim settings As WatermarkSettings = NewTextSettings()
                settings.WriteCsvLog = True
                settings.CsvLogPath = Path.Combine(folder, "log.csv")

                Dim engine As New WatermarkEngine()
                Dim result As WatermarkResult = engine.ApplyToFolder(app, folder, settings, False, False)

                Check.AreEqual(0, result.FilesOk, "bukan OK")
                Check.AreEqual(1, result.FilesPartial, "1 SEBAGIAN")
                Dim csv As String = File.ReadAllText(settings.CsvLogPath)
                Check.Contains(csv, "SEBAGIAN", "CSV menandai SEBAGIAN")
                Check.Contains(csv, "OK 0", "ringkasan CSV jujur")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)

        suite.Add("batch sukses penuh -> OK dan CSV konsisten", Sub()
            Dim folder As String = NewTempFolder("ok")
            Try
                File.WriteAllText(Path.Combine(folder, "a.cdr"), "x")
                File.WriteAllText(Path.Combine(folder, "b.cdr"), "x")
                Dim app As New FakeCorelApplication()
                Dim settings As WatermarkSettings = NewTextSettings()
                settings.WriteCsvLog = True
                settings.CsvLogPath = Path.Combine(folder, "log.csv")

                Dim engine As New WatermarkEngine()
                Dim result As WatermarkResult = engine.ApplyToFolder(app, folder, settings, False, False)

                Check.AreEqual(2, result.FilesOk, "2 file OK")
                Check.AreEqual(0, result.FilesFailed, "0 gagal")
                Check.IsTrue(result.Ok, "hasil OK")
                Check.IsTrue(File.Exists(Path.Combine(folder, "a_wm.cdr")), "output a")
                Check.IsTrue(File.Exists(Path.Combine(folder, "b_wm.cdr")), "output b")
                Check.IsTrue(File.Exists(Path.Combine(folder, "a.cdr")), "file asli tetap ada")
                Dim csv As String = File.ReadAllText(settings.CsvLogPath)
                Check.Contains(csv, "OK", "CSV OK")
                Check.Contains(csv, "RINGKASAN", "baris ringkasan")
                Check.Contains(csv, "OK 2", "ringkasan menghitung 2 OK")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)

        suite.Add("mode timpa: file read-only / path di luar folder ditolak", Sub()
            Dim folder As String = NewTempFolder("overwrite")
            Try
                Dim source As String = Path.Combine(folder, "a.cdr")
                File.WriteAllText(source, "x")
                Check.AreEqual("", BatchFilePolicy.ValidateOverwriteTarget(source, folder, False), "valid di folder")

                Dim outside As String = Path.Combine(Path.GetTempPath(), "cpc-outside-" & Guid.NewGuid().ToString("N") & ".cdr")
                Check.IsTrue(BatchFilePolicy.ValidateOverwriteTarget(outside, folder, False).Length > 0,
                             "path di luar folder ditolak")

                ' Mode timpa dijalankan pada file yang valid -> menimpa file asli (disengaja).
                Dim app As New FakeCorelApplication()
                Dim engine As New WatermarkEngine()
                Dim result As WatermarkResult = engine.ApplyToFolder(app, folder, NewTextSettings(), False, True)
                Check.AreEqual(1, app.Opened.Count, "file dibuka")
                Check.AreEqual(1, result.FilesOk, "mode timpa berhasil")
                Check.AreEqual(0, result.SkippedFiles.Count, "tidak ada yang dilewati")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)

        suite.Add("mode file baru: output tidak pernah menimpa file sumber", Sub()
            Dim folder As String = NewTempFolder("nosource")
            Try
                File.WriteAllText(Path.Combine(folder, "a.cdr"), "isi asli")
                Dim app As New FakeCorelApplication()
                Dim engine As New WatermarkEngine()
                engine.ApplyToFolder(app, folder, NewTextSettings(), False, False)
                Check.AreEqual("isi asli", File.ReadAllText(Path.Combine(folder, "a.cdr")), "file asli tidak berubah")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)

        suite.Add("dua sumber dengan output sama -> nama unik (tidak saling menimpa)", Sub()
            Dim folder As String = NewTempFolder("collide")
            Try
                ' "a.cdr" -> "a_wm.cdr"; "a_wm.cdr" dilewati sebagai output lama.
                ' Simulasi benturan: suffix "_wm" dan sumber bernama "a_wm_.cdr" tidak
                ' benturan, jadi uji lewat fungsi kebijakan murni.
                Dim reserved As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Dim first As String = Path.Combine(folder, "a_wm.cdr")
                reserved.Add(first)
                Dim second As String = BatchFilePolicy.ResolveUniqueOutputPath(first, Function(p As String) reserved.Contains(p))
                Check.AreEqual(Path.Combine(folder, "a_wm_2.cdr"), second, "nama kedua unik")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)
    End Sub
End Module
