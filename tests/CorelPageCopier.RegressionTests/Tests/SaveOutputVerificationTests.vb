Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  SaveOutputVerificationTests.vb
'  Review lanjutan PR #2 — verifikasi penyimpanan.
'
'  Dua lapis:
'    A. kelas murni SaveOutputVerification (tanpa file/COM),
'    B. WatermarkEngine.ApplyToFolder dengan FakeCorel, untuk membuktikan
'       pemanggilan SaveAs dipisahkan dari bukti penulisan, dan bahwa file
'       lama yang masih ada TIDAK membuat status menjadi OK.
'
'  Catatan: semua ini tetap BUKAN integration test CorelDRAW (NOT RUN).
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO

Public Module SaveOutputVerificationTests

    ' --- helper data uji ----------------------------------------------------
    Private Function State(exists As Boolean, length As Long, ticks As Long) As SaveOutputSnapshot
        Return SaveOutputSnapshot.DescribeState(exists, length, New DateTime(2026, 10, 27, 3, 0, 0, DateTimeKind.Utc).AddMilliseconds(ticks))
    End Function

    Private Function UnreadableSnapshot(reason As String) As SaveOutputSnapshot
        Dim snapshot As New SaveOutputSnapshot()
        snapshot.Exists = False
        snapshot.CaptureError = reason
        Return snapshot
    End Function

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
        Dim folderPath As String = Path.Combine(Path.GetTempPath(),
                                                "cpc-save-" & tag & "-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(folderPath)
        Return folderPath
    End Function

    Private Sub CleanupFolder(folder As String)
        Try
            If Directory.Exists(folder) Then Directory.Delete(folder, True)
        Catch
        End Try
    End Sub

    ' ========================================================================
    '  A. Kelas murni
    ' ========================================================================
    Private Sub RegisterPure(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("Verifikasi output simpan (murni)")

        suite.Add("file baru muncul -> VerifiedNew", Sub()
            Dim v As SaveOutputVerification = SaveOutputVerification.Verify(State(False, 0, 0), State(True, 120, 5))
            Check.AreEqual(SaveOutputStatus.VerifiedNew, v.Status, "status")
            Check.IsTrue(v.IsVerified, "dianggap terverifikasi")
            Check.AreEqual("BARU", v.StatusLabel, "label")
        End Sub)

        suite.Add("file lama berubah ukuran -> VerifiedChanged", Sub()
            Dim v As SaveOutputVerification = SaveOutputVerification.Verify(State(True, 100, 0), State(True, 140, 0))
            Check.AreEqual(SaveOutputStatus.VerifiedChanged, v.Status, "status")
            Check.IsTrue(v.IsVerified, "terverifikasi")
        End Sub)

        suite.Add("file lama berubah waktu modifikasi -> VerifiedChanged", Sub()
            Dim v As SaveOutputVerification = SaveOutputVerification.Verify(State(True, 100, 0), State(True, 100, 7))
            Check.AreEqual(SaveOutputStatus.VerifiedChanged, v.Status, "status")
        End Sub)

        suite.Add("file lama TIDAK berubah -> Unchanged (bukan sukses)", Sub()
            Dim v As SaveOutputVerification = SaveOutputVerification.Verify(State(True, 100, 3), State(True, 100, 3))
            Check.AreEqual(SaveOutputStatus.Unchanged, v.Status, "status")
            Check.IsFalse(v.IsVerified, "TIDAK dianggap terverifikasi")
            Check.Contains(v.Reason, "tidak berubah", "alasan eksplisit")
            Check.AreEqual("TIDAK BERUBAH", v.StatusLabel, "label")
        End Sub)

        suite.Add("file hilang setelah SaveAs -> Missing", Sub()
            Dim v As SaveOutputVerification = SaveOutputVerification.Verify(State(False, 0, 0), State(False, 0, 0))
            Check.AreEqual(SaveOutputStatus.Missing, v.Status, "status")
            Check.Contains(v.Reason, "tidak ditemukan", "alasan")
            Check.IsFalse(v.IsVerified, "tidak terverifikasi")
        End Sub)

        suite.Add("output 0 byte -> Empty (bukan sukses)", Sub()
            Dim v As SaveOutputVerification = SaveOutputVerification.Verify(State(False, 0, 0), State(True, 0, 9))
            Check.AreEqual(SaveOutputStatus.Empty, v.Status, "status")
            Check.Contains(v.Reason, "0 byte", "alasan")
            Check.IsFalse(v.IsVerified, "tidak terverifikasi")
        End Sub)

        suite.Add("keadaan sesudah tidak terbaca -> Unreadable", Sub()
            Dim v As SaveOutputVerification = SaveOutputVerification.Verify(State(False, 0, 0),
                                                                           UnreadableSnapshot("IOException: akses ditolak"))
            Check.AreEqual(SaveOutputStatus.Unreadable, v.Status, "status")
            Check.Contains(v.Reason, "akses ditolak", "alasan asli diteruskan")
            Check.IsFalse(v.IsVerified, "tidak terverifikasi")
        End Sub)

        suite.Add("keadaan SEBELUM tidak terbaca -> tidak ada bukti (Unreadable)", Sub()
            Dim v As SaveOutputVerification = SaveOutputVerification.Verify(UnreadableSnapshot("path output kosong"),
                                                                           State(True, 120, 5))
            Check.AreEqual(SaveOutputStatus.Unreadable, v.Status, "status")
            Check.Contains(v.Reason, "sebelum SaveAs", "alasan menyebut tahap sebelum")
            Check.IsFalse(v.IsVerified, "tidak terverifikasi")
        End Sub)

        suite.Add("mode timpa tanpa file sebelumnya -> Unreadable (data tidak masuk akal)", Sub()
            Dim v As SaveOutputVerification = SaveOutputVerification.Verify(State(False, 0, 0), State(True, 120, 5),
                                                                           SaveOutputMode.OverwriteInPlace)
            Check.AreEqual(SaveOutputStatus.Unreadable, v.Status, "status")
            Check.Contains(v.Reason, "mode timpa", "alasan menyebut mode")
        End Sub)

        suite.Add("mode timpa dengan file berubah -> VerifiedChanged", Sub()
            Dim v As SaveOutputVerification = SaveOutputVerification.Verify(State(True, 100, 1), State(True, 160, 2),
                                                                           SaveOutputMode.OverwriteInPlace)
            Check.AreEqual(SaveOutputStatus.VerifiedChanged, v.Status, "status")
            Check.IsTrue(v.IsVerified, "terverifikasi")
        End Sub)

        suite.Add("mode timpa dengan file tetap sama -> Unchanged", Sub()
            Dim v As SaveOutputVerification = SaveOutputVerification.Verify(State(True, 100, 1), State(True, 100, 1),
                                                                           SaveOutputMode.OverwriteInPlace)
            Check.AreEqual(SaveOutputStatus.Unchanged, v.Status, "status")
            Check.IsFalse(v.IsVerified, "file lama bukan bukti sukses")
        End Sub)

        suite.Add("Snapshot.Capture membaca file nyata & path kosong dengan jujur", Sub()
            Dim folder As String = NewTempFolder("capture")
            Try
                Dim filePath As String = Path.Combine(folder, "a_wm.cdr")
                File.WriteAllText(filePath, "12345")

                Dim exists As SaveOutputSnapshot = SaveOutputSnapshot.Capture(filePath)
                Check.IsTrue(exists.IsValid, "terbaca")
                Check.IsTrue(exists.Exists, "file ada")
                Check.AreEqual(5L, exists.Length, "ukuran")

                Dim missing As SaveOutputSnapshot = SaveOutputSnapshot.Capture(Path.Combine(folder, "tidak-ada.cdr"))
                Check.IsTrue(missing.IsValid, "terbaca")
                Check.IsFalse(missing.Exists, "file tidak ada")

                Dim empty As SaveOutputSnapshot = SaveOutputSnapshot.Capture("")
                Check.IsFalse(empty.IsValid, "path kosong bukan keadaan valid")
                Check.Contains(empty.CaptureError, "kosong", "alasan")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)
    End Sub

    ' ========================================================================
    '  B. WatermarkEngine (FakeCorel) — mode file baru & mode timpa
    ' ========================================================================
    Private Sub RegisterEngine(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("Verifikasi output simpan (WatermarkEngine)")

        suite.Add("SaveAs melempar -> GAGAL, pemanggilan tercatat gagal", Sub()
            Dim folder As String = NewTempFolder("throw")
            Try
                File.WriteAllText(Path.Combine(folder, "a.cdr"), "x")
                Dim app As New FakeCorelApplication()
                Dim captured As FakeDocument = Nothing
                app.Configure = Sub(doc As FakeDocument)
                                      doc.SaveFails = True
                                      captured = doc
                                  End Sub

                Dim result As WatermarkResult = New WatermarkEngine().ApplyToFolder(app, folder, NewTextSettings(), False, False)

                Check.AreEqual(1, result.FilesFailed, "1 file GAGAL")
                Check.AreEqual(0, result.FilesOk, "0 file OK")
                Dim outcome As BatchFileOutcome = result.FileOutcomes(0)
                Check.IsTrue(outcome.SaveAttempted, "save dicoba")
                Check.IsFalse(outcome.SaveCallSucceeded, "pemanggilan SaveAs gagal")
                Check.IsFalse(outcome.SaveSucceeded, "penyimpanan tidak dianggap sukses")
                Check.Contains(outcome.MessageText(), "SAVE GAGAL", "pesan")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)

        suite.Add("mode file baru: SaveAs tidak menulis -> GAGAL (bukan OK)", Sub()
            Dim folder As String = NewTempFolder("newnofile")
            Try
                File.WriteAllText(Path.Combine(folder, "a.cdr"), "x")
                Dim app As New FakeCorelApplication()
                app.Configure = Sub(doc As FakeDocument)
                                      doc.SaveWritesFile = False
                                  End Sub

                Dim result As WatermarkResult = New WatermarkEngine().ApplyToFolder(app, folder, NewTextSettings(), False, False)
                Check.AreEqual(1, result.FilesFailed, "GAGAL")
                Check.Contains(String.Join(" ", result.Failures.ToArray()), "tidak ditemukan", "alasan")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)

        suite.Add("mode file baru: pemanggilan ok tapi file lama tetap -> TIDAK BERUBAH, GAGAL", Sub()
            Dim folder As String = NewTempFolder("staleoutput")
            Try
                File.WriteAllText(Path.Combine(folder, "a.cdr"), "x")
                ' Output dari eksekusi sebelumnya sudah ada. Isinya dibuat
                ' menyerupai berkas CorelDRAW yang sah (kontainer RIFF/CDRX),
                ' supaya yang diuji benar-benar cabang "file lama masih ada dan
                ' tidak berubah" — bukan cabang "isi bukan CDR".
                Dim stalePath As String = Path.Combine(folder, "a_wm.cdr")
                File.WriteAllBytes(stalePath, FakeDocument.FakeCdrBytes(False, False))
                Dim staleTime As DateTime = File.GetLastWriteTimeUtc(stalePath)

                Dim app As New FakeCorelApplication()
                app.Configure = Sub(doc As FakeDocument)
                                      doc.SaveWritesFile = False   ' SaveAs tidak menulis apa pun
                                  End Sub

                Dim settings As WatermarkSettings = NewTextSettings()
                settings.WriteCsvLog = True
                settings.CsvLogPath = Path.Combine(folder, "log.csv")

                Dim result As WatermarkResult = New WatermarkEngine().ApplyToFolder(app, folder, settings, False, False)

                Check.AreEqual(1, result.FilesFailed, "file GAGAL, bukan OK")
                Dim outcome As BatchFileOutcome = result.FileOutcomes(0)
                Check.IsTrue(outcome.SaveCallSucceeded, "pemanggilan SaveAs tidak melempar")
                Check.IsFalse(outcome.SaveSucceeded, "penyimpanan TIDAK dianggap sukses")
                Check.AreEqual("TIDAK BERUBAH", outcome.SaveVerificationLabel, "label verifikasi")
                Check.Contains(outcome.MessageText(), "SAVE TIDAK TERVERIFIKASI", "pesan membedakan pemanggilan vs bukti")
                Check.AreEqual(staleTime, File.GetLastWriteTimeUtc(stalePath), "file lama tidak tersentuh")
                Check.Contains(File.ReadAllText(settings.CsvLogPath), "GAGAL", "CSV memakai status yang sama")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)

        suite.Add("mode TIMPA: SaveAs tidak menulis, file lama masih ada -> GAGAL (inti butir 2)", Sub()
            Dim folder As String = NewTempFolder("overwritestale")
            Try
                Dim sourcePath As String = Path.Combine(folder, "a.cdr")
                ' Berkas sumber yang sah (RIFF/CDRX): fokus test adalah "file lama
                ' tidak berubah", bukan "isi bukan CDR".
                File.WriteAllBytes(sourcePath, FakeDocument.FakeCdrBytes(False, False))
                Dim originalTime As DateTime = File.GetLastWriteTimeUtc(sourcePath)

                Dim app As New FakeCorelApplication()
                app.Configure = Sub(doc As FakeDocument)
                                      doc.SaveWritesFile = False
                                  End Sub

                Dim settings As WatermarkSettings = NewTextSettings()
                settings.WriteCsvLog = True
                settings.CsvLogPath = Path.Combine(folder, "log.csv")

                ' includeSubfolders:=False, saveInPlace:=True
                Dim result As WatermarkResult = New WatermarkEngine().ApplyToFolder(app, folder, settings, False, True)

                Check.AreEqual(1, result.FilesFailed, "GAGAL: file lama tidak boleh dianggap bukti")
                Check.AreEqual(0, result.FilesOk, "0 file OK")
                Dim outcome As BatchFileOutcome = result.FileOutcomes(0)
                Check.IsTrue(outcome.SaveCallSucceeded, "SaveAs tidak melempar")
                Check.IsFalse(outcome.SaveSucceeded, "tidak dianggap tersimpan")
                Check.AreEqual("TIDAK BERUBAH", outcome.SaveVerificationLabel, "bukti penulisan tidak ada")
                Check.IsTrue(File.Exists(sourcePath), "file sumber memang masih ada (itu masalahnya)")
                Check.AreEqual(originalTime, File.GetLastWriteTimeUtc(sourcePath), "file tidak berubah")
                Check.Contains(File.ReadAllText(settings.CsvLogPath), "GAGAL", "CSV GAGAL")
                Check.DoesNotContain(File.ReadAllText(settings.CsvLogPath), ";OK;", "tidak ada OK palsu")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)

        suite.Add("mode TIMPA: SaveAs benar-benar menulis -> OK (DIPERBARUI)", Sub()
            Dim folder As String = NewTempFolder("overwriteok")
            Try
                Dim sourcePath As String = Path.Combine(folder, "a.cdr")
                File.WriteAllText(sourcePath, "isi asli")

                Dim app As New FakeCorelApplication()
                Dim settings As WatermarkSettings = NewTextSettings()
                Dim result As WatermarkResult = New WatermarkEngine().ApplyToFolder(app, folder, settings, False, True)

                Check.AreEqual(1, result.FilesOk, "OK")
                Dim outcome As BatchFileOutcome = result.FileOutcomes(0)
                Check.IsTrue(outcome.SaveSucceeded, "tersimpan & terverifikasi")
                Check.AreEqual("DIPERBARUI", outcome.SaveVerificationLabel, "terbukti berubah")
                Check.Contains(outcome.MessageText(), "save OK", "pesan")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)

        suite.Add("mode file baru: output 0 byte -> GAGAL (KOSONG)", Sub()
            Dim folder As String = NewTempFolder("emptyout")
            Try
                File.WriteAllText(Path.Combine(folder, "a.cdr"), "x")
                Dim app As New FakeCorelApplication()
                app.Configure = Sub(doc As FakeDocument)
                                      doc.SaveWritesEmptyFile = True
                                  End Sub

                Dim result As WatermarkResult = New WatermarkEngine().ApplyToFolder(app, folder, NewTextSettings(), False, False)
                Check.AreEqual(1, result.FilesFailed, "GAGAL")
                Dim outcome As BatchFileOutcome = result.FileOutcomes(0)
                Check.AreEqual("KOSONG", outcome.SaveVerificationLabel, "label")
                Check.Contains(outcome.MessageText(), "0 byte", "alasan ukuran")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)

        suite.Add("verifikasi TIDAK membuka ulang dokumen (dokumen aktif tidak terganggu)", Sub()
            Dim folder As String = NewTempFolder("noopen")
            Try
                File.WriteAllText(Path.Combine(folder, "a.cdr"), "x")
                File.WriteAllText(Path.Combine(folder, "b.cdr"), "x")
                Dim app As New FakeCorelApplication()

                Dim result As WatermarkResult = New WatermarkEngine().ApplyToFolder(app, folder, NewTextSettings(), False, False)

                Check.AreEqual(2, app.Opened.Count, "hanya 2 pembukaan: satu per file sumber")
                For Each doc As FakeDocument In app.Opened
                    Check.IsFalse(doc.OriginalPath.EndsWith("_wm.cdr", StringComparison.OrdinalIgnoreCase),
                                  "output tidak pernah dibuka untuk diperiksa")
                Next
                Check.AreEqual(2, result.FilesOk, "dua file OK")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)
    End Sub

    Public Sub Register(runner As TestRunner)
        RegisterPure(runner)
        RegisterEngine(runner)
    End Sub
End Module
