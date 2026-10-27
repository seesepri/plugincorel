Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  SaveOutputContentTests.vb
'  Review follow-up KEDUA pada PR #2 — bukti penyimpanan lebih kuat.
'
'  Fokus (butir 2 review):
'    a. bentuk isi file diperiksa (RIFF/CDR atau ZIP), sehingga "ada + tidak
'       kosong" tidak otomatis berarti VALID,
'    b. perubahan isi (sidik jari) dihitung sebagai bukti, sehingga penulisan
'       ulang dengan panjang sama tidak dilaporkan "tidak berubah" hanya karena
'       granularitas waktu filesystem,
'    c. mode file baru & mode timpa diuji terpisah, termasuk kegagalan SaveAs,
'       target lama yang masih ada, output kosong, dan output tidak valid,
'    d. tidak ada pembukaan ulang dokumen (tetap murni operasi file).
'
'  Catatan: ini BUKAN integration test CorelDRAW (NOT RUN).
' ============================================================================

Imports System
Imports System.IO
Imports System.Drawing

Public Module SaveOutputContentTests

    ' --- helper -------------------------------------------------------------
    Private Function NewTempFolder(tag As String) As String
        Dim folderPath As String = Path.Combine(Path.GetTempPath(),
                                                "cpc-content-" & tag & "-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(folderPath)
        Return folderPath
    End Function

    Private Sub CleanupFolder(folder As String)
        Try
            If Directory.Exists(folder) Then Directory.Delete(folder, True)
        Catch
        End Try
    End Sub

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

    Private Function Utc(second As Integer) As DateTime
        Return New DateTime(2026, 10, 27, 4, 0, second, DateTimeKind.Utc)
    End Function

    Public Sub Register(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("Bukti penyimpanan: bentuk isi & sidik jari (review kedua)")

        ' ------------------------------------------------------------------
        ' A. Klasifikasi bentuk berkas (murni, tanpa file)
        ' ------------------------------------------------------------------
        suite.Add("ClassifySignature: RIFF/CDR, ZIP, dan bentuk asing dibedakan", Sub()
            Dim riff As Byte() = FakeDocument.FakeCdrBytes(False, False)
            Check.AreEqual(SaveOutputSnapshot.SignatureRiffCdr,
                           SaveOutputSnapshot.ClassifySignature(riff, riff.Length), "RIFF + CDRX")

            ' RIFF tetapi bukan form type CorelDRAW (mis. WAVE).
            Dim wave As Byte() = System.Text.Encoding.ASCII.GetBytes("RIFF" & ChrW(0) & ChrW(0) & ChrW(0) & ChrW(0) & "WAVE")
            Check.AreEqual(SaveOutputSnapshot.SignatureUnknown,
                           SaveOutputSnapshot.ClassifySignature(wave, wave.Length), "RIFF non-CDR ditolak")

            Dim zip As Byte() = FakeDocument.FakeCdrBytes(False, True)
            Check.AreEqual(SaveOutputSnapshot.SignatureZip,
                           SaveOutputSnapshot.ClassifySignature(zip, zip.Length), "ZIP (CDR X4+/2022)")

            Dim junk As Byte() = System.Text.Encoding.ASCII.GetBytes("ini bukan cdr")
            Check.AreEqual(SaveOutputSnapshot.SignatureUnknown,
                           SaveOutputSnapshot.ClassifySignature(junk, junk.Length), "teks ditolak")

            Dim tiny As Byte() = New Byte() {&H52, &H49}
            Check.AreEqual(SaveOutputSnapshot.SignatureUnknown,
                           SaveOutputSnapshot.ClassifySignature(tiny, 2), "terlalu pendek -> tidak dikenali")
        End Sub)

        ' ------------------------------------------------------------------
        ' B. Output tidak valid ditolak walau ada & tidak kosong
        ' ------------------------------------------------------------------
        suite.Add("output tidak valid (bukan CDR) -> Invalid, bukan sukses", Sub()
            Dim before As SaveOutputSnapshot = SaveOutputSnapshot.DescribeState(False, 0, DateTime.MinValue)
            Dim after As SaveOutputSnapshot = SaveOutputSnapshot.DescribeState(
                True, 4096, Utc(1), "fingerprint-a", SaveOutputSnapshot.SignatureUnknown)

            Dim verification As SaveOutputVerification =
                SaveOutputVerification.Verify(before, after, SaveOutputMode.NewFile, True)
            Check.AreEqual(SaveOutputStatus.Invalid, verification.Status, "status Invalid")
            Check.IsFalse(verification.IsVerified, "tidak dianggap sukses")
            Check.AreEqual("TIDAK VALID", verification.StatusLabel, "label laporan")
            Check.Contains(verification.Reason, "bukan berkas CorelDRAW", "alasan menyebut bentuk berkas")
        End Sub)

        suite.Add("output tidak valid + mode timpa -> Invalid (file sumber rusak)", Sub()
            Dim before As SaveOutputSnapshot = SaveOutputSnapshot.DescribeState(
                True, 4096, Utc(0), "fingerprint-lama", SaveOutputSnapshot.SignatureRiffCdr)
            Dim after As SaveOutputSnapshot = SaveOutputSnapshot.DescribeState(
                True, 100, Utc(5), "fingerprint-baru", SaveOutputSnapshot.SignatureUnknown)

            Dim verification As SaveOutputVerification =
                SaveOutputVerification.Verify(before, after, SaveOutputMode.OverwriteInPlace, True)
            Check.AreEqual(SaveOutputStatus.Invalid, verification.Status, "Invalid walau isi berubah")
            Check.IsFalse(verification.IsVerified, "bukan sukses")
        End Sub)

        suite.Add("bentuk tidak diperiksa bila tidak diminta (kompatibilitas perilaku lama)", Sub()
            Dim before As SaveOutputSnapshot = SaveOutputSnapshot.DescribeState(False, 0, DateTime.MinValue)
            Dim after As SaveOutputSnapshot = SaveOutputSnapshot.DescribeState(
                True, 100, Utc(1), "fp", SaveOutputSnapshot.SignatureUnknown)
            Dim verification As SaveOutputVerification = SaveOutputVerification.Verify(before, after)
            Check.AreEqual(SaveOutputStatus.VerifiedNew, verification.Status,
                           "tanpa requireCorelSignature, aturan lama tetap berlaku")
        End Sub)

        suite.Add("isi tidak terbaca: bentuk tidak dipastikan, tapi sukses tidak diklaim lebih", Sub()
            ' Snapshot tanpa ContentRead (mis. file masih dikunci CorelDRAW).
            Dim before As SaveOutputSnapshot = SaveOutputSnapshot.DescribeState(False, 0, DateTime.MinValue)
            Dim after As SaveOutputSnapshot = SaveOutputSnapshot.DescribeState(True, 2048, Utc(2))
            Dim verification As SaveOutputVerification =
                SaveOutputVerification.Verify(before, after, SaveOutputMode.NewFile, True)
            Check.AreEqual(SaveOutputStatus.VerifiedNew, verification.Status, "bukti BARU tetap dipakai")
            Check.Contains(verification.Reason, "Bentuk berkas tidak dapat diperiksa", "keterbatasan dilaporkan")
        End Sub)

        ' ------------------------------------------------------------------
        ' C. Sidik jari isi sebagai bukti perubahan
        ' ------------------------------------------------------------------
        suite.Add("ukuran & waktu sama tetapi isi berbeda -> DIPERBARUI (bukan TIDAK BERUBAH)", Sub()
            Dim before As SaveOutputSnapshot = SaveOutputSnapshot.DescribeState(
                True, 8192, Utc(3), "fp-sebelum", SaveOutputSnapshot.SignatureRiffCdr)
            Dim after As SaveOutputSnapshot = SaveOutputSnapshot.DescribeState(
                True, 8192, Utc(3), "fp-sesudah", SaveOutputSnapshot.SignatureRiffCdr)

            Dim verification As SaveOutputVerification =
                SaveOutputVerification.Verify(before, after, SaveOutputMode.OverwriteInPlace, True)
            Check.AreEqual(SaveOutputStatus.VerifiedChanged, verification.Status, "VerifiedChanged dari isi file")
            Check.Contains(verification.Reason, "isi file berubah", "alasan menyebut dasar bukti")
        End Sub)

        suite.Add("ukuran, waktu, DAN isi sama -> TIDAK BERUBAH (file lama bukan bukti)", Sub()
            Dim before As SaveOutputSnapshot = SaveOutputSnapshot.DescribeState(
                True, 8192, Utc(3), "fp-sama", SaveOutputSnapshot.SignatureRiffCdr)
            Dim after As SaveOutputSnapshot = SaveOutputSnapshot.DescribeState(
                True, 8192, Utc(3), "fp-sama", SaveOutputSnapshot.SignatureRiffCdr)

            Dim verification As SaveOutputVerification =
                SaveOutputVerification.Verify(before, after, SaveOutputMode.OverwriteInPlace, True)
            Check.AreEqual(SaveOutputStatus.Unchanged, verification.Status, "Unchanged")
            Check.Contains(verification.Reason, "isi file sama", "alasan menyebut isi dibandingkan")
        End Sub)

        ' ------------------------------------------------------------------
        ' D. Snapshot nyata: bentuk + sidik jari + kegagalan yang jujur
        ' ------------------------------------------------------------------
        suite.Add("Capture membaca bentuk & sidik jari file nyata", Sub()
            Dim folder As String = NewTempFolder("capture")
            Try
                Dim validPath As String = Path.Combine(folder, "valid.cdr")
                File.WriteAllBytes(validPath, FakeDocument.FakeCdrBytes(False, False))
                Dim snapshot As SaveOutputSnapshot = SaveOutputSnapshot.Capture(validPath, True)
                Check.IsTrue(snapshot.Exists, "file ada")
                Check.IsTrue(snapshot.ContentRead, "isi terbaca")
                Check.AreEqual(SaveOutputSnapshot.SignatureRiffCdr, snapshot.SignatureKind, "bentuk RIFF/CDR")
                Check.IsTrue(snapshot.Fingerprint.Length >= 32, "sidik jari dihitung")

                Dim invalidPath As String = Path.Combine(folder, "invalid.cdr")
                File.WriteAllBytes(invalidPath, FakeDocument.FakeCdrBytes(True, False))
                Dim invalidSnapshot As SaveOutputSnapshot = SaveOutputSnapshot.Capture(invalidPath, True)
                Check.AreEqual(SaveOutputSnapshot.SignatureUnknown, invalidSnapshot.SignatureKind,
                               "isi tidak valid dikenali")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)

        suite.Add("sidik jari membedakan dua file berukuran sama", Sub()
            Dim folder As String = NewTempFolder("fingerprint")
            Try
                Dim first As String = Path.Combine(folder, "a.cdr")
                Dim second As String = Path.Combine(folder, "b.cdr")
                Dim firstBytes As Byte() = FakeDocument.FakeCdrBytes(False, False)
                Dim secondBytes As Byte() = CType(firstBytes.Clone(), Byte())
                secondBytes(secondBytes.Length - 1) = secondBytes(secondBytes.Length - 1) Xor &HFF
                File.WriteAllBytes(first, firstBytes)
                File.WriteAllBytes(second, secondBytes)

                Dim a As SaveOutputSnapshot = SaveOutputSnapshot.Capture(first, True)
                Dim b As SaveOutputSnapshot = SaveOutputSnapshot.Capture(second, True)
                Check.AreEqual(a.Length, b.Length, "panjang sama")
                Check.IsTrue(a.Fingerprint <> b.Fingerprint, "sidik jari berbeda")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)

        suite.Add("Capture pada path kosong / folder -> kegagalan dilaporkan, bukan dilempar", Sub()
            Dim empty As SaveOutputSnapshot = SaveOutputSnapshot.Capture("", True)
            Check.AreEqual("path output kosong", empty.CaptureError, "path kosong")
            Check.IsFalse(empty.IsValid, "tidak valid")

            Dim folderPath As String = NewTempFolder("asfile")
            Try
                ' Folder bukan berkas: FileInfo.Exists = False, jadi hasilnya
                ' "tidak ada" (bukan exception, bukan pula klaim ada).
                Dim asFolder As SaveOutputSnapshot = SaveOutputSnapshot.Capture(folderPath, True)
                Check.IsFalse(asFolder.Exists, "folder tidak dianggap berkas")
                Check.Contains(asFolder.Describe(), "tidak ada", "keadaan dilaporkan apa adanya")
            Finally
                CleanupFolder(folderPath)
            End Try
        End Sub)

        ' ------------------------------------------------------------------
        ' E. Engine: mode file baru
        ' ------------------------------------------------------------------
        suite.Add("mode file baru: SaveAs menulis isi tidak valid -> GAGAL (TIDAK VALID)", Sub()
            Dim folder As String = NewTempFolder("newinvalid")
            Try
                File.WriteAllText(Path.Combine(folder, "a.cdr"), "x")
                Dim app As New FakeCorelApplication()
                app.Configure = Sub(doc As FakeDocument)
                                      doc.SaveWritesInvalidContent = True
                                  End Sub
                Dim settings As WatermarkSettings = NewTextSettings()
                settings.WriteCsvLog = True
                settings.CsvLogPath = Path.Combine(folder, "log.csv")

                Dim result As WatermarkResult = New WatermarkEngine().ApplyToFolder(app, folder, settings, False, False)
                Check.AreEqual(0, result.FilesOk, "0 file OK")
                Check.AreEqual(1, result.FilesFailed, "1 GAGAL")
                Check.IsTrue(result.FileOutcomes(0).SaveCallSucceeded, "pemanggilan SaveAs tidak melempar")
                Check.AreEqual("TIDAK VALID", result.FileOutcomes(0).SaveVerificationLabel, "label verifikasi")
                Check.Contains(String.Join(" ", result.Failures.ToArray()), "bukan berkas CorelDRAW",
                               "kegagalan menyebut bentuk berkas")
                Check.DoesNotContain(File.ReadAllText(settings.CsvLogPath), ";OK;", "tidak ada OK palsu")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)

        suite.Add("mode file baru: output ZIP (CDR X4+/2022) diterima -> OK", Sub()
            Dim folder As String = NewTempFolder("newzip")
            Try
                File.WriteAllText(Path.Combine(folder, "a.cdr"), "x")
                Dim app As New FakeCorelApplication()
                app.Configure = Sub(doc As FakeDocument)
                                      doc.SaveWritesZipContainer = True
                                  End Sub
                Dim result As WatermarkResult =
                    New WatermarkEngine().ApplyToFolder(app, folder, NewTextSettings(), False, False)

                Check.AreEqual(1, result.FilesOk, "OK: paket ZIP adalah bentuk CDR modern")
                Check.AreEqual("BARU", result.FileOutcomes(0).SaveVerificationLabel, "label BARU")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)

        suite.Add("mode file baru: target lama yang masih ada & tidak berubah -> GAGAL", Sub()
            Dim folder As String = NewTempFolder("newstale")
            Try
                File.WriteAllText(Path.Combine(folder, "a.cdr"), "x")
                Dim stalePath As String = Path.Combine(folder, "a_wm.cdr")
                File.WriteAllBytes(stalePath, FakeDocument.FakeCdrBytes(False, False))

                Dim app As New FakeCorelApplication()
                app.Configure = Sub(doc As FakeDocument)
                                      doc.SaveWritesFile = False
                                  End Sub
                Dim result As WatermarkResult =
                    New WatermarkEngine().ApplyToFolder(app, folder, NewTextSettings(), False, False)

                Check.AreEqual(1, result.FilesFailed, "GAGAL")
                Check.AreEqual("TIDAK BERUBAH", result.FileOutcomes(0).SaveVerificationLabel,
                               "target lama tidak dianggap bukti")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)

        ' ------------------------------------------------------------------
        ' F. Engine: mode timpa
        ' ------------------------------------------------------------------
        suite.Add("mode TIMPA: SaveAs menulis 0 byte -> GAGAL + peringatan sumber rusak", Sub()
            Dim folder As String = NewTempFolder("overwriteempty")
            Try
                Dim sourcePath As String = Path.Combine(folder, "a.cdr")
                File.WriteAllBytes(sourcePath, FakeDocument.FakeCdrBytes(False, False))

                Dim app As New FakeCorelApplication()
                app.Configure = Sub(doc As FakeDocument)
                                      doc.SaveWritesEmptyFile = True
                                  End Sub
                Dim settings As WatermarkSettings = NewTextSettings()
                Dim result As WatermarkResult = New WatermarkEngine().ApplyToFolder(app, folder, settings, False, True)

                Check.AreEqual(1, result.FilesFailed, "GAGAL")
                Check.AreEqual("KOSONG", result.FileOutcomes(0).SaveVerificationLabel, "label KOSONG")
                Check.IsTrue(result.Warnings.Count > 0, "peringatan sumber rusak dilaporkan")
                Check.Contains(String.Join(" ", result.Warnings.ToArray()), "TIDAK LENGKAP",
                               "peringatan menyebut berkas sumber")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)

        suite.Add("mode TIMPA: SaveAs menulis isi tidak valid (sumber tertimpa) -> GAGAL", Sub()
            Dim folder As String = NewTempFolder("overwriteinvalid")
            Try
                Dim sourcePath As String = Path.Combine(folder, "a.cdr")
                File.WriteAllBytes(sourcePath, FakeDocument.FakeCdrBytes(False, False))

                Dim app As New FakeCorelApplication()
                app.Configure = Sub(doc As FakeDocument)
                                      doc.SaveWritesInvalidContent = True
                                  End Sub
                Dim result As WatermarkResult =
                    New WatermarkEngine().ApplyToFolder(app, folder, NewTextSettings(), False, True)

                Check.AreEqual(1, result.FilesFailed, "GAGAL")
                Check.AreEqual("TIDAK VALID", result.FileOutcomes(0).SaveVerificationLabel, "label")
                Check.Contains(String.Join(" ", result.Warnings.ToArray()), "TIDAK LENGKAP",
                               "pengguna diberi tahu sumbernya rusak")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)

        suite.Add("mode TIMPA: SaveAs menulis ulang CDR valid -> OK (DIPERBARUI)", Sub()
            Dim folder As String = NewTempFolder("overwritevalid")
            Try
                Dim sourcePath As String = Path.Combine(folder, "a.cdr")
                File.WriteAllBytes(sourcePath, FakeDocument.FakeCdrBytes(False, False))

                Dim app As New FakeCorelApplication()
                Dim result As WatermarkResult =
                    New WatermarkEngine().ApplyToFolder(app, folder, NewTextSettings(), False, True)

                Check.AreEqual(1, result.FilesOk, "OK")
                Check.AreEqual("DIPERBARUI", result.FileOutcomes(0).SaveVerificationLabel, "label")
                Check.AreEqual(0, result.Warnings.Count, "tidak ada peringatan kerusakan")
                Check.IsTrue(result.FileOutcomes(0).SaveSucceeded, "penyimpanan terbukti")
            Finally
                CleanupFolder(folder)
            End Try
        End Sub)
    End Sub
End Module
