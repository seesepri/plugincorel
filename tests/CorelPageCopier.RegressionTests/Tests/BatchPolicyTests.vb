Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  BatchPolicyTests.vb
'  Regression test Temuan D (status batch) & E (pencegahan pemrosesan ulang,
'  benturan nama, mode timpa).
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.IO

Public Module BatchPolicyTests

    Public Sub Register(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("Kebijakan file batch (Temuan D/E)")

        ' --- akhiran output -------------------------------------------------
        suite.Add("NormalizeOutputSuffix default & sanitasi", Sub()
            Check.AreEqual("_wm", BatchFilePolicy.NormalizeOutputSuffix(Nothing), "default")
            Check.AreEqual("_wm", BatchFilePolicy.NormalizeOutputSuffix(""), "kosong -> default")
            Check.AreEqual("_out", BatchFilePolicy.NormalizeOutputSuffix("out"), "tanpa garis bawah diberi awalan")
            Check.AreEqual("_v2", BatchFilePolicy.NormalizeOutputSuffix("_v2.cdr"), "ekstensi dibuang")
            Check.AreEqual("_wm", BatchFilePolicy.NormalizeOutputSuffix("///"), "karakter ilegal dibuang")
        End Sub)

        suite.Add("IsGeneratedOutput mendeteksi output _wm (case-insensitive)", Sub()
            Check.IsTrue(BatchFilePolicy.IsGeneratedOutput("C:\a\desain_wm.cdr", "_wm"), "_wm terdeteksi")
            Check.IsTrue(BatchFilePolicy.IsGeneratedOutput("C:\a\Desain_WM.CDR", "_wm"), "huruf besar")
            Check.IsFalse(BatchFilePolicy.IsGeneratedOutput("C:\a\desain.cdr", "_wm"), "file asli bukan output")
            Check.IsFalse(BatchFilePolicy.IsGeneratedOutput("C:\a\desain_wm.pdf", "_wm"), "pdf diabaikan")
        End Sub)

        suite.Add("BuildOutputPath tidak pernah menimpa sumber", Sub()
            Dim source As String = Path.Combine("C:\batch", "desain.cdr")
            Dim output As String = BatchFilePolicy.BuildOutputPath(source, "_wm")
            Check.AreEqual(Path.Combine("C:\batch", "desain_wm.cdr"), output, "path output")
            Check.IsFalse(BatchFilePolicy.SamePath(source, output), "berbeda dari sumber")
            Check.AreEqual("", BatchFilePolicy.ValidateNewOutputPath(source, output), "output valid")
            Check.IsTrue(BatchFilePolicy.ValidateNewOutputPath(source, source).Length > 0,
                         "output == sumber ditolak")
        End Sub)

        ' --- filter sumber --------------------------------------------------
        suite.Add("SelectSources membuang output _wm dari daftar sumber", Sub()
            Dim candidates As New List(Of String)()
            candidates.Add(Path.Combine("C:\batch", "a.cdr"))
            candidates.Add(Path.Combine("C:\batch", "b_wm.cdr"))
            candidates.Add(Path.Combine("C:\batch", "c_wm_wm.cdr"))
            candidates.Add(Path.Combine("C:\batch", "catatan.txt"))

            Dim selection As BatchSourceSelection = BatchFilePolicy.SelectSources(candidates, "_wm")
            Check.AreEqual(1, selection.Sources.Count, "hanya 1 sumber")
            Check.AreEqual(Path.Combine("C:\batch", "a.cdr"), selection.Sources(0), "sumber a.cdr")
            Check.AreEqual(3, selection.Skipped.Count, "3 file dilewati")
            Check.Contains(selection.Skipped(0).Reason + selection.Skipped(1).Reason + selection.Skipped(2).Reason,
                           "output batch sebelumnya", "alasan output lama dicatat")
        End Sub)

        suite.Add("SelectSources: semua output lama -> tidak ada yang diproses", Sub()
            Dim candidates As New List(Of String)()
            candidates.Add(Path.Combine("C:\batch", "a_wm.cdr"))
            Dim selection As BatchSourceSelection = BatchFilePolicy.SelectSources(candidates, "_wm")
            Check.AreEqual(0, selection.Sources.Count, "tidak ada sumber")
            Check.AreEqual(1, selection.Skipped.Count, "satu dilewati")
        End Sub)

        ' --- benturan nama --------------------------------------------------
        suite.Add("ResolveUniqueOutputPath menambah akhiran angka", Sub()
            Dim taken As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            taken.Add(Path.Combine("C:\batch", "a_wm.cdr"))
            Dim resolved As String = BatchFilePolicy.ResolveUniqueOutputPath(
                Path.Combine("C:\batch", "a_wm.cdr"),
                Function(candidate As String) taken.Contains(candidate))
            Check.AreEqual(Path.Combine("C:\batch", "a_wm_2.cdr"), resolved, "nama unik")
        End Sub)

        suite.Add("ResolveUniqueOutputPath memakai kandidat bila bebas", Sub()
            Dim candidate As String = Path.Combine("C:\batch", "a_wm.cdr")
            Check.AreEqual(candidate, BatchFilePolicy.ResolveUniqueOutputPath(candidate, Function(p As String) False), "tetap")
        End Sub)

        ' --- validasi mode timpa -------------------------------------------
        suite.Add("ValidateOverwriteTarget menolak path di luar folder", Sub()
            Dim errorText As String = BatchFilePolicy.ValidateOverwriteTarget(
                Path.Combine("C:\lain", "a.cdr"), "C:\batch", False)
            Check.IsTrue(errorText.Length > 0, "ditolak")
            Check.Contains(errorText, "luar folder", "alasan")
        End Sub)

        suite.Add("IsUnderFolder membedakan subfolder", Sub()
            Check.IsTrue(BatchFilePolicy.IsUnderFolder(Path.Combine("C:\batch", "a.cdr"), "C:\batch", False), "file langsung")
            Check.IsFalse(BatchFilePolicy.IsUnderFolder(Path.Combine("C:\batch", "sub", "a.cdr"), "C:\batch", False), "subfolder ditolak")
            Check.IsTrue(BatchFilePolicy.IsUnderFolder(Path.Combine("C:\batch", "sub", "a.cdr"), "C:\batch", True), "subfolder diizinkan")
            Check.IsFalse(BatchFilePolicy.IsUnderFolder("C:\batch2\a.cdr", "C:\batch", True), "prefix palsu ditolak")
        End Sub)

        ' --- aturan status --------------------------------------------------
        suite.Add("status OK hanya bila semua langkah wajib berhasil", Sub()
            Check.AreEqual(BatchFileStatus.Ok,
                           BatchStatusRules.Evaluate(True, 3, 3, 0, True, True, False), "semua sukses")
            Check.AreEqual(BatchFileStatus.Failed,
                           BatchStatusRules.Evaluate(True, 3, 3, 0, True, False, False), "save gagal -> GAGAL")
            Check.AreEqual(BatchFileStatus.Failed,
                           BatchStatusRules.Evaluate(True, 3, 0, 0, True, True, False), "tidak ada halaman -> GAGAL")
            Check.AreEqual(BatchFileStatus.Failed,
                           BatchStatusRules.Evaluate(True, 3, 3, 0, False, False, False), "tanpa save -> GAGAL")
            Check.AreEqual(BatchFileStatus.Failed,
                           BatchStatusRules.Evaluate(False, 3, 3, 0, True, True, False), "tidak diproses -> GAGAL")
            Check.AreEqual(BatchFileStatus.PartialSuccess,
                           BatchStatusRules.Evaluate(True, 3, 2, 1, True, True, False), "1 halaman gagal -> SEBAGIAN")
            Check.AreEqual(BatchFileStatus.PartialSuccess,
                           BatchStatusRules.Evaluate(True, 3, 2, 0, True, True, False), "kurang 1 halaman -> SEBAGIAN")
            Check.AreEqual(BatchFileStatus.Failed,
                           BatchStatusRules.Evaluate(True, 3, 3, 0, True, True, True), "fatal -> GAGAL")
        End Sub)

        suite.Add("label status dipakai bersama (OK/SEBAGIAN/GAGAL)", Sub()
            Check.AreEqual("OK", BatchStatusRules.Label(BatchFileStatus.Ok), "OK")
            Check.AreEqual("SEBAGIAN", BatchStatusRules.Label(BatchFileStatus.PartialSuccess), "SEBAGIAN")
            Check.AreEqual("GAGAL", BatchStatusRules.Label(BatchFileStatus.Failed), "GAGAL")
        End Sub)

        suite.Add("BatchFileOutcome: save gagal tidak dihitung sukses", Sub()
            Dim outcome As New BatchFileOutcome("C:\batch\a.cdr")
            outcome.ApplyAttempted = True
            outcome.PagesIntended = 2
            outcome.PagesApplied = 2
            outcome.SaveAttempted = True
            outcome.SaveSucceeded = False
            Check.AreEqual(BatchFileStatus.Failed, outcome.Status, "status gagal")
            Check.Contains(outcome.MessageText(), "SAVE GAGAL", "pesan menyebut save gagal")
        End Sub)

        suite.Add("BatchStatusRules.CountStatus menghitung dari satu sumber", Sub()
            Dim outcomes As New List(Of BatchFileOutcome)()
            Dim ok As New BatchFileOutcome("a.cdr")
            ok.ApplyAttempted = True : ok.PagesIntended = 1 : ok.PagesApplied = 1
            ok.SaveAttempted = True : ok.SaveSucceeded = True
            Dim partial_ As New BatchFileOutcome("b.cdr")
            partial_.ApplyAttempted = True : partial_.PagesIntended = 2 : partial_.PagesApplied = 1
            partial_.PagesFailed = 1 : partial_.SaveAttempted = True : partial_.SaveSucceeded = True
            Dim bad As New BatchFileOutcome("c.cdr")
            bad.FatalError = "gagal dibuka"
            outcomes.Add(ok)
            outcomes.Add(partial_)
            outcomes.Add(bad)
            Check.AreEqual(1, BatchStatusRules.CountStatus(outcomes, BatchFileStatus.Ok), "1 OK")
            Check.AreEqual(1, BatchStatusRules.CountStatus(outcomes, BatchFileStatus.PartialSuccess), "1 SEBAGIAN")
            Check.AreEqual(1, BatchStatusRules.CountStatus(outcomes, BatchFileStatus.Failed), "1 GAGAL")
        End Sub)

        suite.Add("EvaluateApplyOnly (operasi satu dokumen)", Sub()
            Check.AreEqual(BatchFileStatus.Ok, BatchStatusRules.EvaluateApplyOnly(True, 2, 2, 0, False), "sukses")
            Check.AreEqual(BatchFileStatus.PartialSuccess, BatchStatusRules.EvaluateApplyOnly(True, 2, 1, 1, False), "sebagian")
            Check.AreEqual(BatchFileStatus.Failed, BatchStatusRules.EvaluateApplyOnly(True, 2, 0, 0, False), "tidak ada halaman")
            Check.AreEqual(BatchFileStatus.Failed, BatchStatusRules.EvaluateApplyOnly(True, 2, 2, 0, True), "ada error")
        End Sub)

        suite.Add("Ringkasan CSV memisahkan SEBAGIAN", Sub()
            Dim entry As BatchLogEntry = BatchLog.SummaryEntry("uji", 3, 1, 1, 1, 2.0)
            Check.Contains(entry.Message, "OK 1", "jumlah OK")
            Check.Contains(entry.Message, "SEBAGIAN 1", "jumlah SEBAGIAN")
            Check.Contains(entry.Message, "GAGAL 1", "jumlah GAGAL")
        End Sub)
    End Sub
End Module
