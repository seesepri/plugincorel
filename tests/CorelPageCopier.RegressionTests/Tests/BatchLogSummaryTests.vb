Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  BatchLogSummaryTests.vb
'  Review follow-up KEDUA pada PR #2 — butir 3.
'
'  Sebelumnya `entry.Message` dirakit di beberapa tempat untuk baris yang sama,
'  sehingga teks ringkasan bisa tertimpa/berlipat. Sekarang satu fungsi
'  (BatchLog.BuildSummaryMessage) menjadi satu-satunya perakit, dan test ini
'  mengunci sifat itu: teks muncul TEPAT SEKALI di baris CSV, dan kedua overload
'  SummaryEntry menghasilkan pesan yang identik.
' ============================================================================

Imports System
Imports System.IO

Public Module BatchLogSummaryTests

    Private Function CountOccurrences(haystack As String, needle As String) As Integer
        Dim count As Integer = 0
        Dim position As Integer = haystack.IndexOf(needle, StringComparison.Ordinal)
        While position >= 0
            count += 1
            position = haystack.IndexOf(needle, position + needle.Length, StringComparison.Ordinal)
        End While
        Return count
    End Function

    Public Sub Register(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("Log CSV: teks ringkasan dirakit satu kali (review kedua)")

        suite.Add("SummaryEntry: pesan sama dengan BuildSummaryMessage (tanpa duplikasi)", Sub()
            Dim entry As BatchLogEntry = BatchLog.SummaryEntry("kualitas 80", 5, 2, 2, 1, 3.5)
            Dim expected As String = BatchLog.BuildSummaryMessage("kualitas 80", 2, 2, 1)
            Check.AreEqual(expected, entry.Message, "pesan dirakit oleh satu fungsi")
            Check.AreEqual(1, CountOccurrences(entry.Message, "SEBAGIAN"), "kata SEBAGIAN muncul sekali")
            Check.AreEqual(1, CountOccurrences(entry.Message, "kualitas 80"), "pengaturan muncul sekali")
            Check.AreEqual("RINGKASAN", entry.Status, "status ringkasan")
        End Sub)

        suite.Add("overload 4-argumen meneruskan partial = 0 (kompatibilitas)", Sub()
            Dim entry As BatchLogEntry = BatchLog.SummaryEntry("uji", 3, 3, 0, 1.0)
            Check.AreEqual(BatchLog.BuildSummaryMessage("uji", 3, 0, 0), entry.Message, "pesan konsisten")
            Check.Contains(entry.Message, "SEBAGIAN 0", "kolom SEBAGIAN tetap ada")
        End Sub)

        suite.Add("baris CSV memuat teks ringkasan tepat sekali", Sub()
            Dim entry As BatchLogEntry = BatchLog.SummaryEntry("kualitas 80", 4, 1, 1, 2, 6.0)
            Dim line As String = BatchLog.BuildLine(entry)
            Check.AreEqual(1, CountOccurrences(line, "OK 1"), "ringkasan tidak berlipat di CSV")
            Check.AreEqual(1, CountOccurrences(line, "GAGAL 2"), "hitungan GAGAL tidak berlipat")
            Check.AreEqual(7, line.Split(";"c).Length, "jumlah kolom tetap 7")
        End Sub)

        suite.Add("teks ringkasan yang memuat pemisah tetap di-escape satu kali", Sub()
            ' Settings bisa memuat ';' (mis. nama preset). Escaping tidak boleh
            ' menghasilkan penggandaan pesan.
            Dim entry As BatchLogEntry = BatchLog.SummaryEntry("preset; khusus", 1, 1, 0, 0, 0.5)
            Dim line As String = BatchLog.BuildLine(entry)
            Dim quote As String = Chr(34)
            Check.Contains(line, quote & "preset; khusus | OK 1, SEBAGIAN 0, GAGAL 0" & quote,
                           "sel dibungkus kutip & pesan utuh")
            Check.AreEqual(1, CountOccurrences(line, "SEBAGIAN 0"), "satu kemunculan saja")
        End Sub)
    End Sub
End Module
