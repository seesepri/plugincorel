Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  BatchLogFormatTests.vb
'  Fokus D — format CSV: aman untuk delimiter, tanda kutip, dan newline;
'  satu baris log = satu baris fisik; ringkasan konsisten dengan jumlah file.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.IO

Public Module BatchLogFormatTests

    ''' <summary>Parser CSV sederhana (pemisah ';', kutip ganda) untuk verifikasi bolak-balik.</summary>
    Private Function ParseCells(line As String) As List(Of String)
        Dim cells As New List(Of String)()
        Dim sb As New System.Text.StringBuilder()
        Dim inQuotes As Boolean = False
        Dim i As Integer = 0
        While i < line.Length
            Dim c As Char = line(i)
            If inQuotes Then
                If c = """"c Then
                    If i + 1 < line.Length AndAlso line(i + 1) = """"c Then
                        sb.Append(""""c)
                        i += 1
                    Else
                        inQuotes = False
                    End If
                Else
                    sb.Append(c)
                End If
            Else
                If c = """"c Then
                    inQuotes = True
                ElseIf c = ";"c Then
                    cells.Add(sb.ToString())
                    sb.Length = 0
                Else
                    sb.Append(c)
                End If
            End If
            i += 1
        End While
        cells.Add(sb.ToString())
        Return cells
    End Function

    Public Sub Register(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("Watermark — format log CSV")

        suite.Add("nama file dengan pemisah & kutip tetap satu sel", Sub()
            Dim entry As New BatchLogEntry()
            entry.FileName = "order;012 ""final"".cdr"
            entry.Status = "OK"
            entry.Message = "selesai"
            Dim line As String = BatchLog.BuildLine(entry)
            Dim cells As List(Of String) = ParseCells(line)
            Check.AreEqual(7, cells.Count, "jumlah kolom tetap 7")
            Check.AreEqual("order;012 ""final"".cdr", cells(1), "nama file utuh")
            Check.AreEqual("OK", cells(4), "status")
        End Sub)

        suite.Add("pesan dengan newline tidak memecah baris CSV", Sub()
            Dim entry As New BatchLogEntry()
            entry.FileName = "a.cdr"
            entry.Status = "GAGAL"
            entry.Message = "baris pertama" & Environment.NewLine & "baris kedua"
            Dim line As String = BatchLog.BuildLine(entry)
            Check.IsFalse(line.Contains(ChrW(13)), "tidak ada CR")
            Check.IsFalse(line.Contains(ChrW(10)), "tidak ada LF")
            Dim cells As List(Of String) = ParseCells(line)
            Check.AreEqual(7, cells.Count, "kolom tidak bergeser")
            Check.Contains(cells(5), "baris pertama", "isi pesan dipertahankan")
            Check.Contains(cells(5), "baris kedua", "isi pesan dipertahankan")
        End Sub)

        suite.Add("tab & karakter kontrol dibersihkan", Sub()
            Dim entry As New BatchLogEntry()
            entry.FileName = "a" & ChrW(9) & "b.cdr"
            entry.Status = "OK"
            entry.Message = "x" & ChrW(0) & "y"
            Dim line As String = BatchLog.BuildLine(entry)
            Check.IsFalse(line.Contains(ChrW(9)), "tanpa tab")
            Check.IsFalse(line.Contains(ChrW(0)), "tanpa NUL")
        End Sub)

        suite.Add("kolom waktu, halaman, shape, durasi terisi", Sub()
            Dim entry As New BatchLogEntry()
            entry.FileName = "a.cdr"
            entry.Pages = 3
            entry.Shapes = 5
            entry.Status = "OK"
            entry.DurationSeconds = 2.5
            Dim cells As List(Of String) = ParseCells(BatchLog.BuildLine(entry))
            Check.AreEqual(BatchLog.Header().Split(";"c).Length, cells.Count, "jumlah kolom = header")
            Check.Contains(cells(0), "-", "timestamp ada")
            Check.AreEqual("3", cells(2), "halaman")
            Check.AreEqual("5", cells(3), "shape")
            Check.AreEqual("2.5", cells(6), "durasi")
        End Sub)

        suite.Add("ringkasan memisahkan OK / SEBAGIAN / GAGAL / DIBATALKAN", Sub()
            Dim entry As BatchLogEntry = BatchLog.SummaryEntry("setelan", 10, 6, 2, 1, 1, 12.5)
            Check.AreEqual("RINGKASAN", entry.Status, "status ringkasan")
            Check.Contains(entry.Message, "OK 6", "OK")
            Check.Contains(entry.Message, "SEBAGIAN 2", "SEBAGIAN")
            Check.Contains(entry.Message, "GAGAL 1", "GAGAL")
            Check.Contains(entry.Message, "DIBATALKAN 1", "DIBATALKAN")
            Check.AreEqual(10, entry.Pages, "jumlah file pada kolom halaman")
        End Sub)

        suite.Add("jumlah status pada ringkasan = jumlah file sebenarnya", Sub()
            Dim tracker As New BatchRunTracker(10)
            For i As Integer = 1 To 6
                Dim ok As New BatchFileOutcome("a" & i & ".cdr")
                ok.ApplyAttempted = True
                ok.PagesIntended = 1
                ok.PagesApplied = 1
                ok.SaveAttempted = True
                ok.SaveSucceeded = True
                tracker.RecordCompleted(ok)
            Next
            For i As Integer = 1 To 2
                Dim partlyDone As New BatchFileOutcome("p" & i & ".cdr")
                partlyDone.ApplyAttempted = True
                partlyDone.PagesIntended = 2
                partlyDone.PagesApplied = 1
                partlyDone.SaveAttempted = True
                partlyDone.SaveSucceeded = True
                tracker.RecordCompleted(partlyDone)
            Next
            Dim failed As New BatchFileOutcome("g.cdr")
            failed.FatalError = "tidak dapat dibuka"
            tracker.RecordCompleted(failed)
            tracker.AddCancelled(1)

            Check.AreEqual(6, tracker.OkCount, "OK")
            Check.AreEqual(2, tracker.PartialCount, "SEBAGIAN")
            Check.AreEqual(1, tracker.FailedCount, "GAGAL")
            Check.AreEqual(1, tracker.CancelledCount, "DIBATALKAN")
            Check.AreEqual(0, tracker.Remaining, "tidak ada sisa")
            Check.IsTrue(tracker.IsConsistent(), "pembukuan konsisten")
            Check.AreEqual("", tracker.ConsistencyMessage(), "tidak ada peringatan konsistensi")
        End Sub)

        suite.Add("pembukuan tidak lengkap terdeteksi (file hilang dari laporan)", Sub()
            Dim tracker As New BatchRunTracker(3)
            Dim ok As New BatchFileOutcome("a.cdr")
            tracker.RecordCompleted(ok)
            Check.IsFalse(tracker.IsConsistent(), "tidak konsisten")
            Check.AreEqual(2, tracker.Remaining, "2 file belum punya status")
            Check.Contains(tracker.ConsistencyMessage(), "tidak konsisten", "pesan")
        End Sub)

        suite.Add("outcome DIBATALKAN tidak dihitung sebagai GAGAL", Sub()
            Dim outcome As BatchFileOutcome = BatchRunTracker.BuildCancelledOutcome("a.cdr", "dibatalkan")
            Check.AreEqual(BatchFileStatus.Cancelled, outcome.Status, "status")
            Check.AreEqual("DIBATALKAN", outcome.StatusLabel(), "label")
            Check.Contains(outcome.MessageText(), "tidak diproses", "pesan")
            Check.DoesNotContain(outcome.MessageText(), "save tidak dijalankan", "tidak menyebut save")
        End Sub)

        suite.Add("status file: PDF tidak valid menurunkan OK menjadi SEBAGIAN", Sub()
            Dim outcome As New BatchFileOutcome("a.cdr")
            outcome.ApplyAttempted = True
            outcome.PagesIntended = 1
            outcome.PagesApplied = 1
            outcome.SaveAttempted = True
            outcome.SaveSucceeded = True
            Check.AreEqual(BatchFileStatus.Ok, outcome.Status, "OK sebelum PDF")
            outcome.ProofPdfInvalid = True
            Check.AreEqual(BatchFileStatus.PartialSuccess, outcome.Status, "SEBAGIAN setelah PDF invalid")
            Check.Contains(outcome.MessageText(), "PDF proof TIDAK VALID", "pesan")
        End Sub)

        suite.Add("file log ditulis dengan BOM UTF-8 & header sekali", Sub()
            Dim folder As String = Path.Combine(Path.GetTempPath(), "cplog-" & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(folder)
            Dim csvPath As String = Path.Combine(folder, "log.csv")
            Try
                Dim entry As New BatchLogEntry()
                entry.FileName = "a.cdr"
                entry.Status = "OK"
                Check.IsTrue(BatchLog.AppendRow(csvPath, entry), "baris pertama")
                Check.IsTrue(BatchLog.AppendRow(csvPath, entry), "baris kedua")
                Dim lines As String() = File.ReadAllLines(csvPath)
                Check.AreEqual(3, lines.Length, "header + 2 baris")
                Check.AreEqual(BatchLog.Header(), lines(0), "header sekali di awal")
                Dim bytes As Byte() = File.ReadAllBytes(csvPath)
                Check.AreEqual(239, CInt(bytes(0)), "BOM UTF-8")
            Finally
                Try
                    Directory.Delete(folder, True)
                Catch
                End Try
            End Try
        End Sub)
    End Sub
End Module
