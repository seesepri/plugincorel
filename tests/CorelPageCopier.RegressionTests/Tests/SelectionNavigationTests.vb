Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  SelectionNavigationTests.vb
'  Regression test untuk MatchNavigator & SelectionOutcome
'  (Temuan Auto Select #9 "tanpa melewati/mengulang halaman" dan #10).
' ============================================================================

Imports System
Imports System.Collections.Generic

Public Module SelectionNavigationTests

    Private Function AlwaysExists(pageIndex As Integer) As Boolean
        Return True
    End Function

    Public Sub Register(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("Auto Select — navigasi & seleksi")

        ' --- urutan halaman -------------------------------------------------
        suite.Add("navigasi: urut, tanpa melewati & tanpa mengulang", Sub()
            Dim nav As New MatchNavigator()
            nav.Start(ShapeCategory.Text, True, New Integer() {5, 2, 9, 2})

            Check.AreEqual(3, nav.Count, "duplikat dibuang")
            Check.AreEqual(2, nav.CurrentPage, "mulai dari halaman terkecil")

            Check.AreEqual(NavigationStep.Moved, nav.MoveNext(ShapeCategory.Text, True, AddressOf AlwaysExists), "langkah 1")
            Check.AreEqual(5, nav.CurrentPage, "halaman 5")
            Check.AreEqual(NavigationStep.Moved, nav.MoveNext(ShapeCategory.Text, True, AddressOf AlwaysExists), "langkah 2")
            Check.AreEqual(9, nav.CurrentPage, "halaman 9")
        End Sub)

        suite.Add("navigasi: berhenti di hasil terakhir (tidak kembali ke awal)", Sub()
            Dim nav As New MatchNavigator()
            nav.Start(ShapeCategory.Bitmap, True, New Integer() {1, 2})
            Check.AreEqual(NavigationStep.Moved, nav.MoveNext(ShapeCategory.Bitmap, True, AddressOf AlwaysExists), "ke halaman 2")
            Check.AreEqual(NavigationStep.AtEnd, nav.MoveNext(ShapeCategory.Bitmap, True, AddressOf AlwaysExists), "berhenti")
            Check.AreEqual(2, nav.CurrentPage, "tetap di halaman terakhir")
            Check.IsTrue(nav.IsAtEnd, "IsAtEnd")
            Check.AreEqual("2/2", nav.PositionLabel, "label posisi")
        End Sub)

        suite.Add("navigasi: satu halaman hasil -> langsung AtEnd", Sub()
            Dim nav As New MatchNavigator()
            nav.Start(ShapeCategory.Vector, True, New Integer() {3})
            Check.AreEqual(NavigationStep.AtEnd, nav.MoveNext(ShapeCategory.Vector, True, AddressOf AlwaysExists), "berhenti")
        End Sub)

        suite.Add("navigasi: belum ada pemindaian -> NotStarted", Sub()
            Dim nav As New MatchNavigator()
            Check.AreEqual(NavigationStep.NotStarted, nav.MoveNext(ShapeCategory.Text, True, AddressOf AlwaysExists), "NotStarted")
            Check.IsFalse(nav.HasResults, "tidak ada hasil")
        End Sub)

        ' --- snapshot kategori (bug lama: navigasi pakai kategori saat ini) ----
        suite.Add("kategori diubah setelah pemindaian -> StaleSnapshot (tidak melompat)", Sub()
            Dim nav As New MatchNavigator()
            nav.Start(ShapeCategory.Text, True, New Integer() {2, 4})
            ' Pengguna mengubah ComboBox dari Teks -> Bitmap lalu menekan "hasil berikutnya".
            Check.AreEqual(NavigationStep.StaleSnapshot,
                           nav.MoveNext(ShapeCategory.Bitmap, True, AddressOf AlwaysExists), "ditolak")
            Check.AreEqual(2, nav.CurrentPage, "posisi tidak dimajukan")
            Check.IsTrue(nav.MatchesSnapshot(ShapeCategory.Text, True), "snapshot tetap Teks")
        End Sub)

        suite.Add("cakupan diubah setelah pemindaian -> StaleSnapshot", Sub()
            Dim nav As New MatchNavigator()
            nav.Start(ShapeCategory.Text, True, New Integer() {2})
            Check.AreEqual(NavigationStep.StaleSnapshot,
                           nav.MoveNext(ShapeCategory.Text, False, AddressOf AlwaysExists), "cakupan berbeda")
        End Sub)

        suite.Add("snapshot dipertahankan untuk pemakaian ulang UI", Sub()
            Dim nav As New MatchNavigator()
            nav.Start(ShapeCategory.Vector, True, New Integer() {7})
            Check.AreEqual(ShapeCategory.Vector, nav.Category, "kategori snapshot")
            Check.IsTrue(nav.AllPages, "cakupan snapshot")
            Check.Contains(nav.StatusText(), "halaman 7", "status text")
            Check.Contains(nav.StatusText(), "Objek vektor", "status memuat kategori")
        End Sub)

        ' --- halaman hilang (dokumen diedit) ----------------------------------
        suite.Add("halaman hasil hilang -> PageMissing, posisi TIDAK dimajukan", Sub()
            Dim nav As New MatchNavigator()
            nav.Start(ShapeCategory.Text, True, New Integer() {2, 5})

            Dim exists As Func(Of Integer, Boolean) = Function(p As Integer) p <> 5
            Check.AreEqual(NavigationStep.PageMissing, nav.MoveNext(ShapeCategory.Text, True, exists), "halaman 5 hilang")
            Check.AreEqual(2, nav.CurrentPage, "tetap di halaman 2 (tidak melompat)")
        End Sub)

        suite.Add("Reset membatalkan sesi (dipakai saat pemindaian gagal)", Sub()
            Dim nav As New MatchNavigator()
            nav.Start(ShapeCategory.Text, True, New Integer() {1, 2})
            nav.Reset()
            Check.IsFalse(nav.HasResults, "tidak ada hasil setelah reset")
            Check.AreEqual(0, nav.Count, "daftar kosong")
            Check.AreEqual(NavigationStep.NotStarted, nav.MoveNext(ShapeCategory.Text, True, AddressOf AlwaysExists), "NotStarted")
        End Sub)

        ' --- SelectionOutcome -------------------------------------------------
        suite.Add("seleksi terverifikasi bila jumlah COM == target", Sub()
            Dim outcome As New SelectionOutcome()
            outcome.ExpectedCount = 3
            outcome.AddedCalls = 3
            outcome.ActualCount = 3
            outcome.ActualCountKnown = True
            Check.IsTrue(outcome.SelectionVerified, "terverifikasi")
            Check.AreEqual(3, outcome.SelectionVerifiedCount, "angka terpilih")
            Check.Contains(outcome.Describe(), "Seleksi terverifikasi", "laporan")
        End Sub)

        suite.Add("AddToSelection tidak berefek -> TIDAK dianggap terverifikasi", Sub()
            Dim outcome As New SelectionOutcome()
            outcome.ExpectedCount = 4
            outcome.AddedCalls = 4
            outcome.ActualCount = 1     ' mis. 3 objek di layer terkunci
            outcome.ActualCountKnown = True
            Check.IsFalse(outcome.SelectionVerified, "tidak terverifikasi")
            Check.AreEqual(1, outcome.SelectionVerifiedCount, "hanya 1 yang boleh diklaim")
            Check.Contains(outcome.Describe(), "BELUM terverifikasi", "laporan jujur")
        End Sub)

        suite.Add("jumlah COM tidak terbaca -> tidak mengklaim terverifikasi", Sub()
            Dim outcome As New SelectionOutcome()
            outcome.ExpectedCount = 2
            outcome.AddedCalls = 2
            outcome.ActualCountKnown = False
            Check.IsFalse(outcome.SelectionVerified, "tidak terverifikasi")
            Check.AreEqual(2, outcome.SelectionVerifiedCount, "fallback ke jumlah panggilan")
        End Sub)

        suite.Add("tidak ada target -> tidak ada operasi (seleksi lama aman)", Sub()
            Dim outcome As SelectionOutcome = SelectionOutcome.NotRun("Tidak ada target seleksi pada halaman 4.")
            Check.IsFalse(outcome.WasAttempted, "tidak dijalankan")
            Check.IsFalse(outcome.SelectionVerified, "tidak terverifikasi")
            Check.AreEqual(0, outcome.SelectionVerifiedCount, "0 objek diklaim")
            Check.Contains(outcome.Describe(), "tidak dijalankan", "laporan")
        End Sub)
    End Sub
End Module
