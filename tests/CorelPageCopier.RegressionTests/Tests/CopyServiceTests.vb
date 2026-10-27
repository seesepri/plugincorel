Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  CopyServiceTests.vb
'  Regression test untuk pemindahan pipeline copy dari CopierForm.CopyClickedCore
'  ke CopyService. Memakai FakeCorel (bukan COM) dan prompt palsu.
'
'  Fokus: setiap jalur batal punya alasan eksplisit, urutan operasi tetap
'  (konfirmasi sebelum mengubah dokumen, satuan dipulihkan), dan objek sumber
'  tidak pernah dihapus.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.IO

''' <summary>Prompt palsu untuk CopyService.</summary>
Public Class FakeCopyPrompt
    Implements ICopyPrompt

    Public Property AllowCopy As Boolean = True
    Public Property AllowWithoutUndo As Boolean = True
    Public Property ConfirmCalls As Integer = 0
    Public Property LastPlan As CopyPlan = Nothing
    Public Property UndoQuestionCalls As Integer = 0
    Public Property RestoreFailureReported As Boolean = False
    Public Property LastRestoreError As String = ""

    Public Function ConfirmCopy(plan As CopyPlan) As Boolean Implements ICopyPrompt.ConfirmCopy
        ConfirmCalls += 1
        LastPlan = plan
        Return AllowCopy
    End Function

    Public Function ConfirmContinueWithoutUndoGroup(reason As String) As Boolean Implements ICopyPrompt.ConfirmContinueWithoutUndoGroup
        UndoQuestionCalls += 1
        Return AllowWithoutUndo
    End Function

    Public Sub NotifyUnitRestoreFailed(targetUnitName As String, errorMessage As String) Implements ICopyPrompt.NotifyUnitRestoreFailed
        RestoreFailureReported = True
        LastRestoreError = errorMessage
    End Sub
End Class

Public Module CopyServiceTests

    Private Function NewConnection(doc As FakeDocument, app As FakeCorelApplication,
                                   Optional prompt As ICorelPrompt = Nothing) As CorelConnectionService
        app.ActiveDocumentInstance = doc
        Dim fakePrompt As ICorelPrompt = If(prompt, New FakeCorelPrompt())
        Dim service As New CorelConnectionService(fakePrompt,
            Function(progId As String) DirectCast(app, Object),
            Function(progId As String) DirectCast(Nothing, Object),
            New String() {"CorelDRAW.Application.24"})
        service.Connect()
        service.RememberLoadedIdentity()
        Return service
    End Function

    Private Function NewDocument(pageCount As Integer) As FakeDocument
        Dim doc As New FakeDocument("C:\uji\copy.cdr")
        For i As Integer = 1 To pageCount
            Dim page As New FakePage(i)
            page.EnsureActiveLayer()
            doc.Pages.Add(page)
        Next
        doc.ActivePage = doc.Pages(1)
        Return doc
    End Function

    Private Function AddSourceShape(doc As FakeDocument) As FakeShape
        Dim layer As FakeLayer = doc.Pages(1).EnsureActiveLayer()
        Dim shape As New FakeShape(3, 100, 100, 50, 30)
        shape.Name = "Kotak sumber"
        shape.Layer = layer
        layer.Items.Add(shape)
        doc.Selection.Add(shape)
        Return shape
    End Function

    Private Sub RegisterValidation(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("CopyService — validasi target (murni)")

        suite.Add("halaman di luar rentang & halaman sumber dilewati dengan catatan", Sub()
            Dim result As CopyTargetValidation = CopyTargetValidation.Validate(
                New Integer() {1, 2, 9, 0}, sourcePage:=1, pageCount:=3)
            Check.AreEqual(1, result.ValidTargets.Count, "hanya satu target valid")
            Check.AreEqual(2, result.ValidTargets(0), "halaman 2")
            Check.AreEqual(3, result.Warnings.Count, "tiga peringatan")
            Check.Contains(String.Join(" ", result.Warnings.ToArray()), "tidak valid", "alasan rentang")
            Check.Contains(String.Join(" ", result.Warnings.ToArray()), "halaman sumber", "alasan sumber")
        End Sub)

        suite.Add("duplikat dibuang, urutan dipertahankan", Sub()
            Dim result As CopyTargetValidation = CopyTargetValidation.Validate(
                New Integer() {3, 2, 3, 2}, sourcePage:=1, pageCount:=3)
            Check.AreEqual(2, result.ValidTargets.Count, "dua target unik")
            Check.AreEqual(3, result.ValidTargets(0), "urutan asli dipertahankan")
            Check.AreEqual(2, result.ValidTargets(1), "urutan asli dipertahankan")
        End Sub)
    End Sub

    Public Sub Register(runner As TestRunner)
        RegisterValidation(runner)
        Dim suite As TestSuite = runner.AddSuite("CopyService (arsitektur)")

        suite.Add("tanpa halaman tujuan -> NoTargetPages, tidak menyentuh dokumen", Sub()
            Dim doc As FakeDocument = NewDocument(3)
            Dim app As New FakeCorelApplication()
            Dim prompt As New FakeCopyPrompt()
            Dim service As New CopyService(NewConnection(doc, app), prompt, New NullProgressSink())

            Dim result As CopyServiceResult = service.Run(1, New Integer() {})
            Check.AreEqual(CopyAbortReason.NoTargetPages, result.AbortReason, "alasan")
            Check.IsFalse(result.Executed, "tidak dijalankan")
            Check.AreEqual(0, prompt.ConfirmCalls, "tidak ada konfirmasi")
            Check.AreEqual(0, doc.CommandGroupsBegun.Count, "tidak ada grup undo")
        End Sub)

        suite.Add("semua target tidak valid -> NoValidTargets + detail", Sub()
            Dim doc As FakeDocument = NewDocument(2)
            Dim app As New FakeCorelApplication()
            Dim prompt As New FakeCopyPrompt()
            Dim service As New CopyService(NewConnection(doc, app), prompt, New NullProgressSink())

            Dim result As CopyServiceResult = service.Run(1, New Integer() {1, 99})
            Check.AreEqual(CopyAbortReason.NoValidTargets, result.AbortReason, "alasan")
            Check.Contains(result.AbortDetail, "tidak valid", "detail diberikan ke UI")
            Check.AreEqual(0, prompt.ConfirmCalls, "tidak ada konfirmasi")
        End Sub)

        suite.Add("tidak ada seleksi -> NoSelection", Sub()
            Dim doc As FakeDocument = NewDocument(2)
            Dim app As New FakeCorelApplication()
            Dim prompt As New FakeCopyPrompt()
            Dim service As New CopyService(NewConnection(doc, app), prompt, New NullProgressSink())

            Dim result As CopyServiceResult = service.Run(1, New Integer() {2})
            Check.AreEqual(CopyAbortReason.NoSelection, result.AbortReason, "alasan")
            Check.AreEqual(0, doc.CommandGroupsBegun.Count, "dokumen tidak diubah")
        End Sub)

        suite.Add("pengguna membatalkan konfirmasi -> UserCancelled, dokumen belum diubah", Sub()
            Dim doc As FakeDocument = NewDocument(2)
            AddSourceShape(doc)
            Dim app As New FakeCorelApplication()
            Dim prompt As New FakeCopyPrompt()
            prompt.AllowCopy = False
            Dim service As New CopyService(NewConnection(doc, app), prompt, New NullProgressSink())

            Dim result As CopyServiceResult = service.Run(1, New Integer() {2})
            Check.AreEqual(CopyAbortReason.UserCancelled, result.AbortReason, "alasan")
            Check.AreEqual(1, prompt.ConfirmCalls, "konfirmasi ditanyakan")
            Check.AreEqual(0, doc.UnitSetAttempts.Count, "satuan belum diubah")
            Check.AreEqual(0, doc.CommandGroupsBegun.Count, "grup undo belum dibuka")
        End Sub)

        suite.Add("konfirmasi memuat identitas, halaman, dan satuan", Sub()
            Dim doc As FakeDocument = NewDocument(3)
            AddSourceShape(doc)
            Dim app As New FakeCorelApplication()
            Dim prompt As New FakeCopyPrompt()
            prompt.AllowCopy = False
            Dim service As New CopyService(NewConnection(doc, app), prompt, New NullProgressSink())

            service.Run(1, New Integer() {2, 3})
            Check.IsNotNull(prompt.LastPlan, "plan tersedia")
            Check.AreEqual(1, prompt.LastPlan.SourcePage, "halaman sumber")
            Check.AreEqual(2, prompt.LastPlan.TargetCount, "dua target")
            Check.AreEqual("mm", prompt.LastPlan.UnitName, "satuan")
            Check.Contains(prompt.LastPlan.DocumentIdentity, "copy.cdr", "identitas dokumen")
        End Sub)

        suite.Add("satuan tidak terbaca -> UnitUnreadable, tidak ada perubahan", Sub()
            Dim doc As FakeDocument = NewDocument(2)
            AddSourceShape(doc)
            doc.FailUnitRead = True
            Dim app As New FakeCorelApplication()
            Dim prompt As New FakeCopyPrompt()
            Dim service As New CopyService(NewConnection(doc, app), prompt, New NullProgressSink())

            Dim result As CopyServiceResult = service.Run(1, New Integer() {2})
            Check.AreEqual(CopyAbortReason.UnitUnreadable, result.AbortReason, "alasan")
            Check.Contains(result.AbortDetail, "tidak terbaca", "detail")
            Check.AreEqual(0, prompt.ConfirmCalls, "tidak ada konfirmasi")
        End Sub)

        suite.Add("grup undo gagal & pengguna menolak lanjut -> UndoGroupDeclined", Sub()
            Dim doc As FakeDocument = NewDocument(2)
            AddSourceShape(doc)
            doc.FailBeginCommandGroup = True
            Dim app As New FakeCorelApplication()
            Dim prompt As New FakeCopyPrompt()
            prompt.AllowWithoutUndo = False
            Dim service As New CopyService(NewConnection(doc, app), prompt, New NullProgressSink())

            Dim result As CopyServiceResult = service.Run(1, New Integer() {2})
            Check.AreEqual(CopyAbortReason.UndoGroupDeclined, result.AbortReason, "alasan")
            Check.AreEqual(1, prompt.UndoQuestionCalls, "pengguna ditanya")
            Check.IsFalse(result.Executed, "tidak ada paste")
        End Sub)

        suite.Add("copy sukses: objek sumber utuh & halaman tujuan bertambah", Sub()
            Dim doc As FakeDocument = NewDocument(2)
            Dim source As FakeShape = AddSourceShape(doc)
            Dim app As New FakeCorelApplication()
            Dim prompt As New FakeCopyPrompt()
            Dim sink As New RecordingProgressSink()
            Dim service As New CopyService(NewConnection(doc, app), prompt, sink)

            Dim result As CopyServiceResult = service.Run(1, New Integer() {2})

            Check.IsTrue(result.Executed, "dijalankan")
            Check.AreEqual(CopyAbortReason.None, result.AbortReason, "tidak ada pembatalan")
            Check.AreEqual(1, result.Summary.TotalCount, "satu halaman diproses")
            Check.AreEqual(1, doc.Pages(2).EnsureActiveLayer().Items.Count, "objek masuk ke halaman 2")
            Check.IsFalse(source.Deleted, "objek sumber TIDAK dihapus")
            Check.IsTrue(source.Copied, "objek sumber disalin (bukan dipindah)")
            Check.AreEqual(1, sink.Messages.Count, "satu pesan progres")
            Check.Contains(sink.LastMessage, "Copy progres: 1/1", "format progres dipertahankan")
            Check.AreEqual(1, sink.PumpCount, "DoEvents dipanggil lewat sink")
        End Sub)

        suite.Add("satuan dinormalkan lalu DIPULIHKAN ke nilai asli", Sub()
            Dim doc As FakeDocument = NewDocument(2)
            AddSourceShape(doc)
            doc.Unit = DocumentUnits.Inch
            doc.UnitSetAttempts.Clear()
            Dim app As New FakeCorelApplication()
            Dim service As New CopyService(NewConnection(doc, app), New FakeCopyPrompt(), New NullProgressSink())

            service.Run(1, New Integer() {2})

            Check.AreEqual(2, doc.UnitSetAttempts.Count, "dua kali set: mm lalu inch")
            Check.AreEqual(DocumentUnits.Millimeter, doc.UnitSetAttempts(0), "dinormalkan ke mm")
            Check.AreEqual(DocumentUnits.Inch, doc.UnitSetAttempts(1), "dipulihkan ke inch")
            Check.AreEqual(DocumentUnits.Inch, doc.Unit, "satuan akhir = asli")
        End Sub)

        suite.Add("pemulihan satuan gagal -> dilaporkan ke prompt & peringatan", Sub()
            Dim doc As FakeDocument = NewDocument(2)
            AddSourceShape(doc)
            doc.Unit = DocumentUnits.Inch
            doc.FailUnitRestoreTo = DocumentUnits.Inch
            Dim app As New FakeCorelApplication()
            Dim prompt As New FakeCopyPrompt()
            Dim service As New CopyService(NewConnection(doc, app), prompt, New NullProgressSink())

            Dim result As CopyServiceResult = service.Run(1, New Integer() {2})

            Check.IsTrue(result.RestoreFailed, "ditandai gagal")
            Check.IsTrue(prompt.RestoreFailureReported, "pengguna diberi tahu")
            Check.Contains(String.Join(" ", result.Warnings.ToArray()), "PEMULIHAN SATUAN DOKUMEN GAGAL",
                           "peringatan eksplisit")
        End Sub)

        suite.Add("grup undo ditutup & halaman sumber diaktifkan kembali", Sub()
            Dim doc As FakeDocument = NewDocument(3)
            AddSourceShape(doc)
            Dim app As New FakeCorelApplication()
            Dim service As New CopyService(NewConnection(doc, app), New FakeCopyPrompt(), New NullProgressSink())

            service.Run(1, New Integer() {2, 3})

            Check.AreEqual(1, doc.CommandGroupsBegun.Count, "satu grup undo")
            Check.IsFalse(doc.CommandGroupOpen, "grup undo ditutup")
            Check.IsTrue(doc.Pages(1).Activated, "halaman sumber diaktifkan kembali")
        End Sub)

        suite.Add("exception di satu halaman -> halaman itu Failed, tidak menghentikan laporan", Sub()
            Dim doc As FakeDocument = NewDocument(2)
            AddSourceShape(doc)
            ' Halaman 2 dibuat tidak bisa diakses: Activate melempar.
            Dim broken As New FakePage(2)
            broken.FailActivate = True
            doc.Pages.All.Clear()
            doc.Pages.Add(New FakePage(1))
            doc.Pages(1).EnsureActiveLayer()
            doc.Pages.Add(broken)
            doc.ActivePage = doc.Pages(1)
            doc.Selection.Clear()
            AddSourceShape(doc)

            Dim app As New FakeCorelApplication()
            Dim service As New CopyService(NewConnection(doc, app), New FakeCopyPrompt(), New NullProgressSink())
            Dim result As CopyServiceResult = service.Run(1, New Integer() {2})

            Check.AreEqual(1, result.Summary.FailedCount, "halaman gagal dicatat")
            Check.IsTrue(result.Executed, "operasi tetap dianggap berjalan")
        End Sub)
    End Sub
End Module
