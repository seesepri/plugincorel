Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  CorelConnectionServiceTests.vb
'  Regression test untuk pemindahan logika koneksi dari CopierForm ke
'  CorelConnectionService. Semua cabang diuji tanpa CorelDRAW: pencarian
'  instance (ROT) dan pembuatan instance disuntik sebagai delegate.
'
'  Yang dijaga (Temuan F audit reliabilitas):
'    - instance yang sedang berjalan diutamakan,
'    - instance kedua tidak pernah dibuka tanpa persetujuan,
'    - dokumen dari instance lain tidak dianggap target yang sama.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.IO

''' <summary>Prompt palsu: mencatat pertanyaan & jawaban yang diberikan test.</summary>
Public Class FakeCorelPrompt
    Implements ICorelPrompt

    Public Property AllowNewInstance As Boolean = False
    Public Property OpenInstanceCalls As Integer = 0
    Public Property LastDetail As String = ""
    Public ReadOnly Property Notifications As New List(Of String)()
    Public Property MismatchReported As Boolean = False

    Public Function ConfirmOpenNewInstance(detail As String) As Boolean Implements ICorelPrompt.ConfirmOpenNewInstance
        OpenInstanceCalls += 1
        LastDetail = detail
        Return AllowNewInstance
    End Function

    Public Sub NotifyNewInstanceOpened(hasActiveDocument As Boolean) Implements ICorelPrompt.NotifyNewInstanceOpened
        Notifications.Add(If(hasActiveDocument, "instance baru + dokumen", "instance baru tanpa dokumen"))
    End Sub

    Public Sub NotifyDocumentMismatch(loaded As DocumentIdentity,
                                      current As DocumentIdentity) Implements ICorelPrompt.NotifyDocumentMismatch
        MismatchReported = True
    End Sub
End Class

Public Module CorelConnectionServiceTests

    Private Function NewDocument(name As String, pageCount As Integer) As FakeDocument
        Dim doc As New FakeDocument("C:\uji\" & name & ".cdr")
        doc.Name = name
        For i As Integer = 1 To pageCount
            Dim page As New FakePage(i)
            page.EnsureActiveLayer()
            doc.Pages.Add(page)
        Next
        doc.ActivePage = doc.Pages(1)
        Return doc
    End Function

    Public Sub Register(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("CorelConnectionService (arsitektur)")

        suite.Add("instance berjalan dengan dokumen -> dipakai tanpa bertanya", Sub()
            Dim running As New FakeCorelApplication()
            running.OpenDocument("C:\uji\a.cdr")
            running.ActiveDocumentInstance = running.Opened(0)
            Dim prompt As New FakeCorelPrompt()
            Dim service As New CorelConnectionService(prompt,
                Function(progId As String) DirectCast(running, Object),
                Function(progId As String) DirectCast(Nothing, Object),
                New String() {"CorelDRAW.Application.24"})

            Dim outcome As ConnectionOutcome = service.Connect()

            Check.IsTrue(outcome.Connected, "terhubung")
            Check.IsTrue(outcome.AttachedToRunning, "memakai instance berjalan")
            Check.AreEqual(0, prompt.OpenInstanceCalls, "tidak bertanya membuka instance baru")
            Check.IsTrue(service.Document IsNot Nothing, "dokumen aktif tersedia")
        End Sub)

        suite.Add("tidak ada instance & pengguna menolak -> exception, tidak membuka instance", Sub()
            Dim prompt As New FakeCorelPrompt()
            prompt.AllowNewInstance = False
            Dim factory As New FakeCorelApplication()
            Dim service As New CorelConnectionService(prompt,
                Function(progId As String) DirectCast(Nothing, Object),
                Function(progId As String) DirectCast(factory, Object),
                New String() {"CorelDRAW.Application.24"})

            Dim message As String = ""
            Try
                service.Connect()
            Catch ex As Exception
                message = ex.Message
            End Try
            Check.Contains(message, "dibatalkan oleh pengguna", "alasan penolakan")
            Check.AreEqual(1, prompt.OpenInstanceCalls, "bertanya satu kali")
            Check.AreEqual(0, factory.Opened.Count, "instance baru tidak dipakai")
        End Sub)

        suite.Add("pengguna menyetujui -> instance baru dipakai + diberitahukan", Sub()
            Dim prompt As New FakeCorelPrompt()
            prompt.AllowNewInstance = True
            Dim factory As New FakeCorelApplication()
            Dim service As New CorelConnectionService(prompt,
                Function(progId As String) DirectCast(Nothing, Object),
                Function(progId As String) DirectCast(factory, Object),
                New String() {"CorelDRAW.Application.24"})

            ' Factory membuat aplikasi baru yang punya dokumen aktif.
            factory.OpenDocument("C:\uji\baru.cdr")
            factory.ActiveDocumentInstance = factory.Opened(0)

            Dim outcome As ConnectionOutcome = service.Connect()
            Check.IsTrue(outcome.Connected, "terhubung")
            Check.IsTrue(outcome.CreatedNewInstance, "menandai instance baru")
            Check.AreEqual(1, prompt.Notifications.Count, "pengguna diberi tahu")
            Check.Contains(prompt.Notifications(0), "instance baru", "jenis pemberitahuan")
        End Sub)

        suite.Add("instance baru tanpa dokumen -> exception jelas, instance dilepas", Sub()
            Dim prompt As New FakeCorelPrompt()
            prompt.AllowNewInstance = True
            Dim factory As New FakeCorelApplication()   ' tanpa ActiveDocumentInstance
            Dim service As New CorelConnectionService(prompt,
                Function(progId As String) DirectCast(Nothing, Object),
                Function(progId As String) DirectCast(factory, Object),
                New String() {"CorelDRAW.Application.24"})

            Dim message As String = ""
            Try
                service.Connect()
            Catch ex As Exception
                message = ex.Message
            End Try
            Check.Contains(message, "belum memiliki dokumen aktif", "alasan")
            Check.IsTrue(prompt.Notifications.Count >= 1, "pengguna diberi tahu")
            Check.Contains(prompt.Notifications(0), "tanpa dokumen", "jenis pemberitahuan")
        End Sub)

        suite.Add("instance berjalan tanpa dokumen -> bertanya sebelum membuka baru", Sub()
            Dim running As New FakeCorelApplication()   ' ada, tetapi tanpa dokumen
            Dim prompt As New FakeCorelPrompt()
            prompt.AllowNewInstance = False
            Dim service As New CorelConnectionService(prompt,
                Function(progId As String) DirectCast(running, Object),
                Function(progId As String) DirectCast(Nothing, Object),
                New String() {"CorelDRAW.Application.24"})

            Dim message As String = ""
            Try
                service.Connect()
            Catch ex As Exception
                message = ex.Message
            End Try
            Check.AreEqual(1, prompt.OpenInstanceCalls, "bertanya")
            Check.Contains(message, "Tidak ada dokumen aktif", "detail dilaporkan")
        End Sub)

        suite.Add("EnsureDocument: dokumen baru diambil dari App", Sub()
            Dim running As New FakeCorelApplication()
            running.OpenDocument("C:\uji\a.cdr")
            running.ActiveDocumentInstance = running.Opened(0)
            Dim prompt As New FakeCorelPrompt()
            Dim service As New CorelConnectionService(prompt,
                Function(progId As String) DirectCast(running, Object),
                Function(progId As String) DirectCast(Nothing, Object),
                New String() {"CorelDRAW.Application.24"})
            service.Connect()

            Dim doc As Object = service.EnsureDocument()
            Check.IsNotNull(doc, "dokumen tersedia")
            Check.AreEqual("C:\uji\a.cdr", DirectCast(doc, FakeDocument).FilePath, "dokumen yang benar")
        End Sub)

        suite.Add("EnsureDocument: tanpa dokumen -> exception (bukan NullReference)", Sub()
            Dim prompt As New FakeCorelPrompt()
            Dim service As New CorelConnectionService(prompt, Nothing, Nothing, New String() {})
            Dim message As String = ""
            Try
                service.EnsureDocument()
            Catch ex As Exception
                message = ex.Message
            End Try
            Check.Contains(message, "Belum terhubung", "pesan jelas")
        End Sub)

        suite.Add("identitas dokumen berbeda -> EnsureSameDocument False + dilaporkan", Sub()
            Dim running As New FakeCorelApplication()
            running.OpenDocument("C:\uji\a.cdr")
            running.ActiveDocumentInstance = running.Opened(0)
            Dim prompt As New FakeCorelPrompt()
            Dim service As New CorelConnectionService(prompt,
                Function(progId As String) DirectCast(running, Object),
                Function(progId As String) DirectCast(Nothing, Object),
                New String() {"CorelDRAW.Application.24"})
            service.Connect()
            service.RememberLoadedIdentity()

            ' Dokumen berganti (file lain) tanpa mengganti instance.
            running.ActiveDocumentInstance = New FakeDocument("C:\uji\lain.cdr")
            ' Perilaku lama: ConnectToCorel dipanggil di awal setiap operasi,
            ' jadi dokumen aktif dibaca ulang dari instance sebelum dibandingkan.
            service.Connect()

            Check.IsFalse(service.EnsureSameDocument(), "ditolak")
            Check.IsTrue(prompt.MismatchReported, "pengguna diberi tahu")
        End Sub)

        suite.Add("dokumen sama -> EnsureSameDocument True (jumlah halaman boleh berubah)", Sub()
            Dim running As New FakeCorelApplication()
            running.OpenDocument("C:\uji\a.cdr")
            running.ActiveDocumentInstance = running.Opened(0)
            Dim prompt As New FakeCorelPrompt()
            Dim service As New CorelConnectionService(prompt,
                Function(progId As String) DirectCast(running, Object),
                Function(progId As String) DirectCast(Nothing, Object),
                New String() {"CorelDRAW.Application.24"})
            service.Connect()
            service.RememberLoadedIdentity()

            running.ActiveDocumentInstance.Pages.Add(New FakePage(2))
            Check.IsTrue(service.EnsureSameDocument(), "tetap dianggap dokumen yang sama")
            Check.IsFalse(prompt.MismatchReported, "tidak ada peringatan")
        End Sub)

        suite.Add("Release melepas referensi (Connect berikutnya mencari ulang)", Sub()
            Dim running As New FakeCorelApplication()
            running.OpenDocument("C:\uji\a.cdr")
            running.ActiveDocumentInstance = running.Opened(0)
            Dim resolveCount As Integer = 0
            Dim prompt As New FakeCorelPrompt()
            Dim service As New CorelConnectionService(prompt,
                Function(progId As String)
                    resolveCount += 1
                    Return DirectCast(running, Object)
                End Function,
                Function(progId As String) DirectCast(Nothing, Object),
                New String() {"CorelDRAW.Application.24"})

            service.Connect()
            service.Release()
            Check.IsTrue(service.App Is Nothing, "App dilepas")
            Check.IsTrue(service.Document Is Nothing, "Document dilepas")
            service.Connect()
            Check.IsTrue(resolveCount >= 2, "pencarian instance diulang setelah Release")
        End Sub)

        suite.Add("CorelComProbe: ProgID default 2022 lebih dulu", Sub()
            Dim ids As String() = CorelComProbe.DefaultProgramIds()
            Check.AreEqual("CorelDRAW.Application.24", ids(0), "ProgID 2022 di urutan pertama")
            Check.AreEqual("CorelDRAW.Application", ids(1), "fallback generik")
        End Sub)
    End Sub
End Module
