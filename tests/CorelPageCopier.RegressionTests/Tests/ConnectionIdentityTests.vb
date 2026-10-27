Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  ConnectionIdentityTests.vb
'  Regression test Temuan F: dokumen dari instance berbeda tidak dianggap
'  target yang sama, dan identitas dokumen bisa ditampilkan.
' ============================================================================

Imports System

Public Module ConnectionIdentityTests

    Public Sub Register(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("Koneksi & identitas dokumen (Temuan F)")

        suite.Add("dokumen sama di instance sama -> SameTarget True", Sub()
            Dim first As New DocumentIdentity("a.cdr", "C:\x\a.cdr", 3, 111)
            Dim second As New DocumentIdentity("a.cdr", "C:\x\a.cdr", 4, 111)
            Check.IsTrue(first.SameTarget(second), "dokumen sama (jumlah halaman boleh berubah)")
        End Sub)

        suite.Add("dokumen sama di instance BERBEDA -> bukan target yang sama", Sub()
            Dim first As New DocumentIdentity("a.cdr", "C:\x\a.cdr", 3, 111)
            Dim second As New DocumentIdentity("a.cdr", "C:\x\a.cdr", 3, 222)
            Check.IsFalse(first.SameTarget(second), "instance berbeda harus ditolak")
            Check.IsFalse(first.Key = second.Key, "key berbeda")
        End Sub)

        suite.Add("dokumen berbeda di instance sama -> bukan target yang sama", Sub()
            Dim first As New DocumentIdentity("a.cdr", "C:\x\a.cdr", 3, 111)
            Dim second As New DocumentIdentity("b.cdr", "C:\x\b.cdr", 3, 111)
            Check.IsFalse(first.SameTarget(second), "path berbeda")
        End Sub)

        suite.Add("dokumen tanpa path dibandingkan lewat nama", Sub()
            Dim first As New DocumentIdentity("Tanpa Judul-1", "", 2, 5)
            Dim second As New DocumentIdentity("Tanpa Judul-1", "", 2, 5)
            Check.IsTrue(first.SameTarget(second), "nama sama")
            Check.IsFalse(first.SameTarget(New DocumentIdentity("Tanpa Judul-2", "", 2, 5)), "nama berbeda")
        End Sub)

        suite.Add("Describe menampilkan nama file + jumlah halaman", Sub()
            Dim identity As New DocumentIdentity("a.cdr", "C:\x\a.cdr", 7, 1)
            Check.Contains(identity.Describe(), "a.cdr", "nama file")
            Check.Contains(identity.Describe(), "7 halaman", "jumlah halaman")
        End Sub)

        suite.Add("Describe menangani jumlah halaman tidak terbaca", Sub()
            Dim identity As New DocumentIdentity("a.cdr", "C:\x\a.cdr", -1, 1)
            Check.Contains(identity.Describe(), "tidak terbaca", "jujur soal kegagalan baca")
        End Sub)

        suite.Add("FromDocument membaca dokumen palsu", Sub()
            ' Path dibuat eksplisit (bukan lewat konstruktor) supaya test tetap
            ' sama di Windows maupun Linux.
            Dim doc As New FakeDocument()
            doc.FilePath = "C:\x\uji.cdr"
            doc.Name = "uji.cdr"
            doc.Pages.Add(New FakePage(1))
            doc.Pages.Add(New FakePage(2))
            Dim identity As DocumentIdentity = DocumentIdentity.FromDocument(doc, 42)
            Check.AreEqual("uji.cdr", identity.Name, "nama")
            Check.AreEqual("C:\x\uji.cdr", identity.FilePath, "path")
            Check.AreEqual(2, identity.PageCount, "jumlah halaman")
            Check.AreEqual(42, identity.InstanceToken, "token instance")
        End Sub)

        suite.Add("FromDocument aman untuk dokumen Nothing", Sub()
            Dim identity As DocumentIdentity = DocumentIdentity.FromDocument(Nothing, 7)
            Check.IsTrue(identity.IsEmpty, "identitas kosong")
            Check.AreEqual(-1, identity.PageCount, "jumlah halaman tidak diketahui")
        End Sub)

        suite.Add("InstanceTokenFor memakai handle jendela bila ada", Sub()
            Dim app As New FakeAppWithWindow()
            app.AppWindow.Handle = 9876
            Check.AreEqual(9876, DocumentIdentity.InstanceTokenFor(app), "handle jendela")
            Dim appWithoutWindow As New FakeAppWithWindow()
            appWithoutWindow.AppWindow.FailHandle = True
            Check.IsTrue(DocumentIdentity.InstanceTokenFor(appWithoutWindow) <> 0, "fallback hash proxy")
        End Sub)

        suite.Add("ComObjectHealth: objek Nothing dianggap mati", Sub()
            Check.IsFalse(ComObjectHealth.IsAlive(Nothing), "Nothing mati")
            Check.IsTrue(ComObjectHealth.IsAlive(New FakeDocument()), "objek .NET dianggap hidup")
        End Sub)
    End Sub

    ''' <summary>Objek aplikasi palsu dengan AppWindow.Handle (meniru Application.AppWindow).</summary>
    Public Class FakeAppWithWindow
        Public ReadOnly Property AppWindow As FakeAppWindow = New FakeAppWindow()
    End Class

    Public Class FakeAppWindow
        Private _handle As Integer = 0
        Public Property FailHandle As Boolean = False

        Public Property Handle As Integer
            Get
                If FailHandle Then Throw New InvalidOperationException("simulasi: AppWindow tidak tersedia")
                Return _handle
            End Get
            Set(value As Integer)
                _handle = value
            End Set
        End Property
    End Class
End Module
