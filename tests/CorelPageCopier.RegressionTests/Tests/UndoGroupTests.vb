Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  UndoGroupTests.vb
'  Regression test Temuan C (undo grouping): kegagalan tidak ditelan, dan
'  EndCommandGroup hanya dipanggil bila group benar-benar terbuka.
' ============================================================================

Imports System

Public Module UndoGroupTests

    Public Sub Register(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("Undo grouping (Temuan C)")

        suite.Add("group berhasil dibuka & ditutup -> satu Ctrl+Z dijamin", Sub()
            Dim guard As New UndoGroupGuard("uji")
            guard.MarkBeginAttempted()
            guard.MarkBeginSucceeded()
            Check.IsTrue(guard.IsOpen, "terbuka")
            Check.IsTrue(guard.ShouldCallEnd, "perlu End")
            guard.MarkEndSucceeded()
            Check.IsTrue(guard.SingleUndoGuaranteed, "jaminan satu undo")
            Check.IsFalse(guard.HasProblem, "tanpa masalah")
            Check.Contains(guard.Describe(), "satu Ctrl+Z", "laporan menjanjikan undo")
        End Sub)

        suite.Add("group gagal dibuka -> tidak menjanjikan satu undo", Sub()
            Dim guard As New UndoGroupGuard("uji")
            guard.MarkBeginAttempted()
            guard.MarkBeginFailed("simulasi: tidak didukung")
            Check.IsTrue(guard.BeginFailed, "kegagalan tercatat")
            Check.IsFalse(guard.SingleUndoGuaranteed, "tidak boleh janji satu undo")
            Check.IsFalse(guard.ShouldCallEnd, "End TIDAK boleh dipanggil")
            Check.IsTrue(guard.HasProblem, "wajib dilaporkan")
            Check.Contains(guard.Describe(), "TIDAK dapat dibuka", "laporan eksplisit")
            Check.Contains(guard.Describe(), "simulasi: tidak didukung", "detail error disertakan")
        End Sub)

        suite.Add("End gagal -> dilaporkan, bukan ditelan", Sub()
            Dim guard As New UndoGroupGuard("uji")
            guard.MarkBeginAttempted()
            guard.MarkBeginSucceeded()
            guard.MarkEndFailed("simulasi: end gagal")
            Check.IsTrue(guard.EndFailed, "kegagalan End tercatat")
            Check.IsTrue(guard.HasProblem, "wajib dilaporkan")
            Check.IsFalse(guard.SingleUndoGuaranteed, "tidak menjanjikan undo")
            Check.Contains(guard.Describe(), "EndCommandGroup GAGAL", "laporan eksplisit")
        End Sub)

        suite.Add("End hanya sekali (ShouldCallEnd false setelahnya)", Sub()
            Dim guard As New UndoGroupGuard("uji")
            guard.MarkBeginAttempted()
            guard.MarkBeginSucceeded()
            guard.MarkEndSucceeded()
            Check.IsFalse(guard.ShouldCallEnd, "tidak memanggil End dua kali")
        End Sub)

        suite.Add("tanpa Begin, End tidak dipanggil", Sub()
            Dim guard As New UndoGroupGuard("uji")
            Check.IsFalse(guard.ShouldCallEnd, "End tidak dipanggil")
            Check.IsFalse(guard.SingleUndoGuaranteed, "tidak ada jaminan")
            Check.IsFalse(guard.HasProblem, "bukan kegagalan")
        End Sub)
    End Sub
End Module
