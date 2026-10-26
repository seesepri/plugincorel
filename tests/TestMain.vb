Option Strict Off
Imports System

''' <summary>Runner unit test sederhana: tanpa framework, hanya PASS/FAIL.</summary>
Module TestMain
    Public Failures As Integer = 0
    Public Total As Integer = 0

    Public Sub Check(cond As Boolean, name As String)
        Total += 1
        If cond Then
            Console.WriteLine("  PASS  " & name)
        Else
            Failures += 1
            Console.WriteLine("  FAIL  " & name)
        End If
    End Sub

    Sub Section(title As String)
        Console.WriteLine()
        Console.WriteLine(title & ":")
    End Sub

    Sub Main()
        TokenTests.Run()
        ImportLogoTests.Run()
        LayoutAndResultTests.Run()

        Console.WriteLine()
        Console.WriteLine("Total assertion: " & Total & " | gagal: " & Failures)
        If Failures = 0 Then
            Console.WriteLine("SEMUA TES LULUS")
        Else
            Console.WriteLine(Failures & " TES GAGAL")
            Environment.Exit(1)
        End If
    End Sub
End Module
