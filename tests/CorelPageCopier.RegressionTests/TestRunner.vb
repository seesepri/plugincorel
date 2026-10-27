Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  TestRunner.vb
'  Harness unit test minimal (tanpa framework eksternal) untuk logika yang
'  dapat diuji TANPA CorelDRAW.
'
'  Jalankan:  dotnet run --project tests/CorelPageCopier.RegressionTests
'  Exit code 0 = semua lulus, 1 = ada yang gagal.
' ============================================================================

Imports System
Imports System.Collections.Generic

Public Class AssertionException
    Inherits Exception

    Public Sub New(message As String)
        MyBase.New(message)
    End Sub
End Class

''' <summary>Assertion sederhana: melempar AssertionException bila gagal.</summary>
Public Class Check

    Public Shared Sub That(condition As Boolean, message As String)
        If Not condition Then Throw New AssertionException(message)
    End Sub

    Public Shared Sub IsTrue(condition As Boolean, message As String)
        That(condition, message)
    End Sub

    Public Shared Sub IsFalse(condition As Boolean, message As String)
        That(Not condition, message)
    End Sub

    Public Shared Sub IsNothing(value As Object, message As String)
        That(value Is Nothing, message)
    End Sub

    Public Shared Sub IsNotNull(value As Object, message As String)
        That(value IsNot Nothing, message)
    End Sub

    Public Shared Sub AreEqual(expected As Object, actual As Object, message As String)
        Dim expectedText As String = If(expected Is Nothing, "(null)", expected.ToString())
        Dim actualText As String = If(actual Is Nothing, "(null)", actual.ToString())
        If Not String.Equals(expectedText, actualText, StringComparison.Ordinal) Then
            Throw New AssertionException(message & " — expected [" & expectedText & "] tetapi [" & actualText & "]")
        End If
    End Sub

    ''' <summary>Bandingkan angka dengan toleransi (untuk perhitungan geometri).</summary>
    Public Shared Sub Near(expected As Double, actual As Double, tolerance As Double, message As String)
        If Math.Abs(expected - actual) > tolerance Then
            Throw New AssertionException(message & " — expected " & expected.ToString() & " ±" &
                                          tolerance.ToString() & " tetapi " & actual.ToString())
        End If
    End Sub

    Public Shared Sub Contains(haystack As String, needle As String, message As String)
        Dim text As String = If(haystack, "")
        If text.IndexOf(needle, StringComparison.OrdinalIgnoreCase) < 0 Then
            Throw New AssertionException(message & " — teks tidak memuat [" & needle & "]. Teks: " & text)
        End If
    End Sub

    Public Shared Sub DoesNotContain(haystack As String, needle As String, message As String)
        Dim text As String = If(haystack, "")
        If text.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0 Then
            Throw New AssertionException(message & " — teks TIDAK boleh memuat [" & needle & "]. Teks: " & text)
        End If
    End Sub

    Public Shared Sub Fail(message As String)
        Throw New AssertionException(message)
    End Sub
End Class

''' <summary>Satu kelompok test (nama + aksi).</summary>
Public Class TestCase
    Public Name As String = ""
    Public Body As Action = Nothing

    Public Sub New(groupName As String, caseName As String, body As Action)
        Name = groupName & ": " & caseName
        Me.Body = body
    End Sub
End Class

''' <summary>Koleksi test + runner.</summary>
Public Class TestSuite
    Private ReadOnly _cases As New List(Of TestCase)()
    Private ReadOnly _group As String

    Public Sub New(groupName As String)
        _group = groupName
    End Sub

    Public Sub Add(caseName As String, body As Action)
        _cases.Add(New TestCase(_group, caseName, body))
    End Sub

    Public ReadOnly Property Cases As List(Of TestCase)
        Get
            Return _cases
        End Get
    End Property
End Class

''' <summary>Menjalankan semua suite dan mencetak ringkasan.</summary>
Public Class TestRunner

    Private ReadOnly _suites As New List(Of TestSuite)()

    Public Function AddSuite(groupName As String) As TestSuite
        Dim suite As New TestSuite(groupName)
        _suites.Add(suite)
        Return suite
    End Function

    Public Function RunAll() As Integer
        Dim passed As Integer = 0
        Dim failed As Integer = 0
        Dim failures As New List(Of String)()

        For Each suite As TestSuite In _suites
            For Each testCase As TestCase In suite.Cases
                Try
                    testCase.Body.Invoke()
                    passed += 1
                    Console.WriteLine("  PASS  " & testCase.Name)
                Catch ex As AssertionException
                    failed += 1
                    failures.Add(testCase.Name & " — " & ex.Message)
                    Console.WriteLine("  FAIL  " & testCase.Name)
                    Console.WriteLine("        " & ex.Message)
                Catch ex As Exception
                    failed += 1
                    failures.Add(testCase.Name & " — " & ex.GetType().Name & ": " & ex.Message)
                    Console.WriteLine("  ERROR " & testCase.Name)
                    Console.WriteLine("        " & ex.GetType().Name & ": " & ex.Message)
                End Try
            Next
        Next

        Console.WriteLine()
        Console.WriteLine("Ringkasan: " & passed & " lulus, " & failed & " gagal, " & (passed + failed) & " total.")
        If failed > 0 Then
            Console.WriteLine()
            Console.WriteLine("Test gagal:")
            For Each failure As String In failures
                Console.WriteLine("  • " & failure)
            Next
            Return 1
        End If
        Return 0
    End Function
End Class
