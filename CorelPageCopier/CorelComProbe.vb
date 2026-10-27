Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  CorelComProbe.vb
'  Satu-satunya tempat yang menghubungi COM "di luar" dokumen:
'    - mencari instance CorelDRAW yang SUDAH berjalan (Running Object Table),
'    - membuat instance baru dari ProgID.
'
'  Desain:
'    Pemanggilan Marshal.GetActiveObject dilakukan lewat reflection supaya file
'    ini tetap bisa dikompilasi di proyek test (net8.0) tanpa mengubah cara
'    kerja late-bound COM. Exception asli dari COM diteruskan apa adanya
'    (TargetInvocationException dibuka), jadi perilaku pemanggil tidak berubah.
'
'  Tidak ada interop/PIA: tetap late bound, sama seperti sebelum refactor.
' ============================================================================

Imports System
Imports System.Reflection
Imports System.Runtime.ExceptionServices
Imports System.Runtime.InteropServices

Public Class CorelComProbe

    ''' <summary>ProgID yang dicoba, urut prioritas (2022 lebih dulu).</summary>
    Public Shared Function DefaultProgramIds() As String()
        Return New String() {"CorelDRAW.Application.24", "CorelDRAW.Application"}
    End Function

    ''' <summary>
    ''' Ambil instance CorelDRAW yang sedang berjalan dari ROT.
    ''' Mengembalikan Nothing bila tidak ada; exception COM lain diteruskan.
    ''' </summary>
    Public Shared Function GetRunningObject(progId As String) As Object
        Dim method As MethodInfo = Nothing
        Try
            method = GetType(Marshal).GetMethod("GetActiveObject", New Type() {GetType(String)})
        Catch ex As Exception
            Throw New PlatformNotSupportedException(
                "Marshal.GetActiveObject tidak tersedia di runtime ini.", ex)
        End Try
        If method Is Nothing Then
            Throw New PlatformNotSupportedException(
                "Marshal.GetActiveObject tidak tersedia di runtime ini.")
        End If

        Try
            Return method.Invoke(Nothing, New Object() {progId})
        Catch tiex As TargetInvocationException
            If tiex.InnerException IsNot Nothing Then
                ExceptionDispatchInfo.Capture(tiex.InnerException).Throw()
            End If
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Buat instance CorelDRAW baru dari ProgID (Activator, seperti sebelum
    ''' refactor). Mengembalikan Nothing bila ProgID tidak terdaftar.
    ''' </summary>
    Public Shared Function CreateInstance(progId As String) As Object
        Dim appType As Type = Nothing
        Try
            ' Type.GetTypeFromProgID ada di .NET Framework; di runtime lain
            ' bisa tidak tersedia -> diperlakukan sebagai "tidak terdaftar".
            ' 'getType' bentrok dengan kata kunci GetType -> typeLookup.
            Dim typeLookup As MethodInfo = GetType(Type).GetMethod(
                "GetTypeFromProgID", New Type() {GetType(String), GetType(Boolean)})
            If typeLookup Is Nothing Then Return Nothing
            appType = DirectCast(typeLookup.Invoke(Nothing, New Object() {progId, False}), Type)
        Catch tiex As TargetInvocationException
            If tiex.InnerException IsNot Nothing Then
                ExceptionDispatchInfo.Capture(tiex.InnerException).Throw()
            End If
            Throw
        End Try

        If appType Is Nothing Then Return Nothing
        Return Activator.CreateInstance(appType)
    End Function
End Class
