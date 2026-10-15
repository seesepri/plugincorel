Option Explicit
Public Sub LaunchCorelPageCopier()
    On Error GoTo HandleError
    Dim exePath As String
    exePath = "C:\Program Files\Corel Page Copier\CorelPageCopier.exe"
    If Dir$(exePath) = "" Then
        MsgBox "Corel Page Copier tidak ditemukan:" & vbCrLf & exePath & vbCrLf & "Edit path EXE di macro ToolbarLauncher.bas.", vbExclamation, "Corel Page Copier"
        Exit Sub
    End If
    CreateObject("WScript.Shell").Run """" & exePath & """", 1, False
    Exit Sub
HandleError:
    MsgBox "Gagal menjalankan Corel Page Copier: " & Err.Description, vbCritical, "Corel Page Copier"
End Sub
