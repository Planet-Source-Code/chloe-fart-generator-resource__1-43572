Attribute VB_Name = "basCOMMON"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Public Sub ExecuteLink(LINK As String)
    On Error Resume Next
    Dim lRet As Long
    If LINK <> "" Then
        lRet = ShellExecute(0, "open", LINK, "", App.Path, SW_SHOWNORMAL)
        If lRet >= 0 And lRet <= 32 Then
            MsgBox "Error jumping to:" & LINK, 48, "Warning"
        End If
    End If
End Sub



