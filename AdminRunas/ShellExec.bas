Attribute VB_Name = "ShellExec"
Option Explicit

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
   (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWDEFAULT As Long = 10

Public Sub ShellExe(sFile As Variant)
Dim Res As Long
Dim ShellStr As String
ShellStr = "rundll32.exe shell32.dll,ShellExec_RunDLL " + sFile
If AdminUser <> "" Then
    Res = RunAs(AdminUser, AdminPwd, ShellStr)   'CommandLine
    If Res <> 0 Then MsgBox GetErrorMessage(Res)
Else
    Shell ShellStr, vbNormalFocus
End If
End Sub

