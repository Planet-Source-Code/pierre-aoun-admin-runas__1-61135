Attribute VB_Name = "ComDlgCtrl"
Option Explicit
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public ComDlgBx As OPENFILENAME

Public Function ShowOpen() As String
    'Show the 'Open File' dialog
    If GetOpenFileName(ComDlgBx) Then
        ShowOpen = Trim$(ComDlgBx.lpstrFile)
    Else
        ShowOpen = ""
    End If
End Function

