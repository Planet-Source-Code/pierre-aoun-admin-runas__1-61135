Attribute VB_Name = "GetIconExe"
Option Explicit
Private Type PicBmp
   Size As Long
   tType As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type
Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type
Private Declare Function OleCreatePictureIndirect _
Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, _
ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" _
Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal _
nIconIndex As Long, phiconLarge As Long, phiconSmall As _
Long, ByVal nIcons As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal _
hicon As Long) As Long
Public Function GetIconExec(FileName As String, _
IconIndex As Long, UseLargeIcon As Boolean) As Picture
Dim hlargeicon As Long
Dim hsmallicon As Long
Dim selhandle As Long
Dim pic As PicBmp
Dim IPic As IPicture
Dim IID_IDispatch As GUID
If ExtractIconEx(FileName, IconIndex, hlargeicon, _
hsmallicon, 1) > 0 Then
If UseLargeIcon Then
selhandle = hlargeicon
Else
selhandle = hsmallicon
End If
With IID_IDispatch
.Data1 = &H20400
.Data4(0) = &HC0
.Data4(7) = &H46
End With
With pic
.Size = Len(pic)
.tType = vbPicTypeIcon
.hBmp = selhandle
End With
Call OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)
Set GetIconExec = IPic
DestroyIcon hsmallicon
DestroyIcon hlargeicon
End If
End Function





