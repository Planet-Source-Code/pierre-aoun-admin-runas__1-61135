Attribute VB_Name = "Encryption"
'Encryption with drive SN and computer SN

Private Declare Function GetVolumeInformation Lib _
   "kernel32.dll" Alias "GetVolumeInformationA" _
   (ByVal lpRootPathName As String, _
   ByVal lpVolumeNameBuffer As String, _
   ByVal nVolumeNameSize As Integer, _
   lpVolumeSerialNumber As Long, _
   lpMaximumComponentLength As Long, _
   lpFileSystemFlags As Long, _
   ByVal lpFileSystemNameBuffer As String, _
   ByVal nFileSystemNameSize As Long) As Long
   'Drive serial number
 Private Function DriveSerialNumber(ByVal Drive As String) As String
 On Error Resume Next
    Dim lAns As Long
    Dim lRet As Long
    Dim sVolumeName As String, sDriveType As String
    Dim sDrive As String
    sDrive = Drive
    If Len(sDrive) = 1 Then
        sDrive = sDrive & ":\"
    ElseIf Len(sDrive) = 2 And Right(sDrive, 1) = ":" Then
        sDrive = sDrive & "\"
    End If
    sVolumeName = String$(255, Chr$(0))
    sDriveType = String$(255, Chr$(0))
    lRet = GetVolumeInformation(sDrive, sVolumeName, _
    255, lAns, 0, 0, sDriveType, 255)
DriveSerialNumber = CStr(lAns)
End Function
    'computer SN
Private Function ComputerSN() As String
On Error Resume Next
    Dim SNPC As String
    Dim winmgmt1, SNSet, SN
SNPC = ""
winmgmt1 = "winmgmts:{impersonationLevel=impersonate}!//" & SNPC & ""
Set SNSet = GetObject(winmgmt1).InstancesOf("Win32_BIOS")
For Each SN In SNSet
   ComputerSN = CStr(SN.SerialNumber)
Next
End Function

Public Function GetSNMachine() As String
GetSNMachine = ComputerSN + DriveSerialNumber("c:")
End Function

Public Function Encrypt(ByVal Mots As Collection, ByVal SerialN As String) As String
    Const MotNum = 10
    Dim GlobalMot As String
    Dim AscNum As Integer
    Dim i As Integer
    Dim j As Integer
    Dim EncSN As String
    Dim SN256 As String
    Dim Mot(0 To MotNum) As String
    Dim Mots256(0 To MotNum) As String
    Dim Mots256Xor(0 To MotNum) As String
    On Error Resume Next
    For i = 1 To MotNum
        Mot(i) = Mots.Item(i)
        If Mot(i) = "" Then Mot(i) = "~Nothing~"
    Next i
    EncSN = "" 'Small encryption for SerialN
    If SerialN = "" Then SerialN = "0123456789"
    For i = 1 To Len(SerialN)
        EncSN = EncSN + Mid(SerialN, i, 1) + Chr(Asc("A") + i) + Chr(Asc("Z") + i)
    Next i
    
   'disp SN to 256
    SN256 = ""
    For i = 1 To 256 Step Len(EncSN)
        SN256 = SN256 + EncSN
    Next i
    SN256 = Left(SN256 + EncSN, 256)
    Mots256(0) = Chr(Len(EncSN)) + Left(SN256, 255)
    
    For j = 1 To MotNum
        Mots256(j) = Chr(Len(Mot(j)))
        For i = 1 To 256 Step Len(Mot(j))
            Mots256(j) = Mots256(j) + Mot(j)
        Next i
        Mots256(j) = Mots256(j) + Mot(j)
        Mots256(j) = Left(Mots256(j), 256)
    Next j
    For j = 0 To MotNum
    Mots256Xor(j) = ""
        For i = 1 To 256
           Mots256Xor(j) = Mots256Xor(j) + Chr(Asc(Mid(Mots256(j), i, 1)) Xor Asc(Mid(SN256, i, 1)))
        Next i
    Next j
    GlobalMot = ""
    For i = 1 To 256 Step 8
      For j = 0 To MotNum
        GlobalMot = GlobalMot + Mid(Mots256Xor(j), i, 8)
      Next j
    Next i
   Encrypt = GlobalMot
End Function

Public Sub Decrypt(ByVal MotEnc As String, ByVal SerialN As String, Data As Collection)
    Const MotNum = 10
    Dim GlobalMot As String
    Dim AscNum As Integer
    Dim i As Integer
    Dim j As Integer
    Dim EncSN As String
    Dim SN256 As String
    Dim TestSN As String
    Dim Mot(0 To MotNum) As String
    Dim Mots256(0 To MotNum) As String
    Dim Mots256Xor(0 To MotNum) As String
    Dim buf As String
    Set Data = New Collection
    On Error Resume Next
    EncSN = "" 'Small encryption for SerialN
    If SerialN = "" Then SerialN = "0123456789"
    For i = 1 To Len(SerialN)
        EncSN = EncSN + Mid(SerialN, i, 1) + Chr(Asc("A") + i) + Chr(Asc("Z") + i)
    Next i
    SN256 = ""
    For i = 1 To 256 Step Len(EncSN)
        SN256 = SN256 + EncSN
    Next i
    SN256 = SN256 + EncSN
    SN256 = Left(SN256, 256)
    TestSN = Chr(Len(EncSN)) + Left(SN256, 255)
    '------------------
    If Len(MotEnc) <> (MotNum + 1) * 256 Then GoTo EncErr
    GlobalMot = MotEnc
     For i = 1 To Len(GlobalMot) Step 8 * (MotNum + 1)
        For j = 0 To MotNum
            Mots256Xor(j) = Mots256Xor(j) + Mid(GlobalMot, i + j * 8, 8)
        Next j
    Next i
    
    For j = 0 To MotNum
    Mots256(j) = ""
        For i = 1 To 256
           Mots256(j) = Mots256(j) + Chr(Asc(Mid(Mots256Xor(j), i, 1)) Xor Asc(Mid(SN256, i, 1)))
        Next i
    Next j
     If Mots256(0) <> TestSN Then GoTo EncErr
    For i = 1 To MotNum
        AscNum = Asc(Left(Mots256(i), 1))
        buf = Mid(Mots256(i), 2, AscNum)
        If buf = "~Nothing~" Then buf = ""
        Data.Add buf
    Next i
    Exit Sub
EncErr:
    'MsgBox "Encryption Error!"
    For i = 1 To MotNum
        Data.Add ""
    Next i
End Sub

