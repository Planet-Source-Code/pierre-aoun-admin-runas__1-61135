Attribute VB_Name = "RunasAp"
Public AdminUser As String
Public AdminPwd As String
Public Type STARTUPINFOW
cb As Long
lpReserved As Long
lpDesktop As Long
lpTitle As Long
dwX As Long
dwY As Long
dwXSize As Long
dwYSize As Long
dwXCountChars As Long
dwYCountChars As Long
dwFillAttribute As Long
dwFlags As Long
wShowWindow As Integer
cbReserved2 As Integer
lpReserved2 As Long
hStdInput As Long
hStdOutput As Long
hStdError As Long
End Type
Public Type PROCESS_INFORMATION
hProcess As Long
hThread As Long
dwProcessId As Long
dwThreadId As Long
End Type
Public Const LOGON_WITH_PROFILE As Long = &H1&
Public Const LOGON_NETCREDENTIALS_ONLY As Long = &H2&
Public Const WAIT_TIMEOUT = 258&
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Public Declare Function GetCommandLine Lib "kernel32" Alias _
"GetCommandLineA" () As Long
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" ( _
ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias _
"GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function CreateProcessWithLogonW Lib "advapi32" ( _
ByVal lpUsername As Long, ByVal lpDomain As Long, _
ByVal lpPassword As Long, ByVal dwLogonFlags As Long, _
ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, _
ByVal dwCreationFlags As Long, lpEnvironment As Any, _
ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFOW, _
lpProcessInfo As PROCESS_INFORMATION) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias _
"GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function FormatMessage Lib "kernel32" Alias _
"FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, _
ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
ByVal lpBuffer As String, ByVal nSize As Long, _
Arguments As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" ( _
ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" ( _
ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Function AppPath() As String
Dim lpStr As Long, i As Long
Dim Buffer As String
Dim exePath As String
lpStr = GetCommandLine()
Buffer = Space$(512)
lstrcpy Buffer, lpStr
Buffer = Left$(Buffer, InStr(Buffer & vbNullChar, vbNullChar) - 1)
If Left$(Buffer, 1) = """" Then
i = InStr(2, Buffer, """")
exePath = Mid$(Buffer, 2, i - 2)
Else
i = InStr(Buffer, " ")
exePath = Left$(Buffer, i - 1)
End If
AppPath = Left(exePath, Len(exePath) - InStr(1, StrReverse(exePath), _
"\"))
End Function

Public Function AppExeName() As String
Dim lpStr As Long, i As Long
Dim Buffer As String
Dim exePath As String

lpStr = GetCommandLine()
Buffer = Space$(512)
lstrcpy Buffer, lpStr
Buffer = Left$(Buffer, InStr(Buffer & vbNullChar, vbNullChar) - 1)
If Left$(Buffer, 1) = """" Then
i = InStr(2, Buffer, """")
exePath = Mid$(Buffer, 2, i - 2)
Else
i = InStr(Buffer, " ")
exePath = Left$(Buffer, i - 1)
End If
AppExeName = Mid(exePath, Len(exePath) - InStr(1, _
StrReverse(exePath), "\") + 2)
End Function

Public Function CommandLine() As String
Dim lpStr As Long, i As Long
Dim Buffer As String
Dim cmdLine As String

lpStr = GetCommandLine()
Buffer = Space$(512)
lstrcpy Buffer, lpStr
Buffer = Left$(Buffer, InStr(Buffer & vbNullChar, vbNullChar) - 1)
If Left$(Buffer, 1) = """" Then
i = InStr(2, Buffer, """")
cmdLine = LTrim$(Mid$(Buffer, i + 1))
Else
i = InStr(Buffer, " ")
cmdLine = LTrim$(Mid$(Buffer, i))
End If
CommandLine = cmdLine
End Function

Function UserName() As String
Dim lpBuffer As String
Dim nSize As Long
Dim lError As Long
lpBuffer = Space(255)
nSize = Len(lpBuffer)
Call GetUserName(lpBuffer, nSize)
UserName = Left(lpBuffer, InStr(1, lpBuffer, Chr(0)) - 1)
End Function

Public Function GetErrorMessage(Error As Long) As String
Dim Buffer As String
Dim lBuffer As Long
Buffer = String(1024, 0)
lBuffer = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, Error, _
0, Buffer, 200, ByVal 0&)
GetErrorMessage = Left(Buffer, lBuffer)
End Function

Public Function ComputerName() As String
Dim lpBuffer As String
Dim nSize As Long
Dim lError As Long
lpBuffer = Space(255)
nSize = Len(lpBuffer)
Call GetComputerName(lpBuffer, nSize)
ComputerName = Left(lpBuffer, nSize)
End Function
Public Function RunAs(sUser As String, sPwd As String, _
sCmdLine As String, Optional Parameters As String = "", _
Optional Directory As String = "", _
Optional WindowStyle As VbAppWinStyle = vbNormalFocus, _
Optional Wait As Boolean = False, Optional Timeout As Long = -1, _
Optional Terminate As Boolean = False, _
Optional hProcess As Long) As Long
Dim SInfo As STARTUPINFOW
Dim PInfo As PROCESS_INFORMATION
Dim aUser() As String
Dim sDomain As String
Dim sUsername As String
Dim sDir As String
Dim sCmd As String
Dim Res As Long

aUser = Split(sUser, "\")
If UBound(aUser) = 1 Then
sDomain = aUser(0)
sUsername = aUser(1)
Else
sDomain = ComputerName
sUsername = sUser
End If

SInfo.dwFlags = STARTF_USESHOWWINDOW
SInfo.wShowWindow = WindowStyle

If Directory = "" Then
sDir = CurDir
Else
sDir = Directory
End If

If Parameters <> "" Then
sCmd = sCmdLine & " " & Parameters
Else
sCmd = sCmdLine
End If

Res = CreateProcessWithLogonW(StrPtr(sUsername), StrPtr(sDomain), _
StrPtr(sPwd), LOGON_WITH_PROFILE, 0&, StrPtr(sCmd), 0&, ByVal 0&, _
StrPtr(sDir), SInfo, PInfo)

If Res <> 0 Then
hProcess = PInfo.hProcess
If Wait Then
If Timeout > 0 Then Timeout = Timeout * 1000
If WaitForSingleObject(PInfo.hProcess, _
Timeout) = WAIT_TIMEOUT Then
RunAs = WAIT_TIMEOUT
If Terminate Then
If TerminateProcess(PInfo.hProcess, 0) = 0 Then
RunAs = Err.LastDllError
End If
End If
End If
End If
Else
RunAs = Err.LastDllError
hProcess = 0
End If
End Function



