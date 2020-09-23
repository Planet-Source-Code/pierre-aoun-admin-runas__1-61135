VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Run as V1.0"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9585
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "email: pierre_aoun@hotmail.com"
      Top             =   840
      Width           =   3855
   End
   Begin VB.Frame fraUser 
      Caption         =   "Run as:"
      Height          =   1215
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   4575
      Begin VB.CheckBox chkPass 
         Caption         =   "(Save)    Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   750
         Width           =   1695
      End
      Begin VB.TextBox txtPass 
         Height          =   350
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtUser 
         Height          =   350
         Left            =   1920
         TabIndex        =   12
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Domain\user:"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Customize buttons ( Drag && Drop option )"
      Height          =   5295
      Left            =   4920
      TabIndex        =   2
      Top             =   1440
      Width           =   4455
      Begin VB.CommandButton cmdRunProg 
         Height          =   615
         Index           =   9
         Left            =   720
         OLEDropMode     =   1  'Manual
         TabIndex        =   20
         Top             =   4560
         Width           =   3375
      End
      Begin VB.CommandButton cmdRunProg 
         Height          =   615
         Index           =   8
         Left            =   720
         OLEDropMode     =   1  'Manual
         TabIndex        =   19
         Top             =   3840
         Width           =   3375
      End
      Begin VB.CommandButton cmdRunProg 
         Height          =   615
         Index           =   7
         Left            =   720
         OLEDropMode     =   1  'Manual
         TabIndex        =   10
         Top             =   3120
         Width           =   3375
      End
      Begin VB.CommandButton cmdRunProg 
         Height          =   615
         Index           =   6
         Left            =   720
         OLEDropMode     =   1  'Manual
         TabIndex        =   9
         Top             =   2400
         Width           =   3375
      End
      Begin VB.CommandButton cmdRunProg 
         Height          =   615
         Index           =   5
         Left            =   720
         OLEDropMode     =   1  'Manual
         TabIndex        =   8
         Top             =   1680
         Width           =   3375
      End
      Begin VB.CommandButton cmdRunProg 
         Height          =   615
         Index           =   4
         Left            =   720
         OLEDropMode     =   1  'Manual
         TabIndex        =   7
         Top             =   960
         Width           =   3375
      End
      Begin VB.CommandButton cmdRunProg 
         Height          =   615
         Index           =   3
         Left            =   720
         OLEDropMode     =   1  'Manual
         TabIndex        =   6
         Top             =   240
         Width           =   3375
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   9
         Left            =   120
         Stretch         =   -1  'True
         ToolTipText     =   "Choisir le programme à démarrer"
         Top             =   4635
         Width           =   495
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   8
         Left            =   120
         Stretch         =   -1  'True
         ToolTipText     =   "Choisir le programme à démarrer"
         Top             =   3900
         Width           =   495
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   7
         Left            =   120
         Stretch         =   -1  'True
         ToolTipText     =   "Choisir le programme à démarrer"
         Top             =   3180
         Width           =   495
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   6
         Left            =   120
         Stretch         =   -1  'True
         ToolTipText     =   "Choisir le programme à démarrer"
         Top             =   2460
         Width           =   495
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   5
         Left            =   120
         Stretch         =   -1  'True
         ToolTipText     =   "Choisir le programme à démarrer"
         Top             =   1740
         Width           =   495
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   4
         Left            =   120
         Stretch         =   -1  'True
         ToolTipText     =   "Choisir le programme à démarrer"
         Top             =   1005
         Width           =   495
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   3
         Left            =   120
         Stretch         =   -1  'True
         ToolTipText     =   "Choisir le programme à démarrer"
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Command ( Drag && Drop here )"
      Height          =   2055
      Left            =   240
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   4680
      Width           =   4575
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Height          =   495
         Left            =   2520
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox cmbRun 
         Height          =   315
         ItemData        =   "frmMain.frx":0442
         Left            =   120
         List            =   "frmMain.frx":0458
         OLEDropMode     =   1  'Manual
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   4335
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "Ok"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fixed Programs"
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   4575
      Begin VB.CommandButton cmdRunProg 
         Caption         =   "Computer Management"
         Height          =   615
         Index           =   2
         Left            =   960
         TabIndex        =   18
         Top             =   1920
         Width           =   3255
      End
      Begin VB.CommandButton cmdRunProg 
         Caption         =   "Services"
         Height          =   615
         Index           =   1
         Left            =   960
         TabIndex        =   17
         Top             =   1200
         Width           =   3255
      End
      Begin VB.CommandButton cmdRunProg 
         Caption         =   "Explorer"
         Height          =   615
         Index           =   0
         Left            =   960
         TabIndex        =   16
         Top             =   480
         Width           =   3255
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Index           =   2
         Left            =   240
         Picture         =   "frmMain.frx":048F
         Stretch         =   -1  'True
         ToolTipText     =   "Choisir le programme à démarrer"
         Top             =   1920
         Width           =   615
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Index           =   1
         Left            =   240
         Stretch         =   -1  'True
         ToolTipText     =   "Choisir le programme à démarrer"
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Index           =   0
         Left            =   240
         Stretch         =   -1  'True
         ToolTipText     =   "Choisir le programme à démarrer"
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Admin Runas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   5040
      TabIndex        =   22
      Top             =   240
      Width           =   4095
   End
   Begin VB.Menu mnAddProg 
      Caption         =   "mnAddProg"
      Visible         =   0   'False
      Begin VB.Menu AddProg 
         Caption         =   "Add Program..."
      End
      Begin VB.Menu Suprimer 
         Caption         =   "Delete"
      End
      Begin VB.Menu OpenWith 
         Caption         =   "Open With..."
      End
      Begin VB.Menu mnIcon 
         Caption         =   "Icon..."
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MAX_PATH = 255
Dim SystemDir As String
Dim WindowsDir As String
Dim indexGnrl As Integer
Dim TempString(0 To 30) As String
Dim TempImage(0 To 30) As String
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
 Private Sub PutSettings()
    Dim i As Integer
    Dim DataEnc As String
    Dim Data As Collection
    Set Data = New Collection
    
    Data.Add txtUser.Text
    Data.Add txtPass.Text
    Data.Add CStr(chkPass.Value)
    For i = 3 To 9
        Data.Add TempString(i)
        SaveSetting "Pierre Programs", "RunAs", "Icon" + CStr(i - 2), TempImage(i)
    Next i
    DataEnc = Encrypt(Data, GetSNMachine)
    SaveSetting "Pierre Programs", "RunAs", "Settings", DataEnc
End Sub
Private Sub GetSettings()
On Error Resume Next
    Dim i As Integer
    Dim DataEnc As String
    Dim Data As Collection
    Dim IconPic As String
    DataEnc = GetSetting("Pierre Programs", "RunAs", "Settings", "")
    Decrypt DataEnc, GetSNMachine, Data
    txtUser.Text = Data.Item(1)
    If CInt(Data.Item(3)) = 1 Then
        chkPass.Value = 1
        txtPass.Text = Data.Item(2)
    End If
     
    For i = 3 To 9
        TempString(i) = Data.Item(i + 1)
        cmdRunProg(i).Caption = Dir(TempString(i))
        If TempString(i) = "" Then cmdRunProg(i).Caption = ""
        cmdRunProg(i).ToolTipText = TempString(i)
        IconPic = GetSetting("Pierre Programs", "RunAs", "Icon" + CStr(i - 2), "")
        If IconPic <> "" Then
            imgIcon(i).Picture = GetIconExec(IconPic, 0, True)
            imgIcon(i).Picture = LoadPicture(IconPic)
            TempImage(i) = IconPic
        End If
    Next i
End Sub
Private Sub AddProg_Click()
On Error Resume Next
Call Add_Prog(indexGnrl)
PutSettings
End Sub
Private Sub chkPass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PutSettings
End Sub
Private Sub cmbRun_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then cmdRun_Click
End Sub
Private Sub cmbRun_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
cmbRun.Text = Data.Files.Item(1)
End Sub
Private Sub cmdRun_Click()
On Error Resume Next
 AdminUser = txtUser.Text
AdminPwd = txtPass.Text
 If cmbRun.Text <> "" Then ShellExe cmbRun.Text
End Sub

Private Sub cmdRunProg_Click(Index As Integer)
Dim progP As String
On Error Resume Next
AdminUser = txtUser.Text
AdminPwd = txtPass.Text
progP = TempString(Index)
If progP = "" Then Exit Sub
Select Case Index
    Case Else
       ShellExe progP
End Select
End Sub
Private Sub cmdRunProg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then  'Right Click
 indexGnrl = Index
 Me.PopupMenu mnAddProg
End If
End Sub

Private Sub cmdRunProg_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim FileGet As String
FileGet = Data.Files.Item(1)
If FileGet <> "" Then
   cmdRunProg(Index).Caption = Dir(FileGet)
   imgIcon(Index).Picture = GetIconExec(FileGet, 0, True)
   TempString(Index) = FileGet
End If
PutSettings
End Sub
Private Sub cmdBrowse_Click()
On Error Resume Next
Dim FileGet As String
    ComDlgBx.lStructSize = Len(ComDlgBx)
    ComDlgBx.hwndOwner = Me.hwnd
    ComDlgBx.hInstance = App.hInstance
    ComDlgBx.lpstrFilter = "Programmes (*.exe)" + Chr$(0) + "*.exe" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    ComDlgBx.lpstrFile = Space$(254)
    ComDlgBx.nMaxFile = 255
    ComDlgBx.lpstrFileTitle = Space$(254)
    ComDlgBx.nMaxFileTitle = 255
    ComDlgBx.lpstrInitialDir = ""
    ComDlgBx.lpstrTitle = "Ovrir un programme"
    ComDlgBx.flags = 0
FileGet = ShowOpen
If FileGet <> "" Then cmbRun.Text = FileGet
End Sub

Private Sub Form_Load()
On Error Resume Next
    Dim sRet As String, lngRet As Long
    If App.PrevInstance Then End: Exit Sub
    
    'Get System Path, Start
    sRet = String$(MAX_PATH, 0)
    lngRet = GetSystemDirectory(sRet, MAX_PATH)
    SystemDir = Left(sRet, lngRet)
    If Right(SystemDir, 1) <> "\" Then SystemDir = SystemDir + "\"
    
    sRet = ""
    sRet = String$(MAX_PATH, 0)
    lngRet = GetWindowsDirectory(sRet, MAX_PATH)
    WindowsDir = Left(sRet, lngRet)
    If Right(WindowsDir, 1) <> "\" Then WindowsDir = WindowsDir + "\"
    imgIcon(0).Picture = GetIconExec(WindowsDir + "explorer.exe", 0, True)
    TempString(0) = "iexplore " + Left(WindowsDir, 2)
    imgIcon(1).Picture = GetIconExec(SystemDir + "filemgmt.dll", 0, True)
    TempString(1) = SystemDir + "services.msc"

    TempString(2) = SystemDir + "compmgmt.msc"
GetSettings

End Sub

 

Private Sub Form_Unload(Cancel As Integer)
PutSettings
End Sub

Private Sub Frame2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
cmbRun.Text = Data.Files.Item(1)
End Sub

Private Sub Add_Prog(Index As Integer)
On Error Resume Next
If Index > 2 Then
  Dim FileGet As String
    ComDlgBx.lStructSize = Len(ComDlgBx)
    'Set the parent window
    ComDlgBx.hwndOwner = Me.hwnd
    'Set the application's instance
    ComDlgBx.hInstance = App.hInstance
    'Select a filter
    ComDlgBx.lpstrFilter = "Programmes (*.exe)" + Chr$(0) + "*.exe" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    'create a buffer for the file
    ComDlgBx.lpstrFile = Space$(254)
    'set the maximum length of a returned file
    ComDlgBx.nMaxFile = 255
    'Create a buffer for the file title
    ComDlgBx.lpstrFileTitle = Space$(254)
    'Set the maximum length of a returned file title
    ComDlgBx.nMaxFileTitle = 255
    'Set the initial directory
    ComDlgBx.lpstrInitialDir = ""
    'Set the title
    ComDlgBx.lpstrTitle = "Ovrir un programme"
    'No flags
    ComDlgBx.flags = 0
    
FileGet = ShowOpen
    If FileGet <> "" Then
        cmdRunProg(Index).Caption = Dir(FileGet)
        imgIcon(Index).Picture = GetIconExec(FileGet, 0, True)
        TempString(Index) = FileGet
        TempImage(Index) = FileGet
    End If
    
End If
End Sub

Private Sub mnIcon_Click()
On Error Resume Next
If indexGnrl > 2 Then
  Dim FileGet As String
    ComDlgBx.lStructSize = Len(ComDlgBx)
    'Set the parent window
    ComDlgBx.hwndOwner = Me.hwnd
    'Set the application's instance
    ComDlgBx.hInstance = App.hInstance
    'Select a filter
    ComDlgBx.lpstrFilter = "*.exe;*.dll;Images" + Chr$(0) + "*.exe;*.wmf;*.dll;*.bmp;*.gif;*.jpg;*.jpeg" + Chr$(0) _
                    + "All Files (*.*)" + Chr$(0) + "*.*"
                    
    'create a buffer for the file
    ComDlgBx.lpstrFile = Space$(254)
    'set the maximum length of a returned file
    ComDlgBx.nMaxFile = 255
    'Create a buffer for the file title
    ComDlgBx.lpstrFileTitle = Space$(254)
    'Set the maximum length of a returned file title
    ComDlgBx.nMaxFileTitle = 255
    'Set the initial directory
    ComDlgBx.lpstrInitialDir = ""
    'Set the title
    ComDlgBx.lpstrTitle = "Ovrir un programme"
    'No flags
    ComDlgBx.flags = 0
    
FileGet = ShowOpen
    If FileGet <> "" Then
        imgIcon(indexGnrl).Picture = GetIconExec(FileGet, 0, True)
        imgIcon(indexGnrl).Picture = LoadPicture(FileGet)
        TempImage(indexGnrl) = FileGet
    End If
    PutSettings
End If
End Sub

Private Sub OpenWith_Click()
Dim Res As Long
Dim ShellStr As String
ShellStr = "rundll32.exe shell32.dll,OpenAs_RunDLL " + TempString(indexGnrl)
If AdminUser <> "" Then
    Res = RunAs(AdminUser, AdminPwd, ShellStr)    'CommandLine
    If Res <> 0 Then MsgBox GetErrorMessage(Res)
Else
    Shell ShellStr, vbNormalFocus
End If
End Sub

Private Sub Suprimer_Click()
On Error Resume Next
        cmdRunProg(indexGnrl).Caption = ""
        imgIcon(indexGnrl).Picture = Nothing
        TempString(indexGnrl) = ""
        TempImage(indexGnrl) = ""
        PutSettings
End Sub

Private Sub txtUser_LostFocus()
PutSettings
End Sub
