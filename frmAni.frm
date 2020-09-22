VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAni 
   Caption         =   "Create Animated Cursor"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7065
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmAni.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   257
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   471
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog ComDia 
      Left            =   3600
      Top             =   3285
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   3375
      Width           =   1140
   End
   Begin VB.PictureBox picPrev 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   4320
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   11
      Top             =   90
      Width           =   540
   End
   Begin VB.TextBox txtInterval 
      Height          =   285
      Left            =   4890
      TabIndex        =   9
      Text            =   "10"
      Top             =   2460
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   3375
      Width           =   1335
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create ANI && Test"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   2880
      Width           =   2685
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "->"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.FileListBox fleFile 
      Height          =   2235
      Left            =   2190
      Pattern         =   "*.cur"
      TabIndex        =   3
      Top             =   90
      Width           =   2055
   End
   Begin VB.DirListBox drDir 
      Height          =   2115
      Left            =   120
      TabIndex        =   2
      Top             =   450
      Width           =   2055
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   315
      Left            =   135
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.ListBox lstFiles 
      Height          =   2205
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   $"frmAni.frx":030A
      Height          =   1050
      Left            =   270
      TabIndex        =   13
      Top             =   2700
      Width           =   3615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "msecs."
      Height          =   195
      Left            =   5430
      TabIndex        =   10
      Top             =   2490
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interval between icons:"
      Height          =   195
      Left            =   3090
      TabIndex        =   8
      Top             =   2490
      Width           =   1650
   End
End
Attribute VB_Name = "frmAni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SystemCursor As Long
Dim vbCursor As Long
Dim inTest As Boolean

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Const AniDesc = "Created using Visual Basic 6 with Neil Crosby's vbIconMaker" + "  " '59  characters +2 spaces
Const AniCreator = " vbIconMaker " '13 characters
Const AppTitle = AniCreator

Private Sub cmdAdd_Click()
    If fleFile.FileName = "" Then
        MsgBox "You must select a file to add !", vbExclamation + vbOKOnly, AppTitle
        Exit Sub
    End If
    Dim CPath As String
    CPath = drDir.Path
    If Right$(CPath, 1) <> "\" Then CPath = CPath + "\"
    CPath = CPath + fleFile.FileName
    lstFiles.AddItem CPath
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdCreate_Click()
Dim cdDir, Pos

On Error GoTo Fout
        Dim Ret As String, nFileSize As Long, nInfoSize As Long, Cnt As Long, sSave As String
    If Val(txtInterval.Text) < 1 Then
        MsgBox "Invalid interval", vbCritical + vbOKOnly, AppTitle
        Exit Sub
    End If
    If lstFiles.ListCount = 0 Then
        MsgBox "Please, select some files !", vbExclamation + vbOKOnly, AppTitle
        Exit Sub
    End If
    '======Get name of file to save==========
ComDia.CancelError = True
ComDia.FileName = "My Ani"
cdDir = GetSetting("vbIconMaker", "ComDiaSettings", "cdAniDirSetting")
If cdDir = "" Then GoTo NoRegVal  'first time
ComDia.InitDir = cdDir
NoRegVal: ComDia.flags = cdlOFNOverwritePrompt + cdlOFNNoReadOnlyReturn
ComDia.Filter = "Animated Cursors (*.ani)|*.ani"
ComDia.ShowSave
Pos = InStrRev(ComDia.FileName, "\")
cdDir = Mid(ComDia.FileName, 1, Pos)
SaveSetting "vbIconMaker", "ComDiaSettings", "cdAniDirSetting", cdDir
'===============
    
    Ret = ComDia.FileName
    
    
    If Ret <> "" Then
        For Cnt = 0 To lstFiles.ListCount - 1
            nFileSize = nFileSize + FileLen(lstFiles.List(Cnt))
        Next
        nFileSize = 97 + Len(AniDesc) + Len(AniCreator) + nFileSize + lstFiles.ListCount * 8
        If Dir(Ret) <> "" Then Kill Ret
        Open Ret For Binary As #1
            Put #1, , "RIFF"
            Put #1, , nFileSize
            Put #1, , "ACONLIST"
            nInfoSize = Len(AniDesc) + Len(AniCreator) + 21 '*2
            Put #1, , nInfoSize
            Put #1, , "INFOINAM"
            nInfoSize = Len(AniDesc) + 1
            Put #1, , nInfoSize
            Put #1, , AniDesc + Chr$(0)
            Put #1, , "IART"
            nInfoSize = Len(AniCreator) + 1
            Put #1, , nInfoSize
            Put #1, , AniCreator + Chr$(0)
            Put #1, , "anih"
            nInfoSize = 36
            Put #1, , nInfoSize
            Put #1, , nInfoSize
            nInfoSize = lstFiles.ListCount
            Put #1, , nInfoSize
            Put #1, , nInfoSize
            sSave = String(16, 0)
            Put #1, , sSave
            nInfoSize = IIf(Fix(Val(txtInterval.Text) / 10) > 0, Fix(Val(txtInterval.Text) / 10), 1)
            Put #1, , nInfoSize
            nInfoSize = 1
            Put #1, , nInfoSize
            Put #1, , "LIST"
            nInfoSize = (nFileSize - (97 + Len(AniDesc) + Len(AniCreator))) + 4
            Put #1, , nInfoSize
            Put #1, , "fram"
            For Cnt = 0 To lstFiles.ListCount - 1
                Put #1, , "icon"
                nInfoSize = FileLen(lstFiles.List(Cnt))
                Put #1, , nInfoSize
                sSave = String(nInfoSize, 0)
                Open lstFiles.List(Cnt) For Binary As #2
                    Get #2, , sSave
                Close #2
                Put #1, , sSave
            Next
        Close
    End If
'==END Saving .ani file=======
inTest = True
vbCursor = LoadCursorFromFile(Ret)

 SetCapture Me.hwnd
 SetCursor vbCursor

Exit Sub

Fout: If Err.Number = 32755 Then Exit Sub 'user pressed cancel
    MsgBox Error$, vbCritical + vbOKOnly, AppTitle
End Sub
Private Sub cmdDelete_Click()
    If lstFiles.ListIndex <> -1 Then lstFiles.RemoveItem lstFiles.ListIndex
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub drDir_Change()
    fleFile.Path = drDir.Path
End Sub

Private Sub drDir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SaveSetting "vbIconMaker", "ComDiaSettings", "drDirSetting", drDir.Path
    SaveSetting "vbIconMaker", "ComDiaSettings", "drvDriveSetting", drvDrive.Drive
End Sub

Private Sub drvDrive_Change()
On Local Error GoTo Fout
    drDir.Path = drvDrive.Drive
Exit Sub

Fout:
    drvDrive.Drive = "C:\"
End Sub
Private Sub fleFile_Click()
On Local Error Resume Next
    Dim CPath As String
    CPath = drDir.Path
    If Right$(CPath, 1) <> "\" Then CPath = CPath + "\"
    CPath = CPath + fleFile.FileName
    picPrev.Picture = LoadPicture(CPath)
End Sub
Private Sub fleFile_DblClick()
    cmdAdd_Click
End Sub

Private Sub Form_Click()
 ReleaseCapture
 SetCursor SystemCursor
 SystemCursor = 0
End Sub

Private Sub Form_Load()
    lstFiles.Clear
    picPrev.Width = 32
    picPrev.Height = 32
    On Error Resume Next
drvDrive.Drive = GetSetting("vbIconMaker", "ComDiaSettings", "drvDriveSetting")
drDir.Path = GetSetting("vbIconMaker", "ComDiaSettings", "drDirSetting")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form1.Show
End Sub

Private Sub Label3_Click()
If inTest Then
 ReleaseCapture
 SetCursor SystemCursor
 SystemCursor = 0
 inTest = False
End If
End Sub
