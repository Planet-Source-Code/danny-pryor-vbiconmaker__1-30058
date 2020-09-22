VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ViewAni 
   Caption         =   "View .ani Files"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
   Icon            =   "ViewAni.frx":0000
   LinkTopic       =   "ViewAni"
   ScaleHeight     =   3825
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog ComDia 
      Left            =   1125
      Top             =   2745
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   420
      Left            =   3870
      TabIndex        =   1
      Top             =   3375
      Width           =   1290
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open .ani file"
      Height          =   435
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   1380
   End
End
Attribute VB_Name = "ViewAni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Dim SystemCursor As Long
Dim vbCursor As Long
Dim inView As Boolean

Private Sub cmdOpen_Click()
Dim cdDir, cdIndex, Pos
On Error GoTo ex
'==========
cdDir = GetSetting("vbIconMaker", "ComDiaSettings", "cdViewAniDirSetting")
cdIndex = GetSetting("vbIconMaker", "ComDiaSettings", "cdViewAniIndexSetting")
If cdDir = "" Or cdIndex = "" Then GoTo NoRegVal 'first time
ComDia.FilterIndex = cdIndex
ComDia.InitDir = cdDir
'===========
NoRegVal: ComDia.CancelError = True
ComDia.FileName = ""
ComDia.flags = cdlOFNFileMustExist
ComDia.Filter = "Animated Cursors (*.ani)|*.ani"
ComDia.ShowOpen
'============
Pos = InStrRev(ComDia.FileName, "\")
cdDir = Mid(ComDia.FileName, 1, Pos)
cdIndex = ComDia.FilterIndex
SaveSetting "vbIconMaker", "ComDiaSettings", "cdViewAniDirSetting", cdDir
SaveSetting "vbIconMaker", "ComDiaSettings", "cdViewAniIndexSetting", cdIndex
'============
inView = True
vbCursor = LoadCursorFromFile(ComDia.FileName)
SetCapture Me.hwnd
 SetCursor vbCursor
Exit Sub
ex: If Err.Number = 32755 Then Exit Sub 'user pressed cancel
MsgBox "Error # " & Err.Number & " - " & Err.Description
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Click()
If inView Then
 ReleaseCapture
 SetCursor SystemCursor
 SystemCursor = 0
 inView = False
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form1.Show
End Sub
