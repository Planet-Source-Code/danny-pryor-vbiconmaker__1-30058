VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Set Cursor Hotspots"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   Icon            =   "SetCursorHotspots.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4905
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCur 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   555
      Left            =   2520
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   9
      Top             =   180
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picTestArea 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   495
      MousePointer    =   99  'Custom
      ScaleHeight     =   1755
      ScaleWidth      =   3105
      TabIndex        =   7
      Top             =   2025
      Width           =   3165
   End
   Begin VB.TextBox txtX 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1485
      TabIndex        =   4
      Top             =   90
      Width           =   420
   End
   Begin VB.TextBox txtY 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1485
      TabIndex        =   3
      Top             =   450
      Width           =   420
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2250
      TabIndex        =   2
      Top             =   4005
      Width           =   1005
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   945
      TabIndex        =   1
      Top             =   4005
      Width           =   1005
   End
   Begin VB.PictureBox picIco 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   555
      Left            =   405
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   0
      Top             =   180
      Width           =   555
   End
   Begin VB.Label Label2 
      Caption         =   "Click On Image To Set X,Y Hotspots"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   315
      TabIndex        =   11
      Top             =   990
      Width           =   3390
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   4680
      Y1              =   1575
      Y2              =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "VB Cursor files seem to only display in Black and White."
      Height          =   240
      Left            =   180
      TabIndex        =   10
      Top             =   4545
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Hotspot Test Area - Click && Drag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   630
      TabIndex        =   8
      Top             =   1755
      Width           =   2895
   End
   Begin VB.Label lblX 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1215
      TabIndex        =   6
      Top             =   135
      Width           =   195
   End
   Begin VB.Label lblY 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1215
      TabIndex        =   5
      Top             =   495
      Width           =   195
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ImageData() As Byte, IconData(11) As String
Dim canDraw As Boolean, setHot As Boolean

Private Sub cmdCancel_Click()
CancelIt = True
Unload Form3
End Sub



Private Sub cmdOK_Click()
If setHot = True Then
    Unload Form3
    setHot = False
Else
    MsgBox "Hotspots not set."
End If
End Sub

Private Sub Form_Load()
picIco.Picture = LoadPicture(App.Path & "\temp.ico")
canDraw = False
GetIcon
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtX.Text = ""
txtY.Text = ""
End Sub

Private Sub picIco_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtX.Text = X
txtY.Text = Y
End Sub

Private Sub picIco_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
GetCursor
setHot = True
End Sub
Private Sub picTestArea_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
canDraw = True
End Sub
Private Sub picTestArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If canDraw = True Then picTestArea.PSet (X, Y), vbBlue
End Sub
Private Sub picTestArea_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
canDraw = False
End Sub
Private Sub GetCursor()
Dim CurDir As CURSORDIR, CurDirEntry As CURSORDIRENTRY
Dim XHotspot, YHotspot, i
XHotspot = Val(txtX.Text)
YHotspot = Val(txtY.Text)
CurDir.idReserved = Val(IconData(0))
CurDir.idType = 2
CurDir.idCount = Val(IconData(2))
CurDirEntry.bWidth = IconData(3)
CurDirEntry.bHeight = IconData(4)
CurDirEntry.bColorCount = IconData(5)
CurDirEntry.bReserved = IconData(6)
CurDirEntry.wXHotspot = XHotspot
CurDirEntry.wYHotspot = YHotspot
CurDirEntry.dwBytesInRes = IconData(9)
CurDirEntry.dwImageOffset = IconData(10)
Open Form1.ComDia.FileName For Binary As #1
        Put #1, , CurDir
        Put #1, , CurDirEntry
        For i = LBound(ImageData) To UBound(ImageData)
        Put #1, , ImageData(i) 'as byte
        Next i
    Close #1
DoEvents
picCur.Picture = LoadPicture(Form1.ComDia.FileName)
DoEvents
picTestArea.MouseIcon = picCur.Picture
DoEvents
End Sub
Private Sub GetIcon()
 'Create some working variables
  Dim hFile As Integer
  Dim tmp As String
  
 'Create the variables to hold the icon info
 Dim FileHeader As ICONDIR2
 Dim InfoHeader As ICONDIRENTRY2, i, j, iBit As Byte
 
 i = 0
  hFile = FreeFile
  Open App.Path & "\temp.ico" For Binary Access Read As #hFile
    'Read the file header info
    Get #hFile, , FileHeader
    Get #hFile, , InfoHeader
    'Get image data bytes
    For j = 1 To (InfoHeader.dwBytesInRes)
    ReDim Preserve ImageData(i) As Byte
   Get #hFile, , iBit ' Read next byte of image data.
    ImageData(j - 1) = iBit
    i = i + 1
    Next j
  Close #hFile
  IconData(0) = FileHeader.idReserved
  IconData(1) = FileHeader.idType
  IconData(2) = FileHeader.idCount
  IconData(3) = InfoHeader.bWidth
  IconData(4) = InfoHeader.bHeight
  IconData(5) = InfoHeader.bColorCount
  IconData(6) = InfoHeader.bReserved
  IconData(7) = InfoHeader.wPlanes
  IconData(8) = InfoHeader.wBitCount
  IconData(9) = InfoHeader.dwBytesInRes
  IconData(10) = InfoHeader.dwImageOffset
End Sub

