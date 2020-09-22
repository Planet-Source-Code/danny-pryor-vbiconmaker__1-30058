VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Form1 
   Caption         =   "VB Icon Maker-32X32 Pixel"
   ClientHeight    =   6615
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9090
   FillStyle       =   0  'Solid
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MouseIcon       =   "Form1.frx":030A
   ScaleHeight     =   441
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   606
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCBsprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   5985
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   48
      Top             =   4905
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picCBmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   5985
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   47
      Top             =   4170
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picCB 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   6015
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   45
      Top             =   3585
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   7755
      LinkTimeout     =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   6735
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton cmdRedo 
      Caption         =   "Redo"
      Height          =   390
      Left            =   90
      TabIndex        =   43
      Top             =   4140
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "Undo"
      Height          =   390
      Left            =   90
      TabIndex        =   42
      Top             =   3660
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   6120
      LinkTimeout     =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   6750
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   5595
      LinkTimeout     =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   6780
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicIcon 
      AutoRedraw      =   -1  'True
      DragIcon        =   "Form1.frx":074C
      Height          =   480
      Left            =   5025
      LinkTimeout     =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   6780
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picReal16 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2820
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   37
      Top             =   90
      Width           =   240
   End
   Begin VB.PictureBox picHand 
      Height          =   420
      Left            =   4845
      Picture         =   "Form1.frx":089E
      ScaleHeight     =   360
      ScaleWidth      =   405
      TabIndex        =   36
      Top             =   285
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4050
      Top             =   330
   End
   Begin VB.CommandButton cmdRegion 
      Height          =   330
      Left            =   450
      Picture         =   "Form1.frx":09F0
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Select Region & Clipboard Functions"
      Top             =   2520
      Width           =   330
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   135
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   34
      Top             =   5085
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picA 
      Height          =   375
      Left            =   5760
      Picture         =   "Form1.frx":0D91
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   33
      Top             =   45
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdText 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Text"
      Top             =   2520
      Width           =   330
   End
   Begin VB.PictureBox picFlood 
      Height          =   375
      Left            =   5355
      Picture         =   "Form1.frx":0EE3
      ScaleHeight     =   315
      ScaleWidth      =   360
      TabIndex        =   31
      Top             =   315
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CommandButton cmdCircleDraw 
      Height          =   330
      Left            =   450
      Picture         =   "Form1.frx":1035
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Ellipse"
      Top             =   1710
      Width           =   330
   End
   Begin VB.CommandButton cmdFillCircleDraw 
      Height          =   330
      Left            =   45
      Picture         =   "Form1.frx":1737
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Filled Ellipse"
      Top             =   1710
      Width           =   330
   End
   Begin VB.CommandButton cmdFillBox 
      Height          =   330
      Left            =   45
      Picture         =   "Form1.frx":1E39
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Filled Rectangle"
      Top             =   2115
      Width           =   330
   End
   Begin VB.CommandButton cmdRect 
      Height          =   330
      Left            =   450
      Picture         =   "Form1.frx":24BB
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Rectangle"
      Top             =   2115
      Width           =   330
   End
   Begin VB.CommandButton cmdLine 
      Height          =   330
      Left            =   45
      Picture         =   "Form1.frx":2B3D
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Line"
      Top             =   1305
      Width           =   330
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "More Colors"
      Height          =   360
      Left            =   6885
      TabIndex        =   21
      Top             =   4815
      Width           =   1140
   End
   Begin VB.PictureBox picBasic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      DrawWidth       =   16
      Height          =   1710
      Left            =   6300
      MousePointer    =   99  'Custom
      ScaleHeight     =   114
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   20
      Top             =   3045
      Width           =   2250
   End
   Begin VB.PictureBox picColorpicker 
      Height          =   330
      Left            =   330
      Picture         =   "Form1.frx":31BF
      ScaleHeight     =   270
      ScaleWidth      =   225
      TabIndex        =   19
      Top             =   5670
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picEraser 
      Height          =   240
      Left            =   15
      Picture         =   "Form1.frx":3311
      ScaleHeight     =   180
      ScaleWidth      =   225
      TabIndex        =   18
      Top             =   5835
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picPencil 
      Height          =   330
      Left            =   30
      Picture         =   "Form1.frx":3463
      ScaleHeight     =   270
      ScaleWidth      =   180
      TabIndex        =   17
      Top             =   5475
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picTransparent 
      Height          =   285
      Left            =   7035
      ScaleHeight     =   225
      ScaleWidth      =   270
      TabIndex        =   15
      Top             =   180
      Width           =   330
   End
   Begin VB.CommandButton cmdPicker 
      Caption         =   "Pick Color From Image"
      Height          =   390
      Left            =   6525
      TabIndex        =   14
      Top             =   5265
      Width           =   1845
   End
   Begin VB.PictureBox pic16Color 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      DrawWidth       =   16
      FillColor       =   &H00FFFFFF&
      FillStyle       =   2  'Horizontal Line
      Height          =   645
      Left            =   6315
      MousePointer    =   99  'Custom
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   13
      Top             =   2040
      Width           =   2250
   End
   Begin VB.CommandButton cmdErase 
      Height          =   330
      Left            =   450
      Picture         =   "Form1.frx":35B5
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Eraser"
      Top             =   900
      Width           =   330
   End
   Begin VB.CommandButton cmdFlood 
      Height          =   330
      Left            =   450
      Picture         =   "Form1.frx":363F
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Flood"
      Top             =   1305
      Width           =   330
   End
   Begin VB.CommandButton cmdPaint 
      Height          =   330
      Left            =   45
      Picture         =   "Form1.frx":36DC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Pixel Paint"
      Top             =   900
      Width           =   330
   End
   Begin VB.PictureBox PicMseColor 
      BackColor       =   &H000000FF&
      Height          =   330
      Left            =   7965
      ScaleHeight     =   270
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   810
      Width           =   375
   End
   Begin MSComDlg.CommonDialog ComDia 
      Left            =   75
      Top             =   2970
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picReal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3150
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   6
      ToolTipText     =   "Dbl Clik for 16x16 view"
      Top             =   90
      Width           =   480
   End
   Begin VB.PictureBox PicMseColorR 
      BackColor       =   &H0000FFFF&
      Height          =   330
      Left            =   8190
      ScaleHeight     =   270
      ScaleWidth      =   315
      TabIndex        =   26
      Top             =   900
      Width           =   375
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   7  'Diagonal Cross
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   1035
      MouseIcon       =   "Form1.frx":3766
      MousePointer    =   99  'Custom
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   1
      Top             =   855
      Width           =   4815
      Begin VB.PictureBox picMove 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2220
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   46
         Top             =   1290
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.PictureBox picTest 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   765
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   46
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   690
      End
      Begin PicClip.PictureClip PicClipMove 
         Left            =   450
         Top             =   1380
         _ExtentX        =   582
         _ExtentY        =   503
         _Version        =   393216
      End
      Begin PicClip.PictureClip picClip 
         Left            =   0
         Top             =   0
         _ExtentX        =   582
         _ExtentY        =   503
         _Version        =   393216
      End
      Begin VB.Shape shRect 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Height          =   420
         Left            =   765
         Top             =   945
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   10
         FillColor       =   &H00C0E0FF&
         Height          =   420
         Left            =   90
         Shape           =   2  'Oval
         Top             =   1755
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0000C000&
         BorderWidth     =   10
         Visible         =   0   'False
         X1              =   123
         X2              =   183
         Y1              =   12
         Y2              =   12
      End
   End
   Begin VB.Image imgSaveDown 
      Height          =   360
      Left            =   870
      Picture         =   "Form1.frx":38B8
      Top             =   90
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgSaveOver 
      Height          =   360
      Left            =   870
      Picture         =   "Form1.frx":3FBA
      ToolTipText     =   "Save As"
      Top             =   90
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgOpenDown 
      Height          =   360
      Left            =   480
      Picture         =   "Form1.frx":46BC
      Top             =   90
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgOpenOver 
      Height          =   360
      Left            =   480
      Picture         =   "Form1.frx":4DBE
      ToolTipText     =   "Open"
      Top             =   90
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgNewDown 
      Height          =   360
      Left            =   90
      Picture         =   "Form1.frx":54C0
      Top             =   90
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgNewOver 
      Height          =   360
      Left            =   90
      Picture         =   "Form1.frx":5B42
      ToolTipText     =   "New"
      Top             =   90
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblMsePos 
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
      Left            =   990
      TabIndex        =   41
      Top             =   5835
      Width           =   2190
   End
   Begin VB.Shape shSel 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   4
      Height          =   420
      Left            =   0
      Top             =   855
      Width           =   420
   End
   Begin VB.Image imgSave 
      Height          =   360
      Left            =   870
      Picture         =   "Form1.frx":61C4
      ToolTipText     =   "Save As"
      Top             =   90
      Width           =   360
   End
   Begin VB.Image imgOpen 
      Height          =   360
      Left            =   480
      Picture         =   "Form1.frx":6366
      ToolTipText     =   "Open"
      Top             =   90
      Width           =   360
   End
   Begin VB.Image imgNew 
      Height          =   345
      Left            =   90
      Picture         =   "Form1.frx":6508
      ToolTipText     =   "New"
      Top             =   90
      Width           =   360
   End
   Begin VB.Label lblRGB 
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
      Left            =   6120
      TabIndex        =   24
      Top             =   5760
      Width           =   2940
   End
   Begin VB.Label Label4 
      Caption         =   "Click a color below to Change Current Paint Color."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6165
      TabIndex        =   23
      Top             =   1350
      Width           =   2715
   End
   Begin VB.Label Label4 
      Caption         =   "Some Basic Colors."
      Height          =   195
      Index           =   1
      Left            =   6375
      TabIndex        =   22
      Top             =   2835
      Width           =   1680
   End
   Begin VB.Label Label5 
      Caption         =   "The Color             Will be transparent in the Saved Icon. RGB(197, 197, 197)"
      Height          =   510
      Left            =   6210
      TabIndex        =   16
      Top             =   270
      Width           =   2850
   End
   Begin VB.Label Label4 
      Caption         =   "The 16 Named Colors."
      Height          =   240
      Index           =   0
      Left            =   6345
      TabIndex        =   12
      Top             =   1800
      Width           =   1680
   End
   Begin VB.Label lblPath 
      Caption         =   "Untitled"
      Height          =   285
      Left            =   1710
      TabIndex        =   9
      Top             =   6210
      Width           =   7305
   End
   Begin VB.Label Label3 
      Caption         =   "Expanded View for Editing"
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
      Left            =   3600
      TabIndex        =   11
      Top             =   5835
      Width           =   2355
   End
   Begin VB.Label Label2 
      Caption         =   "Edit View of Image"
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
      Left            =   3690
      TabIndex        =   10
      Top             =   45
      Width           =   1770
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   57
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblEdit 
      Caption         =   "Icon being Edited:"
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
      Left            =   90
      TabIndex        =   8
      Top             =   6210
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5865
      Picture         =   "Form1.frx":6B72
      Top             =   2445
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Current Paint Color"
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
      Left            =   6165
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   10
      Height          =   330
      Left            =   1815
      Top             =   360
      Width           =   315
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Existing Icon or Image File"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save As"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu MnuEditOpts 
         Caption         =   "&Undo"
         Index           =   0
         Shortcut        =   ^Z
      End
      Begin VB.Menu MnuEditOpts 
         Caption         =   "&Redo"
         Index           =   1
      End
      Begin VB.Menu MnuEditOpts 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu MnuEditOpts 
         Caption         =   "Cu&t"
         Index           =   3
         Shortcut        =   ^X
      End
      Begin VB.Menu MnuEditOpts 
         Caption         =   "&Copy"
         Index           =   4
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuEditOpts 
         Caption         =   "&Paste"
         Index           =   5
         Shortcut        =   ^V
      End
      Begin VB.Menu MnuEditOpts 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu MnuEditOpts 
         Caption         =   "&Delete"
         Index           =   7
         Shortcut        =   {DEL}
      End
      Begin VB.Menu MnuEditOpts 
         Caption         =   "C&lear Clipboard"
         Index           =   8
      End
      Begin VB.Menu MnuEditOpts 
         Caption         =   "C&ancel Select"
         Index           =   9
      End
   End
   Begin VB.Menu mnuOther 
      Caption         =   "&Other Options"
      Begin VB.Menu mnuExtract 
         Caption         =   "&Extract Icon from a file"
      End
      Begin VB.Menu mnuAni 
         Caption         =   "Create &Animated Cursor"
      End
      Begin VB.Menu nmuAni 
         Caption         =   "&View .ani File"
      End
      Begin VB.Menu mnuChgPix 
         Caption         =   "&Change Pixel to Transparent"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAssoc 
         Caption         =   "A&ssociate program with ico files"
      End
   End
   Begin VB.Menu nmuEffects 
      Caption         =   "E&ffects"
      Begin VB.Menu mnuHorz 
         Caption         =   "Flip &Horizonally(Mirror)"
      End
      Begin VB.Menu mnuVert 
         Caption         =   "Flip &Vertically"
      End
      Begin VB.Menu mnuRotateRight 
         Caption         =   "Rotate 90 Degrees Right"
      End
      Begin VB.Menu mnuRotateLeft 
         Caption         =   "Rotate 90 Degrees &Left"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuPopCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPopPaste 
         Caption         =   "&Paste"
         Visible         =   0   'False
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuPopCancel 
         Caption         =   "C&ancel Select"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim okToMove As Boolean, prevX%, prevY%
Dim xSav, ySav, xStart, yStart
Dim pX%, pY%, pXOff%, pYOff%
'==========
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'==========
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private CurrentFile$
Private CurrentName$
Dim cvtY, cvtX
Dim filePath
Dim pixelDraw As Boolean, canDraw As Boolean ', Dirty As Boolean
Dim ColorChg, bkClr, clr As Long, r As Integer, g As Integer, b As Integer
Dim j, p, x1, y1, colorSave, eraseIt As Boolean, pickColor As Boolean
Dim chkPix As Boolean, chgColor 'As Integer
Dim lineDraw As Boolean, lineOKDraw As Boolean
Dim rectDraw As Boolean, fillBoxDraw As Boolean, rectOKDraw As Boolean
Dim circleDraw As Boolean, fillCircleDraw As Boolean, circleOKDraw As Boolean
Dim textDraw As Boolean
Dim selRegion As Boolean
Dim lineX1, lineY1
Dim pasteIt As Boolean
'==For Select Region==================
Dim XHi, XLo, YHi, YLo, xDelLo, xdelHi, yDelLo, ydelHi
''Dim canMove As Boolean
Dim canSelect As Boolean
Dim moveIt As Boolean
Dim selectIt As Boolean
Dim xOff, yOff, setDiff As Boolean
Dim xMove, yMove

'========Used with Flood Area===================
Dim floodDraw As Boolean
Private Declare Function ExtFloodFill Lib "gdi32.dll" (ByVal hdc As Long, ByVal nXStart As Long, ByVal nYStart As Long, ByVal crColor As Long, ByVal fuFillType As Long) As Long
Const FLOODFILLBORDER = 0
Const FLOODFILLSURFACE = 1
'===END Flood Area Data==================
'=======Used to Flip Image==========
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Const SRCCOPY = &HCC0020

Private Sub cmdMore_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If chkPix = True Then
MsgBox "Click on pixel in Expanded View"
Exit Sub
End If
   ComDia.CancelError = True
   On Error GoTo ErrHandler
   ComDia.flags = cdlCCFullOpen
   ComDia.ShowColor
   If ComDia.Color = RGB(197, 197, 197) Then ComDia.Color = RGB(196, 196, 196)
If Button = 1 Then
PicMseColor.BackColor = ComDia.Color
Else
PicMseColorR.BackColor = ComDia.Color
End If
eraseIt = False
   Exit Sub
ErrHandler:
   ' User pressed Cancel button.
   Exit Sub

End Sub

Private Sub cmdRedo_Click()
DoReDo
picContainer.SetFocus
End Sub



Private Sub Form_Initialize()
Dim r, c, i, j
Dim ColorRect As RECT
  ''  picPaste.Top = 4
  ''  picPaste.Left = 4
cvtX = Screen.TwipsPerPixelX
cvtY = Screen.TwipsPerPixelY
''Clipboard.Clear
Form1.Caption = "VB Icon Maker Version " & App.Major & "." & App.Minor & " - True Color 32X32 Pixel"
chgColor = 0
picReal.ScaleMode = 3
picReal.ScaleHeight = 32
picReal.ScaleWidth = 32
picReal.Height = 32
picReal.Width = 32
picReal.Top = 6
picReal.Left = 211
picContainer.ScaleMode = 3
picContainer.ScaleHeight = 321
picContainer.ScaleWidth = 321
picContainer.Height = 321
picContainer.Width = 321
picContainer.Top = 57
picContainer.Left = 69
Shape2.Left = 64
Shape2.Top = 52
Shape2.Width = 331
Shape2.Height = 331
Line1.X2 = Form1.ScaleWidth
Set pic16Color.MouseIcon = picColorpicker.Picture
Set picBasic.MouseIcon = picColorpicker.Picture
'draw named colors
j = 0
For r = 1 To 2
        i = 1
    For c = 10 To 136 Step 18
        pic16Color.Line ((i * 18) - 8, ((r - 1) * 18) + 10)-((i * 18) - 8, ((r - 1) * 18) + 10), QBColor(j), BF
        i = i + 1
        j = j + 1
    Next c
Next r

'draw some basic colors
picBasic.DrawWidth = 16
picBasic.Line (10, 10)-(10, 10), RGB(255, 128, 128), BF
picBasic.Line (10, 28)-(10, 28), RGB(255, 0, 0), BF
picBasic.Line (28, 10)-(28, 10), RGB(255, 255, 128), BF
picBasic.Line (28, 28)-(28, 28), RGB(128, 255, 0), BF

picBasic.Line (46, 10)-(46, 10), RGB(0, 255, 128), BF
picBasic.Line (46, 28)-(46, 28), RGB(0, 255, 64), BF
picBasic.Line (64, 10)-(64, 10), RGB(128, 255, 255), BF
picBasic.Line (64, 28)-(64, 28), RGB(0, 255, 255), BF

picBasic.Line (82, 10)-(82, 10), RGB(0, 128, 255), BF
picBasic.Line (82, 28)-(82, 28), RGB(0, 128, 192), BF
picBasic.Line (100, 10)-(100, 10), RGB(255, 128, 192), BF
picBasic.Line (100, 28)-(100, 28), RGB(128, 128, 192), BF

picBasic.Line (118, 10)-(118, 10), RGB(255, 128, 255), BF
picBasic.Line (118, 28)-(118, 28), RGB(255, 0, 255), BF
picBasic.Line (136, 10)-(136, 10), RGB(128, 64, 64), BF
picBasic.Line (136, 28)-(136, 28), RGB(128, 0, 0), BF
'============
picBasic.Line (10, 46)-(10, 46), RGB(255, 128, 64), BF
picBasic.Line (10, 64)-(10, 64), RGB(255, 128, 0), BF
picBasic.Line (28, 46)-(28, 46), RGB(0, 255, 0), BF
picBasic.Line (28, 64)-(28, 64), RGB(0, 128, 0), BF

picBasic.Line (46, 46)-(46, 46), RGB(0, 128, 128), BF
picBasic.Line (46, 64)-(46, 64), RGB(0, 128, 128), BF
picBasic.Line (64, 46)-(64, 46), RGB(0, 64, 128), BF
picBasic.Line (64, 64)-(64, 64), RGB(0, 0, 255), BF

picBasic.Line (82, 46)-(82, 46), RGB(128, 128, 255), BF
picBasic.Line (82, 64)-(82, 64), RGB(0, 0, 160), BF
picBasic.Line (100, 46)-(100, 46), RGB(128, 0, 64), BF
picBasic.Line (100, 64)-(100, 64), RGB(128, 0, 128), BF

picBasic.Line (118, 46)-(118, 46), RGB(255, 0, 128), BF
picBasic.Line (118, 64)-(118, 64), RGB(128, 0, 255), BF
picBasic.Line (136, 46)-(136, 46), RGB(64, 0, 0), BF
picBasic.Line (136, 64)-(136, 64), RGB(0, 0, 0), BF
'=========
picBasic.Line (10, 82)-(10, 82), RGB(128, 64, 0), BF
picBasic.Line (10, 100)-(10, 100), RGB(128, 128, 0), BF
picBasic.Line (28, 82)-(28, 82), RGB(0, 64, 0), BF
picBasic.Line (28, 100)-(28, 100), RGB(128, 128, 64), BF

picBasic.Line (46, 82)-(46, 82), RGB(0, 64, 64), BF
picBasic.Line (46, 100)-(46, 100), RGB(128, 128, 128), BF
picBasic.Line (64, 82)-(64, 82), RGB(0, 0, 128), BF
picBasic.Line (64, 100)-(64, 100), RGB(64, 128, 128), BF

picBasic.Line (82, 82)-(82, 82), RGB(0, 0, 64), BF
picBasic.Line (82, 100)-(82, 100), RGB(64, 0, 64), BF
picBasic.Line (100, 82)-(100, 82), RGB(64, 0, 64), BF
picBasic.Line (100, 100)-(100, 100), RGB(64, 0, 128), BF

picBasic.Line (118, 82)-(118, 82), RGB(232, 144, 56), BF
picBasic.Line (118, 100)-(118, 100), RGB(255, 204, 153), BF
picBasic.Line (136, 82)-(136, 82), RGB(102, 51, 0), BF
picBasic.Line (136, 100)-(136, 100), RGB(18, 201, 55), BF
'With Form1.pic16Color
For r = 1 To 2
    For c = 0 To 143 Step 18
        SetRect ColorRect, (c + 2), (r * 18) - 16, c + 18, ((r - 1) * 18) + 18
        DrawEdge pic16Color.hdc, ColorRect, 2, 15
    Next c
Next r
'With Form1.picBasic
For r = 1 To 6
    For c = 0 To 143 Step 18
        SetRect ColorRect, (c + 2), (r * 18) - 16, c + 18, ((r - 1) * 18) + 18
        DrawEdge picBasic.hdc, ColorRect, 2, 15
    Next c
Next r
End Sub

Private Sub Form_Load()
Dim F, testExt, Answ
CheckSettings
PrepIconHeader

picTransparent.BackColor = RGB(197, 197, 197)
picReal.BackColor = RGB(197, 197, 197)
bkClr = RGB(197, 197, 197)
picContainer.BackColor = RGB(197, 197, 197) '&H8000000F
For F = 0 To picContainer.ScaleHeight Step 10
    picContainer.Line (0, F)-(picContainer.ScaleWidth, F), &H4040&
Next F
For F = 0 To picContainer.ScaleWidth Step 10
    picContainer.Line (F, 0)-(F, picContainer.ScaleHeight), &H4040&
Next F
picReal.Picture = Image1.Picture
PaintDown
Form1.Refresh
UpdateUndo
'*****Add Code to allow for Drag and Drop and Associations
If Command$ <> "" Then
filePath = Command
If LCase(Left(Command, 6)) = "/open " Then
       filePath = StrConv(Mid(Command, 7), vbProperCase)
End If
filePath = LongFileName(filePath)
picReal.Picture = LoadPicture()
lblPath.Caption = filePath
testExt = Mid(filePath, Len(filePath) - 2, 3)
    If LCase(testExt) = "ico" Or LCase(testExt) = "jpg" Or LCase(testExt) = "gif" Or LCase(testExt) = "bmp" Or LCase(testExt) = "cur" Then
        MousePointer = 11
        picReal.BackColor = RGB(197, 197, 197)
        picTest.Picture = LoadPicture(filePath)
        DoEvents
            If picTest.ScaleWidth > 32 Or picTest.ScaleHeight > 32 Then
                    ComDia.FileName = filePath
                    picReal.Picture = LoadPicture()
                    frmScroll.Show '1, Me
                    While Not iDone
                        DoEvents
                    Wend
                    iDone = False
            Else
                picReal = LoadPicture(filePath)
            End If
                PaintDown
                    cmdUndo.Visible = False
        DeleteCollections
                UpdateUndo
                Form1.Refresh
                MousePointer = 0
                lblPath.Caption = filePath
    Else
        ExtractRequest
    End If
End If
Form1.Show
cmdPaint_Click
Dirty = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRGB.Caption = ""
lblMsePos.Caption = ""
imgNew.Visible = True
imgNewOver.Visible = False
imgNewDown.Visible = False
imgOpen.Visible = True
imgOpenOver.Visible = False
imgOpenDown.Visible = False
imgSave.Visible = True
imgSaveOver.Visible = False
imgSaveDown.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim reply


If Dirty = True Then
        reply = MsgBox("Do you wish to save current image?", vbYesNoCancel, "File Not Saved")
    If reply = vbCancel Then Cancel = 1 'Stop program from closing
    If reply = vbYes Then mnuSave_Click
End If
    If Dir(App.Path & "\temp.ico") <> "" Then Kill App.Path & "\temp.ico"
'==Be sure all forms are unloaded
Unload Form2
Unload Form3
Unload Form4
Unload frmAni
Unload ViewAni
Unload MDIForm1
'=========================
If chkPix = True Then End
End Sub

Private Sub cmdPaint_Click()
'=Initialize Switches=========
setSwitchesFalse
pixelDraw = True
'=====================
shSel.Left = cmdPaint.Left - 2
shSel.Top = cmdPaint.Top - 2
picContainer.SetFocus
picContainer.MouseIcon = picPencil.Picture
End Sub

Private Sub cmdErase_Click()
'=Initialize Switches=========
setSwitchesFalse
eraseIt = True
pixelDraw = True
'=====================
shSel.Left = cmdErase.Left - 2
shSel.Top = cmdErase.Top - 2
picContainer.MouseIcon = picEraser.Picture
picContainer.SetFocus
End Sub

Private Sub cmdLine_Click()
'=Initialize Switches=========
setSwitchesFalse
lineDraw = True
'=====================
shSel.Left = cmdLine.Left - 2
shSel.Top = cmdLine.Top - 2
picContainer.SetFocus
picContainer.MouseIcon = picPencil.Picture
End Sub

Private Sub cmdFlood_Click()
'=Initialize Switches=========
setSwitchesFalse
floodDraw = True
'=====================
shSel.Left = cmdFlood.Left - 2
shSel.Top = cmdFlood.Top - 2
picContainer.SetFocus
picContainer.MouseIcon = picFlood.Picture
End Sub

Private Sub cmdCircleDraw_Click()
'=Initialize Switches=========
setSwitchesFalse
circleDraw = True
'=====================
shSel.Left = cmdCircleDraw.Left - 2
shSel.Top = cmdCircleDraw.Top - 2
Shape1.Shape = 2 'Oval
Shape1.FillStyle = 1 'Transparent
picReal.FillStyle = 1 'transparent
picContainer.SetFocus
picContainer.MouseIcon = picPencil.Picture
picContainer.FillStyle = 1 'transparent
End Sub

Private Sub cmdFillCircleDraw_Click()
'=Initialize Switches=========
setSwitchesFalse
fillCircleDraw = True
'=====================
shSel.Left = cmdFillCircleDraw.Left - 2
shSel.Top = cmdFillCircleDraw.Top - 2
Shape1.Shape = 2 'Oval
Shape1.FillStyle = 0 'Solid
picReal.FillStyle = 0 'Solid
picContainer.FillStyle = 0 'Solid
picContainer.SetFocus
picContainer.MouseIcon = picPencil.Picture
End Sub

Private Sub cmdFillBox_Click()
'=Initialize Switches=========
setSwitchesFalse
fillBoxDraw = True
'=====================
shSel.Left = cmdFillBox.Left - 2
shSel.Top = cmdFillBox.Top - 2
Shape1.Shape = 0 'Rectangle
Shape1.FillStyle = 0 'Solid
picReal.FillStyle = 0 'Solid
picContainer.FillStyle = 0 'Solid
picContainer.SetFocus
picContainer.MouseIcon = picPencil.Picture
End Sub

Private Sub cmdRect_Click()
'=Initialize Switches=========
setSwitchesFalse
rectDraw = True
'=====================
shSel.Left = cmdRect.Left - 2
shSel.Top = cmdRect.Top - 2
Shape1.Shape = 0 'Rectangle
Shape1.FillStyle = 1 'Transparent
picReal.FillStyle = 1 'transparent
picContainer.SetFocus
picContainer.MouseIcon = picPencil.Picture
picContainer.FillStyle = 1 'transparent

End Sub

Private Sub cmdText_Click()
'=Initialize Switches=========
setSwitchesFalse
textDraw = True
'=====================
shSel.Left = cmdText.Left - 2
shSel.Top = cmdText.Top - 2
picContainer.SetFocus
picContainer.MouseIcon = picA.Picture
End Sub

Private Sub cmdRegion_Click()
setSwitchesFalse
picContainer.MouseIcon = picPencil.Picture
shSel.Left = cmdRegion.Left - 2
shSel.Top = cmdRegion.Top - 2
selRegion = True
Timer1.Enabled = True
moveIt = False
selectIt = True

End Sub

Private Sub cmdUndo_Click()
DoUnDo
picContainer.SetFocus
End Sub


Private Sub imgNew_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgNew.Visible = False
imgNewOver.Visible = True
End Sub
Private Sub imgNewOver_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgNewDown.Visible = True
End Sub
Private Sub imgNewOver_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgOpenOver.Visible = False
imgSaveOver.Visible = False
imgOpen.Visible = True
imgSave.Visible = True
End Sub

Private Sub imgNewOver_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Allows user to change mind.
If Y / cvtY < 0 Or Y / cvtY > 24 Or X / cvtX < 0 Or X / cvtX > 24 Then Exit Sub
mnuNew_Click
imgNewDown.Visible = False
picContainer.SetFocus
End Sub

Private Sub imgOpen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgOpen.Visible = False
imgOpenOver.Visible = True
End Sub

Private Sub imgOpenOver_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgOpenDown.Visible = True
End Sub

Private Sub imgOpenOver_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgOpen.Visible = False
imgOpenOver.Visible = True
imgNewOver.Visible = False
imgSaveOver.Visible = False
imgSave.Visible = True
imgNew.Visible = True
End Sub

Private Sub imgOpenOver_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y / cvtY < 0 Or Y / cvtY > 24 Or X / cvtX < 0 Or X / cvtX > 24 Then Exit Sub
mnuOpen_Click
imgOpenDown.Visible = False
picContainer.SetFocus
End Sub

Private Sub imgSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSave.Visible = False
imgSaveOver.Visible = True
End Sub

Private Sub imgSaveOver_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSaveDown.Visible = True
End Sub

Private Sub imgSaveOver_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSave.Visible = False
imgSaveOver.Visible = True
imgOpenOver.Visible = False
imgNewOver.Visible = False
imgOpen.Visible = True
imgNew.Visible = True
End Sub

Private Sub imgSaveOver_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y / cvtY < 0 Or Y / cvtY > 24 Or X / cvtX < 0 Or X / cvtX > 24 Then Exit Sub
mnuSave_Click
imgSaveDown.Visible = False
picContainer.SetFocus
End Sub

Private Sub mnuAni_Click()
Form1.Hide
frmAni.Show
End Sub

Private Sub mnuAssoc_Click()
Dim Ret
Ret = MsgBox("This option will associate Window Icon (.ico) files with this program so that double clicking any ico file will open it in this program. Is that what you want to do?", vbYesNo, "Associate icon files")
If Ret = vbNo Then Exit Sub
SetUpIconDblClick
End Sub

Private Sub mnuChgPix_Click()
Dim X, Y, Ret
lineDraw = False
Ret = MsgBox("Click On Pixel Color In Expanded View To Be Made Transparent. ", vbOKCancel, "Change Pixel Color")
If Ret = vbCancel Then Exit Sub
chkPix = True
'=Initialize Switches=========
setSwitchesFalse
'=====================
While chgColor = 0
DoEvents
Wend
MousePointer = 11
picReal.Picture = picReal.Image
For X = 0 To 31
For Y = 0 To 31
If picReal.Point(X, Y) = chgColor Then
picReal.PSet (X, Y), RGB(197, 197, 197)
End If
Next Y
Next X
picReal.Picture = picReal.Image
PaintDown
Form1.Refresh
cmdPaint_Click
MousePointer = 0
chgColor = 0
chkPix = False
Dirty = True
End Sub

Private Sub mnuDelete_Click()
moveIt = True
selectIt = False
End Sub

Private Sub MnuEdit_Click()
    MnuEditOpts(0).Enabled = ColUndo.Count > 1
    MnuEditOpts(1).Enabled = ColRedo.Count > 0
    MnuEditOpts(3).Enabled = shRect.Visible
    MnuEditOpts(4).Enabled = shRect.Visible
    MnuEditOpts(5).Enabled = Clipboard.GetFormat(vbCFBitmap) 'shRect.Visible
    MnuEditOpts(7).Enabled = shRect.Visible
    MnuEditOpts(8).Enabled = Clipboard.GetFormat(vbCFBitmap) 'shRect.Visible
    MnuEditOpts(9).Enabled = shRect.Visible
End Sub

Private Sub MnuEditOpts_Click(Idx%)

    Select Case Idx
           Case 0
                DoUnDo
           Case 1
                DoReDo
           Case 3
                mnuPopCut_Click
           Case 4
                mnuPopCopy_Click
           Case 5
                pasteItNow
           Case 7
                mnuPopDelete_Click
           Case 8
                Clipboard.Clear
           Case 9
                cmdRegion_Click
    End Select
End Sub

Private Sub mnuNew_Click()
chkSave
Dim F

'=======ClearUndo
    UpdateUndo
    cmdUndo.Visible = False
        DeleteCollections
'=Initialize Switches=========
setSwitchesFalse
'=====================
picReal.Picture = LoadPicture()
picTest.Picture = LoadPicture()
picContainer.Picture = LoadPicture()
picContainer.Cls
PaintDown
picReal.BackColor = RGB(197, 197, 197)
picContainer.BackColor = RGB(197, 197, 197)
For F = 0 To picContainer.ScaleHeight Step 10
picContainer.Line (0, F)-(picContainer.ScaleWidth, F), &H4040&
Next F
For F = 0 To picContainer.ScaleWidth Step 10
picContainer.Line (F, 0)-(F, picContainer.ScaleHeight), &H4040&
Next F
lblPath.Caption = "Untitled"
Form1.Refresh
cmdPaint_Click
UpdateUndo
Dirty = False
End Sub

Private Sub mnuOpen_Click()
Dim Answ, cdDir, cdIndex, Pos
imgOpen.Visible = True
imgOpenOver.Visible = False
imgOpenDown.Visible = False
chkSave
'=ClearUndo========
    UpdateUndo
    cmdUndo.Visible = False
    DeleteCollections
'=Initialize Switches=========
setSwitchesFalse
'=====================
ComDia.CancelError = True
On Error GoTo ex
ComDia.FileName = ""
cdDir = GetSetting("vbIconMaker", "ComDiaSettings", "cdOpenDirSetting")
cdIndex = GetSetting("vbIconMaker", "ComDiaSettings", "cdOpenIndexSetting")
If cdDir = "" Or cdIndex = "" Then GoTo NoRegVal 'first time
ComDia.FilterIndex = cdIndex
ComDia.InitDir = cdDir
NoRegVal: ComDia.flags = cdlOFNFileMustExist
ComDia.Filter = "Icons (*.ico;*.cur)|*.ico;*.cur|Images (*.bmp;*.jpg;*gif;*wmf)|*.bmp;*.jpg;*.gif;*wmf"
ComDia.ShowOpen
Pos = InStrRev(ComDia.FileName, "\")
cdDir = Mid(ComDia.FileName, 1, Pos)
cdIndex = ComDia.FilterIndex
SaveSetting "vbIconMaker", "ComDiaSettings", "cdOpenDirSetting", cdDir
SaveSetting "vbIconMaker", "ComDiaSettings", "cdOpenIndexSetting", cdIndex
MousePointer = 11
picReal.Picture = LoadPicture()
picTest.Picture = LoadPicture()
picReal.BackColor = RGB(197, 197, 197)
picTest.Picture = LoadPicture(ComDia.FileName)
DoEvents
If picTest.ScaleWidth > 32 Or picTest.ScaleHeight > 32 Then
    frmScroll.Show '1, Me
    While Not iDone
    DoEvents
    Wend
    iDone = False
Else
    picReal = LoadPicture(ComDia.FileName)
End If
PaintDown
Form1.Refresh
MousePointer = 0
lblPath.Caption = ComDia.FileName
cmdPaint_Click
UpdateUndo
Dirty = False

Exit Sub
ex: If Err.Number = 32755 Then Exit Sub 'user pressed cancel
MsgBox "Error # " & Err.Number & " - " & Err.Description
Exit Sub
End Sub

Private Sub mnuExtract_Click()
Dim Answ, cdDir, cdIndex, Pos
   Dim hImgLarge As Long
   Dim hImgSmall As Long   'the handle to the system image list
   Dim fName As String     'the file name to get icon from
   Dim fnFilter As String  'the file name filter
   Dim r As Long
chkSave
Dirty = False
'ClearUndo
    UpdateUndo
    DeleteCollections
    
'=Initialize Switches=========
setSwitchesFalse
'=====================
   
   On Local Error GoTo cmdLoadErrorHandler
   
  'get the file from the user
   fnFilter$ = "All Files (*.*)|*.*"
'==========
ComDia.FileName = ""
cdDir = GetSetting("vbIconMaker", "ComDiaSettings", "cdExtractDirSetting")
cdIndex = GetSetting("vbIconMaker", "ComDiaSettings", "cdExtractIndexSetting")
If cdDir = "" Or cdIndex = "" Then GoTo NoRegVal 'first time
ComDia.FilterIndex = cdIndex
ComDia.InitDir = cdDir
NoRegVal: ComDia.flags = cdlOFNFileMustExist
'===========

   ComDia.CancelError = True
   ComDia.Filter = fnFilter$
   ComDia.ShowOpen
'============
Pos = InStrRev(ComDia.FileName, "\")
cdDir = Mid(ComDia.FileName, 1, Pos)
cdIndex = ComDia.FilterIndex
SaveSetting "vbIconMaker", "ComDiaSettings", "cdExtractDirSetting", cdDir
SaveSetting "vbIconMaker", "ComDiaSettings", "cdExtractIndexSetting", cdIndex
'============
picReal.Picture = LoadPicture()
picTest.Picture = LoadPicture()
   fName$ = ComDia.FileName
   
'get the system icon associated with that file
   hImgSmall& = SHGetFileInfo(fName$, 0&, _
                              shinfo, Len(shinfo), _
                              BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)

   hImgLarge& = SHGetFileInfo(fName$, 0&, _
                              shinfo, Len(shinfo), _
                              BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
   
   picTest.Picture = LoadPicture()
   picTest.AutoRedraw = True

   picTest.BackColor = RGB(197, 197, 197)
  'draw the associated icon into the pictureboxes
   Call ImageList_Draw(hImgLarge&, shinfo.iIcon, picReal.hdc, 0, 0, ILD_TRANSPARENT)

PaintDown
Form1.Refresh
UpdateUndo
cmdUndo.Visible = False
MousePointer = 0
lblPath.Caption = ComDia.FileName
cmdPaint_Click
Dirty = False
Exit Sub

cmdLoadErrorHandler: If Err.Number = 32755 Then Exit Sub 'user pressed cancel
MsgBox "Error # " & Err.Number & " - " & Err.Description
Exit Sub
End Sub



Private Sub mnuPopCancel_Click()
cmdRegion_Click
End Sub

Private Sub mnuPopCopy_Click()
Clipboard.Clear
Clipboard.SetData picClip.Clip
            '===========Allow another Select Region to be made========
            cmdRegion_Click
UpdateUndo
End Sub

Private Sub mnuPopCut_Click()

Dim c
Clipboard.Clear
Clipboard.SetData picClip.Clip
               For r = yDelLo To ydelHi
                  For c = xDelLo To xdelHi
                       picReal.PSet (c, r), RGB(197, 197, 197)
                   Next c
               Next r
        PaintDown
        UpdateUndo
        Form1.Refresh
        cmdRegion_Click
End Sub

Private Sub mnuPopDelete_Click()
Dim c
Dirty = True
               For r = yDelLo To ydelHi
                  For c = xDelLo To xdelHi
                       picReal.PSet (c, r), RGB(197, 197, 197)
                   Next c
               Next r
        PaintDown
        UpdateUndo
        Form1.Refresh
        cmdRegion_Click
End Sub

Private Sub pasteItNow()
Dim r%, c%
picContainer.Picture = picContainer.Image
Dirty = True
cmdRegion_Click
okToMove = True
pasteIt = True
shRect.Visible = False
picMove.Visible = False
If Clipboard.GetFormat(vbCFBitmap) Then
        picCB.Picture = Clipboard.GetData(vbCFBitmap)
    If picCB.Width > 32 Or picCB.Height > 32 Then
        IsCB = True
        frmScroll.Show '1, Me
        While Not iDone
        DoEvents
        Wend
        iDone = False
    End If
        DoEvents
        picMove.Width = 10 * picCB.Width
        picMove.Height = 10 * picCB.Height
        picCB.Picture = picCB.Image
        '==build mask and sprite
        picCBmask.Width = picCB.Width
        picCBmask.Height = picCB.Height
        picCBsprite.Width = picCB.Width
        picCBsprite.Height = picCB.Height
        For r = 0 To picCB.Height - 1
            For c = 0 To picCB.Width - 1
                If picCB.Point(c, r) <> RGB(197, 197, 197) Then
                    picCBmask.PSet (c, r), vbBlack
                    picCBsprite.PSet (c, r), picCB.Point(c, r)
                Else
                    picCBmask.PSet (c, r), vbWhite
                    picCBsprite.PSet (c, r), vbBlack
                End If
            Next c
        Next r
            picCBmask.Picture = picCBmask.Image
            picCBsprite.Picture = picCBsprite.Image
        '========

End If
        shRect.Move 0, 0, picMove.Width + 4, picMove.Height + 4
        picMove.Move 4, 4
        shRect.Visible = True
        'paint to picContainer image
            StretchBlt picContainer.hdc, ((shRect.Left + 2) \ 10) * 10, ((shRect.Top + 2) \ 10) * 10, picCBmask.Width * 10, picCBmask.Height * 10, picCBmask.hdc, 0, 0, picCBmask.Width, picCBmask.Height, vbSrcAnd
            StretchBlt picContainer.hdc, ((shRect.Left + 2) \ 10) * 10, ((shRect.Top + 2) \ 10) * 10, picCBsprite.Width * 10, picCBsprite.Height * 10, picCBsprite.hdc, 0, 0, picCBsprite.Width, picCBsprite.Height, vbSrcPaint
        DoEvents
        
End Sub

Private Sub mnuRotateLeft_Click()
    Picture2.Cls
    Call MovePixelsLeft
    picReal.Picture = LoadPicture()
    Picture2.Picture = Picture2.Image
    picReal.Picture = Picture2.Picture
    picReal.Refresh
    UpdateUndo
    PaintDown
    Form1.Refresh
    Dirty = True
End Sub
Private Sub MovePixelsLeft()
Dim r, c, p
For c = 0 To 31
    For r = 0 To 31
        p = picReal.Point(c, r)
        Picture2.PSet (r, 31 - c), p
    Next r
Next c
End Sub
Private Sub mnuRotateRight_Click()
    Picture2.Cls
    Call MovePixelsRight
    picReal.Picture = LoadPicture()
    Picture2.Picture = Picture2.Image
    picReal.Picture = Picture2.Picture
    picReal.Refresh
    UpdateUndo
    PaintDown
    Form1.Refresh
    Dirty = True
End Sub
Private Sub MovePixelsRight()
Dim r, c, p
For r = 0 To 31
    For c = 0 To 31
        p = picReal.Point(c, r)
        Picture2.PSet (31 - r, c), p
    Next c
Next r
End Sub


Private Sub mnuSave_Click()
Dim Ret, bmpPicInfo As BITMAPINFO, Answ, cdDir, cdIndex, Pos
Dim sPos, ePos
imgSave.Visible = True
imgSaveOver.Visible = False
imgSaveDown.Visible = False
'==========
cdDir = GetSetting("vbIconMaker", "ComDiaSettings", "cdSaveAsDirSetting")
cdIndex = GetSetting("vbIconMaker", "ComDiaSettings", "cdSaveAsIndexSetting")
If cdDir = "" Or cdIndex = "" Then GoTo NoRegVal 'first time
ComDia.FilterIndex = cdIndex
ComDia.InitDir = cdDir
'===========
NoRegVal: ComDia.CancelError = True
On Error GoTo ExitIt

ComDia.FileName = "Created"
ePos = InStrRev(lblPath.Caption, ".")
sPos = InStrRev(lblPath.Caption, "\")
If ePos > 0 Then
ComDia.FileName = LCase(Mid(lblPath.Caption, sPos + 1, (ePos - sPos) - 1))
End If
ComDia.flags = cdlOFNOverwritePrompt + cdlOFNNoReadOnlyReturn
ComDia.Filter = "Icons (*.ico)|*.ico|Bitmaps (*.bmp)|*.bmp|Cursors (*.cur)|*.cur"
ComDia.ShowSave
'============
Pos = InStrRev(ComDia.FileName, "\")
cdDir = Mid(ComDia.FileName, 1, Pos)
cdIndex = ComDia.FilterIndex
SaveSetting "vbIconMaker", "ComDiaSettings", "cdSaveAsDirSetting", cdDir
SaveSetting "vbIconMaker", "ComDiaSettings", "cdSaveAsIndexSetting", cdIndex
'============

If ComDia.FilterIndex = 1 Then 'Save as Icon
    MousePointer = 11
    Form2.Show vbModal, Me
    If CancelIt = True Then 'User pressed Cancel on SaveAsOptions Form
    CancelIt = False
    MousePointer = 0
    Exit Sub
    End If
    '============
    lblPath.Caption = ComDia.FileName
    If Form2.opt1Bit Then BitCnt = 1
    If Form2.opt4Bit Then BitCnt = 4
    If Form2.opt8Bit Then BitCnt = 8
    If Form2.opt24Bit Then BitCnt = 24
    '=================
        With bmpPicInfo.bmiHeader
        .biBitCount = 24
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bmpPicInfo.bmiHeader)
        .biWidth = 32
        .biHeight = 32
    End With
    IconInfo.iDC = CreateCompatibleDC(0)
    IconInfo.iWidth = 32
    IconInfo.iHeight = 32
    bi24BitInfo.bmiHeader.biWidth = 32
    bi24BitInfo.bmiHeader.biHeight = 32
    IconInfo.iBitmap = CreateDIBSection(IconInfo.iDC, bmpPicInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    SelectObject IconInfo.iDC, IconInfo.iBitmap
    Ret = BitBlt(IconInfo.iDC, 0, 0, 32, 32, picReal.hdc, 0, 0, vbSrcCopy)
    If Ret = 0 Then
    MsgBox "Unable to BitBlt Picture."
    Exit Sub
    End If
    DoEvents
    SaveIcon ComDia.FileName, IconInfo.iDC, IconInfo.iBitmap, BitCnt ', CLng(SaveTypeIn)
    IconInfo.iFileName = ComDia.FileName
    DeleteDC IconInfo.iDC
    DeleteObject IconInfo.iBitmap
    '==================
    picReal.BackColor = RGB(197, 197, 197)
picTest.Picture = LoadPicture(ComDia.FileName)
DoEvents
If picTest.ScaleWidth > 32 Or picTest.ScaleHeight > 32 Then
Answ = MsgBox("Image not in 32X32 pixel format. Do you wish to resize?", vbYesNo, "Image Size Test")
    If Answ = vbYes Then
    picReal.PaintPicture picTest.Image, 0, 0, 32, 32
    Else
        Exit Sub
    End If
Else
picReal = LoadPicture(ComDia.FileName)
End If
PaintDown
Form1.Refresh
'======new method to get rid of black ===========
If BitCnt = 24 Then
    Dim hIcon
PicIcon.Picture = LoadPicture()
PicIcon.Cls
    ExtractIconEx ComDia.FileName, 0, hIcon, 0, 1
    Ret = DrawIconEx(PicIcon.hdc, 0, 0, hIcon, 32, 32, 0, 0, &H3&) 'Const DI_NORMAL = &H3 Both Mask and Image
    If Ret = 0 Then
        MsgBox "Unable to draw PicIcon"
    End If
    PicIcon.Refresh
    Ret = DrawIconEx(picMask.hdc, 0, 0, hIcon, 32, 32, 0, 0, &H1&) 'Const DI_MASK = &H1
    If Ret = 0 Then
        MsgBox "Unable to draw picMask"
    End If
    picMask.Refresh
    Ret = DrawIconEx(PicImage.hdc, 0, 0, hIcon, 32, 32, 0, 0, &H2&) 'Const DI_IMAGE = &H2
    If Ret = 0 Then
        MsgBox "Unable to draw PicImage"
    End If
    PicImage.Refresh
    DestroyIcon hIcon
    
    WriteDataToFile ComDia.FileName

End If
    '==================
    MousePointer = 0
    
End If
'=========================
If ComDia.FilterIndex = 2 Then 'Save as Bmp
    SavePicture picReal.Image, ComDia.FileName
    lblPath.Caption = ComDia.FileName
End If
'========================
If ComDia.FilterIndex = 3 Then 'Save as Cur
Answ = MsgBox("Do you wish to save in Color to be used in creating an animated (.ani) cursor file later?", vbYesNo, "Color or B-W?")
    MousePointer = 11
    If Answ = vbNo Then
        BitCnt = 1
    Else
        BitCnt = 24
    End If
    '=================
        With bmpPicInfo.bmiHeader
        .biBitCount = 24
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bmpPicInfo.bmiHeader)
        .biWidth = 32
        .biHeight = 32
    End With
    IconInfo.iDC = CreateCompatibleDC(0)
    IconInfo.iWidth = 32
    IconInfo.iHeight = 32
    bi24BitInfo.bmiHeader.biWidth = 32
    bi24BitInfo.bmiHeader.biHeight = 32
    IconInfo.iBitmap = CreateDIBSection(IconInfo.iDC, bmpPicInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    SelectObject IconInfo.iDC, IconInfo.iBitmap
    Ret = BitBlt(IconInfo.iDC, 0, 0, 32, 32, picReal.hdc, 0, 0, vbSrcCopy)
    If Ret = 0 Then
    MsgBox "Unable to BitBlt Picture."
    Exit Sub
    End If
    DoEvents
    If Dir(App.Path & "\temp.ico") <> "" Then Kill App.Path & "\temp.ico"
    SaveIcon App.Path & "\temp.ico", IconInfo.iDC, IconInfo.iBitmap, BitCnt ', CLng(SaveTypeIn)
    DeleteDC IconInfo.iDC
    DeleteObject IconInfo.iBitmap
    DoEvents
    '==================

 Form3.Show vbModal, Me
    If CancelIt = True Then 'User pressed Cancel on SetCursorHotspots Form
    CancelIt = False
    MousePointer = 0
    Exit Sub
    End If
    lblPath.Caption = ComDia.FileName
    '==========
    picReal.BackColor = RGB(197, 197, 197)
picTest.Picture = LoadPicture(ComDia.FileName)
DoEvents
If picTest.ScaleWidth > 32 Or picTest.ScaleHeight > 32 Then
Answ = MsgBox("Image not in 32X32 pixel format. Do you wish to resize?", vbYesNo, "Image Size Test")
    If Answ = vbYes Then
    picReal.PaintPicture picTest.Image, 0, 0, 32, 32
    Else
        Exit Sub
    End If
Else
picReal = LoadPicture(ComDia.FileName)
End If
PaintDown
MousePointer = 0
    '==================
    MousePointer = 0
End If

Dirty = False
Form1.Refresh
Exit Sub
ExitIt: If Err.Number = 32755 Then Exit Sub 'user pressed cancel
MsgBox "Error # " & Err.Number & " - " & Err.Description
Exit Sub
End Sub
Private Sub WriteDataToFile(Fn$)

    Dim MaskString$
    Dim Msg$
    Dim F%, h%, w%
    Dim c1&, c2&, r&, g&, b&, k%, n%

    On Error GoTo WriteError

    F = FreeFile
    Open Fn For Binary Access Write As #F


         For k = Len(Fn) To 1 Step -1
             If Mid(Fn, k, 1) = "\" Then Exit For
         Next

         Put #F, 1, ID
         Put #F, 7, IDE
         Put #F, 23, BIH
         k = 63
         For h = 31 To 0 Step -1
             For w = 0 To 31
                 c1 = GetPixel(PicImage.hdc, w, h)
                 c2 = GetPixel(picMask.hdc, w, h)
                 If c2 = &HFFFFFF Then
                    Put #F, k, 0
                    Put #F, k + 1, 0
                    Put #F, k + 2, 0
                 Else
                    b = c1 \ 65536
                    g = (c1 - b * 65536) \ 256
                    r = c1 - b * 65536 - g * 256
                    Put #F, k, b
                    Put #F, k + 1, g
                    Put #F, k + 2, r
                 End If
                 k = k + 3
             Next
         Next
         k = 0
         n = 0
         For h = 31 To 0 Step -1
             For w = 0 To 31
                 If GetPixel(picMask.hdc, w, h) = &HFFFFFF Then
                    MaskString = MaskString & "1"
                 Else
                    MaskString = MaskString & "0"
                 End If
                 k = k + 1
                 If k = 8 Then
                    k = 0
                    Put #F, n + 3135, BinaryStringToByte(MaskString)
                    MaskString = ""
                    n = n + 1
                 End If
             Next
         Next
    Close #F

    CurrentFile = Fn
    On Error GoTo 0
    Exit Sub

WriteError:

    Screen.MousePointer = 0

    If Err.Number <> cdlCancel Then
       Msg = Err.Description & "."
       Msg = Msg & vbCrLf & vbCrLf
       If CurrentFile = "Untitled" Then
          Msg = Msg & "Unable to save Untitled."
       Else
          Msg = Msg & "Unable to save " & CurrentName
       End If
       MsgBox Msg, vbExclamation, Ttl & " - Error"
    End If
    'bFileSaved = False
    Err.Clear
    Exit Sub

End Sub
Private Function BinaryStringToByte(MS$) As Byte

    Dim k%, Rv As Byte

    For k = 1 To 8
        If Mid(MS, k, 1) = "1" Then Rv = Rv + 2 ^ (8 - k)
    Next

    BinaryStringToByte = Rv

End Function
Private Sub mnuExit_Click()
Call Form_QueryUnload(0, 0)
End
End Sub

Private Sub cmdPicker_Click()
If selRegion Then Exit Sub
If chkPix = True Then
MsgBox "Click on pixel in Expanded View"
Exit Sub
End If
pickColor = True
picContainer.MouseIcon = picColorpicker.Picture
End Sub

Private Sub mnuAbout_Click()
Dim Msg As String
Msg = "Simple VB6 Example for Creating or Editing 32X32 Pixel Icons Or Cursors." & vbCrLf & "Neil Crosby (ncrosby@swbell.net)"
Msg = Msg & vbCrLf & vbCrLf & "You can:" & vbCrLf
Msg = Msg & "Create or Edit an Icon or Cursor in various color depths including True Color." & vbCrLf
Msg = Msg & "Set Cursor Hotspots." & vbCrLf
Msg = Msg & "Create and/or View Animated Cursors." & vbCrLf
Msg = Msg & "Drag and Drop files to the Program's shortcut or exe." & vbCrLf
Msg = Msg & "Extract an icon from any file." & vbCrLf
Msg = Msg & "Select Regions for Import from .bmp, .jpg, .gif, or .wmf image files greater than 32X32 pixels." & vbCrLf
Msg = Msg & vbCrLf & "(A Save option is NOT included to prevent accidentally overlaying system files.)"
MsgBox Msg
End Sub

Public Sub PaintDown()
Dim F
Static Pont
Pont = picContainer.Point(0, 0)
picContainer.PaintPicture picReal.Image, 0, 0, 321, 321
If Pont = &HFFC0FF Then
For F = 0 To picContainer.ScaleHeight Step 10
picContainer.Line (0, F)-(picContainer.ScaleWidth, F), &HFFC0FF
Next F
For F = 0 To picContainer.ScaleWidth Step 10
picContainer.Line (F, 0)-(F, picContainer.ScaleHeight), &HFFC0FF
Next F
Line (picReal.Left - 1, picReal.Top - 1)-(picReal.Left + picReal.Width, picReal.Top + picReal.Height), QBColor(15), B
Else
For F = 0 To picContainer.ScaleHeight Step 10
picContainer.Line (0, F)-(picContainer.ScaleWidth, F), &H4040&
Next F
For F = 0 To picContainer.ScaleWidth Step 10
picContainer.Line (F, 0)-(F, picContainer.ScaleHeight), &H4040&
Next F
Line (picReal.Left - 1, picReal.Top - 1)-(picReal.Left + picReal.Width, picReal.Top + picReal.Height), QBColor(0), B
End If
End Sub

Private Sub mnuHorz_Click()
Dim pX As Long, pY As Long, retval As Long
On Error GoTo errMsg
picTemp.Cls
pX = picReal.ScaleWidth
pY = picReal.ScaleHeight
picTemp.Width = picReal.Width
picTemp.Height = picReal.Height
retval = StretchBlt(picTemp.hdc, pX - 1, 0, -pX, pY, _
picReal.hdc, 0, 0, pX, pY, SRCCOPY)
picReal.Cls
picTemp.Picture = picTemp.Image
picReal.PaintPicture picTemp.Picture, 0, 0, _
picTemp.Width, picTemp.Height, 0, 0, _
picTemp.Width, picTemp.Height, vbSrcCopy
picReal.Picture = picReal.Image
UpdateUndo
PaintDown
Form1.Refresh
Exit Sub
errMsg: MsgBox "Error # " & Err.Number & " " & Err.Description
Err.Clear
picTemp.Cls
picTemp.Picture = LoadPicture()
End Sub

Private Sub mnuVert_Click()
Dim pX As Long, pY As Long, retval As Long
On Error GoTo errMsg
picTemp.Cls
pX = picReal.ScaleWidth
pY = picReal.ScaleHeight
picTemp.Width = picReal.Width
picTemp.Height = picReal.Height
retval = StretchBlt(picTemp.hdc, 0, pY - 1, pX, -pY, _
picReal.hdc, 0, 0, pX, pY, SRCCOPY)
picReal.Cls
picTemp.Picture = picTemp.Image
picReal.PaintPicture picTemp.Picture, 0, 0, _
picTemp.Width, picTemp.Height, 0, 0, _
picTemp.Width, picTemp.Height, vbSrcCopy
picReal.Picture = picReal.Image
UpdateUndo
PaintDown
Form1.Refresh
Exit Sub
errMsg: MsgBox "Error # " & Err.Number & " " & Err.Description
Err.Clear
picTemp.Cls
picTemp.Picture = LoadPicture()

End Sub

Private Sub nmuAni_Click()
Form1.Hide
ViewAni.Show
End Sub

Private Sub pic16Color_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If chkPix = True Then
MsgBox "Click on pixel in Expanded View"
Exit Sub
End If
If Button = 1 Then
PicMseColor.BackColor = pic16Color.Point(X, Y)
Else
PicMseColorR.BackColor = pic16Color.Point(X, Y)
End If
End Sub

Private Sub pic16Color_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
clr = pic16Color.Point(X, Y)
r = clr Mod 256
g = (clr \ 256) Mod 256
b = clr \ 256 \ 256
lblRGB.Caption = "R " & r & " G " & g & " B " & b & "   -   " & clr
End Sub

Private Sub picBasic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If chkPix = True Then
MsgBox "Click on pixel in Expanded View"
Exit Sub
End If
If Button = 1 Then
PicMseColor.BackColor = picBasic.Point(X, Y)
Else
PicMseColorR.BackColor = picBasic.Point(X, Y)
End If
End Sub

Private Sub picBasic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
clr = picBasic.Point(X, Y)
r = clr Mod 256
g = (clr \ 256) Mod 256
b = clr \ 256 \ 256
lblRGB.Caption = "R " & r & " G " & g & " B " & b & "   -   " & clr

End Sub

Private Sub picContainer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cr
'========Move Paste Image===
'In picContainer Mouse down, move and up -- 'Move Paste Image'
'must be before 'Select Region'
If okToMove Then 'set True in PasteItNow
    moveIt = True
    pX = X
    pY = Y
    prevX = X \ 10
    prevY = Y \ 10
    pXOff = X - shRect.Left
    pYOff = Y - shRect.Top
    Exit Sub
End If
'=========Select Region==========
If selRegion = True Then
    If selectIt Then
        canSelect = True
        xStart = (X \ 10) * 10
        yStart = (Y \ 10) * 10
        XLo = (X \ 10) * 10
        YLo = (Y \ 10) * 10
        XHi = (X \ 10) * 10
        YHi = (Y \ 10) * 10
        shRect.Width = Abs(XHi - XLo)
        shRect.Height = Abs(YHi - YLo)
        Exit Sub
    End If
End If
'=========Pick a Color======================
If pickColor = True Then
If Button = 1 Then
PicMseColor.BackColor = picContainer.Point(X, Y)
Else
PicMseColorR.BackColor = picContainer.Point(X, Y)
End If
picContainer.MouseIcon = picPencil.Picture
    eraseIt = False
    pickColor = False
    Exit Sub
End If
'========Draw Text=================
If textDraw Then
Dirty = True
        If Button = 1 Then
            picReal.ForeColor = PicMseColor.BackColor
            Form4.Text1.ForeColor = PicMseColor.BackColor
        Else
            picReal.ForeColor = PicMseColorR.BackColor
            Form4.Text1.ForeColor = PicMseColorR.BackColor
        End If
curX = X \ 10
curY = Y \ 10
DoEvents
Form4.Show 1, Me
        picReal.Picture = picReal.Image
        UpdateUndo
        PaintDown
        Form1.Refresh
End If
'=========Flood an Area===========
If floodDraw Then
Dirty = True
picReal.FillStyle = 0 'Solid
        If Button = 1 Then
            picReal.FillColor = PicMseColor.BackColor
        Else
            picReal.FillColor = PicMseColorR.BackColor
        End If
ExtFloodFill picReal.hdc, X \ 10, Y \ 10, picReal.Point(X \ 10, Y \ 10), FLOODFILLSURFACE
picReal.Picture = picReal.Image
UpdateUndo
PaintDown
Form1.Refresh
End If
'=========Draw a Line=======================
If lineDraw Then
    lineOKDraw = True
    lineX1 = X
    lineY1 = Y
    Exit Sub
End If
'=========Draw a Rectangle Or FillBox=======================
If rectDraw Then
    rectOKDraw = True
    lineX1 = X
    lineY1 = Y
    Exit Sub
End If
If fillBoxDraw Then
    rectOKDraw = True
    lineX1 = X
    lineY1 = Y
    Exit Sub
End If
'=========Draw a Circle or Filled Circle=======================
If circleDraw Then
    circleOKDraw = True
    lineX1 = X
    lineY1 = Y
    Exit Sub
End If
If fillCircleDraw Then
    circleOKDraw = True
    lineX1 = X
    lineY1 = Y
    Exit Sub
End If

'=========Change a Pixel Color to Transparent=======
If chkPix = True Then
chgColor = picContainer.Point(X, Y)
chkPix = False
Exit Sub
End If
'=========Original Draw - One Pixel at a time=======
If pixelDraw = True Then
canDraw = True
End If
End Sub

Private Sub picContainer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cr, xM, yM
'==========Show Mouse Position=======
xM = Int(X / 10) + 1
yM = Int(Y / 10) + 1
lblMsePos.Caption = "Mouse (X,Y) is " & xM & "," & yM
'=========Show color of pixel mouse is over========
clr = picContainer.Point(X, Y)
r = clr Mod 256
g = (clr \ 256) Mod 256
b = clr \ 256 \ 256
lblRGB.Caption = "R " & r & " G " & g & " B " & b & "   -   " & clr
'===========Show Hand on Paste image========
        If okToMove = True And X > shRect.Left And X < shRect.Left + shRect.Width And Y > shRect.Top And Y < shRect.Top + shRect.Height Then
            picContainer.MouseIcon = picHand.Picture
        Else
            picContainer.MouseIcon = picPencil.Picture
        End If
'=========Move Paste image=============
    If moveIt Then
        If X \ 10 = prevX And Y \ 10 = prevY Then
            prevX = X \ 10
            prevY = Y \ 10
            Exit Sub
        Else
            Timer1.Enabled = False
            shRect.Visible = False
            picContainer.Cls
            shRect.Left = X - pXOff
            shRect.Top = Y - pYOff
            StretchBlt picContainer.hdc, ((shRect.Left + 2) \ 10) * 10, ((shRect.Top + 2) \ 10) * 10, picCBmask.Width * 10, picCBmask.Height * 10, picCBmask.hdc, 0, 0, picCBmask.Width, picCBmask.Height, vbSrcAnd
            StretchBlt picContainer.hdc, ((shRect.Left + 2) \ 10) * 10, ((shRect.Top + 2) \ 10) * 10, picCBsprite.Width * 10, picCBsprite.Height * 10, picCBsprite.hdc, 0, 0, picCBsprite.Width, picCBsprite.Height, vbSrcPaint
            prevX = X \ 10
            prevY = Y \ 10
          Exit Sub
        End If
    End If


'========Select Region==========
If selRegion Then

    If X < xStart And Y < yStart Then
            XLo = xStart + 10
            YLo = yStart + 10
            XHi = ((X \ 10) * 10)
            YHi = ((Y \ 10) * 10)
            
    End If
    If X > xStart And Y > yStart Then
            XLo = xStart
            YLo = yStart
            XHi = ((X \ 10) * 10) + 10
            YHi = ((Y \ 10) * 10) + 10
    End If
    If X > xStart And Y < yStart Then
            XLo = xStart
            YLo = yStart + 10
            XHi = ((X \ 10) * 10) + 10
            YHi = ((Y \ 10) * 10)
    End If
    If X < xStart And Y > yStart Then
            XLo = xStart + 10
            YLo = yStart
            XHi = ((X \ 10) * 10)
            YHi = ((Y \ 10) * 10) + 10
    End If
    If XHi < 0 Then XHi = 0
    If YHi < 0 Then YHi = 0
    If XHi > picContainer.ScaleWidth - 1 Then XHi = picContainer.ScaleWidth - 1
    If YHi > picContainer.ScaleHeight - 1 Then YHi = picContainer.ScaleHeight - 1
    If XLo < 0 Then XLo = 0
    If YLo < 0 Then YLo = 0
    If XLo > picContainer.ScaleWidth - 1 Then XLo = picContainer.ScaleWidth - 1
    If YLo > picContainer.ScaleHeight - 1 Then YLo = picContainer.ScaleHeight - 1
    If canSelect = True Then
            shRect.Width = Abs(XHi - XLo)
            shRect.Height = Abs(YHi - YLo)
            shRect.Visible = True
        If XHi > XLo And YHi > YLo Then
            shRect.Top = YLo
            shRect.Left = XLo
        End If
        If XHi > XLo And YHi < YLo Then
            shRect.Top = YHi
            shRect.Left = XLo
        End If
        If XHi < XLo And YHi < YLo Then
            shRect.Top = YHi
            shRect.Left = XHi
        End If
        If XHi < XLo And YHi > YLo Then
            shRect.Top = YLo
            shRect.Left = XHi
        End If
    End If
Exit Sub
End If
'=======Pick A Color and Flood=================
If pickColor Then Exit Sub
If floodDraw Then Exit Sub
'=========Draw a Line=======================
    If lineOKDraw Then
        Line2.x1 = lineX1
        Line2.y1 = lineY1
        Line2.X2 = X
        Line2.Y2 = Y
        If Button = 1 Then
            cr = PicMseColor.BackColor
        Else
            cr = PicMseColorR.BackColor
        End If
        Line2.BorderColor = cr
        Line2.Visible = True

        Exit Sub
    End If
'=========Draw a Rectangle or Fill Box=======================
    If rectDraw Or fillBoxDraw Then
        If rectOKDraw = True Then
                If X > lineX1 Then
                    Shape1.Left = lineX1
                Else
                    Shape1.Left = X
                End If
                
                If Y > lineY1 Then
                    Shape1.Top = lineY1
                Else
                    Shape1.Top = Y
                End If
                Shape1.Width = Abs(X - lineX1)
                Shape1.Height = Abs(Y - lineY1)
                
                If Button = 1 Then
                    cr = PicMseColor.BackColor
                Else
                    cr = PicMseColorR.BackColor
                End If
            Shape1.BorderColor = cr
            Shape1.FillColor = cr
            Shape1.Visible = True
            Exit Sub
        End If
    End If
'=========Draw a Circle or Filled Circle=======================
    If circleDraw Or fillCircleDraw Then
        If circleOKDraw = True Then
                Shape1.Width = Abs(X - lineX1)
                Shape1.Height = Abs(Y - lineY1)
        Shape1.Visible = True
            If X > lineX1 Then
                    Shape1.Left = lineX1 - (Shape1.Width / 2)
                Else
                    Shape1.Left = X + (Shape1.Width / 2)
                End If
                
                If Y > lineY1 Then
                    Shape1.Top = lineY1 - (Shape1.Height / 2)
                Else
                    Shape1.Top = Y + (Shape1.Height / 2)
            End If

                
                If Button = 1 Then
                    cr = PicMseColor.BackColor
                Else
                    cr = PicMseColorR.BackColor
                End If
            Shape1.BorderColor = cr
            Shape1.FillColor = cr
            Shape1.Visible = True
            Exit Sub
        End If
    End If
'=========Change a Pixel Color to Transparent=======

If chkPix = True Then Exit Sub

'=========Original Draw - One Pixel at a time=======
If pixelDraw = True Then
    If canDraw = True Then
                Dirty = True
            If Button = 1 Then
                ColorChg = PicMseColor.BackColor
            Else
                ColorChg = PicMseColorR.BackColor
            End If
                If eraseIt Then ColorChg = RGB(197, 197, 197) 'Transparent picContainer.BackColor
            x1 = 0
            y1 = 0
        For j = 0 To 31
        For p = 0 To 31

                If X < x1 + 10 And X > x1 And Y < y1 + 10 And Y > y1 Then
                    picContainer.Line (x1 + 1, y1 + 1)-(x1 + 9, y1 + 9), ColorChg, BF
                    picReal.PSet (x1 \ 10, y1 \ 10), ColorChg
                End If

            x1 = x1 + 10
                If x1 = 320 Then
                    x1 = 0
                    y1 = y1 + 10
                End If


        Next p
        Next j
    End If
End If

End Sub

Private Sub picContainer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cr, F, xd, yd, r, c
On Error GoTo ExitIt
If pickColor Then Exit Sub
If floodDraw Then Exit Sub
'=======Move Paste Image===========
If moveIt = True Then
'=====paint non transparent to picReal from picCB
xd = (shRect.Left + 2) \ 10
yd = (shRect.Top + 2) \ 10
BitBlt picReal.hdc, xd, yd, picCBmask.Width, picCBmask.Height, picCBmask.hdc, 0, 0, vbSrcAnd
BitBlt picReal.hdc, xd, yd, picCBsprite.Width, picCBsprite.Height, picCBsprite.hdc, 0, 0, vbSrcPaint
picCB.Picture = LoadPicture()
picCBmask.Picture = LoadPicture()
picCBsprite.Picture = LoadPicture()
            picReal.Picture = picReal.Image
            UpdateUndo
            PaintDown
            Form1.Refresh
    cmdRegion_Click
    moveIt = False
    Exit Sub
End If
'=======Select Region==============
If selRegion Then

    If canSelect Then
            'Clip Screen Image
            picReal.Picture = picReal.Image
            picClip.Picture = picReal.Picture
            picContainer.Picture = picContainer.Image
            PicClipMove.Picture = picContainer.Picture
            DoEvents
            ' Get X and Y coordinates of the clipping region.
            picClip.ClipX = (shRect.Left \ 10)
            PicClipMove.ClipX = shRect.Left
            picClip.ClipY = (shRect.Top \ 10)
            PicClipMove.ClipY = shRect.Top
            ' Set the area of the clipping region (in pixels).
        If XHi > 310 Then XHi = 320
        If YHi > 310 Then YHi = 320
            picClip.ClipWidth = (Abs(XHi \ 10 - XLo \ 10))  'shRect.Width
            PicClipMove.ClipWidth = shRect.Width
            picClip.ClipHeight = (Abs(YHi \ 10 - YLo \ 10)) 'shRect.Height
            PicClipMove.ClipHeight = shRect.Height
            If shRect.Width < 2 Or shRect.Height < 2 Then
               MsgBox "Select Area using Click and Drag. After Area is selected, right click in selection for clipboard functions or use Edit Menu."
                cmdRegion_Click
                canSelect = False
                Exit Sub
            End If
            picMove.Width = shRect.Width
            picMove.Height = shRect.Height
            picMove.PaintPicture PicClipMove.Clip, 0, 0
            picMove.Left = shRect.Left
            picMove.Top = shRect.Top
            picMove.Visible = True
            xDelLo = shRect.Left \ 10
            xdelHi = (shRect.Left + shRect.Width) \ 10 - 1
            yDelLo = shRect.Top \ 10
            ydelHi = (shRect.Top + shRect.Height) \ 10 - 1
            '=======
            xOff = shRect.Left
            yOff = shRect.Top
            selectIt = False
    End If
            canSelect = False
            shRect.Move shRect.Left - 1, shRect.Top - 1, shRect.Width + 2, shRect.Height + 2
            Exit Sub
End If
'=========Draw a Line=======================
If lineDraw Then
    Dirty = True
    picContainer.DrawWidth = 10
        If Button = 1 Then
            cr = PicMseColor.BackColor
        Else
            cr = PicMseColorR.BackColor
        End If
    If lineOKDraw Then
        'picContainer.Line (lineX1, lineY1)-(X, Y), cr
        picReal.Line (lineX1 \ 10, lineY1 \ 10)-(X \ 10, Y \ 10), cr
        'color last pixel
        picReal.PSet (X \ 10, Y \ 10), cr
        
        UpdateUndo
        
    End If
        picContainer.DrawWidth = 1
        DoEvents
        picReal.Picture = picReal.Image
        PaintDown
        Form1.Refresh
        lineOKDraw = False
        Line2.Visible = False
        Exit Sub
End If
'=========Draw a Rectangle or Fill Box=======================
If rectDraw Or fillBoxDraw Then
                Dirty = True
        If Button = 1 Then
            cr = PicMseColor.BackColor
        Else
            cr = PicMseColorR.BackColor
        End If
    If rectDraw Then
        picContainer.DrawWidth = 10
        Shape1.FillStyle = 1
        DoEvents
           ' picContainer.Line (lineX1, lineY1)-(X, Y), cr, B
            picReal.Line (lineX1 \ 10, lineY1 \ 10)-(X \ 10, Y \ 10), cr, B
         picContainer.DrawWidth = 1
         Shape1.Visible = False
         rectOKDraw = False
    End If
    If fillBoxDraw Then
           ' picContainer.Line (lineX1, lineY1)-(X, Y), cr, BF
            picReal.Line (lineX1 \ 10, lineY1 \ 10)-(X \ 10, Y \ 10), cr, BF
            Shape1.Visible = False
            rectOKDraw = False
    End If
       
       UpdateUndo
       
        DoEvents
        picReal.Picture = picReal.Image
        PaintDown
        Form1.Refresh
        Exit Sub
End If
'=========Draw a Circle or Filled Circle=======================
If circleDraw Or fillCircleDraw Then
                    Dirty = True
    If Button = 1 Then
            cr = PicMseColor.BackColor
    Else
            cr = PicMseColorR.BackColor
    End If
        picReal.FillColor = cr
        picContainer.FillColor = cr
    If circleDraw Then
        picContainer.DrawWidth = 10
        Shape1.FillStyle = 1
        DoEvents
            picReal.Circle (lineX1 \ 10, lineY1 \ 10), (Shape1.Width \ 10) / 2, cr, , , (Shape1.Height \ 10) / (Shape1.Width \ 10)
            picContainer.DrawWidth = 1
            Shape1.Visible = False
            circleOKDraw = False
    End If
    If fillCircleDraw Then
            picReal.Circle (lineX1 \ 10, lineY1 \ 10), (Shape1.Width \ 10) / 2, cr, , , (Shape1.Height \ 10) / (Shape1.Width \ 10)
            Shape1.Visible = False
            circleOKDraw = False
    End If
           UpdateUndo
           DoEvents
           picReal.Picture = picReal.Image
           PaintDown
           Form1.Refresh
           Exit Sub
End If
'=========Original Draw - One Pixel at a time=======
If pixelDraw = True Then
    canDraw = False
    UpdateUndo
End If
Exit Sub
ExitIt: If circleOKDraw = True Then
Shape1.Visible = False
circleOKDraw = False
End If
End Sub

Private Sub chkSave()
Dim reply
If Dirty = True Then
    reply = MsgBox("Save Current Edited Icon?", vbYesNo, "Save As")
        If reply = vbYes Then mnuSave_Click
End If
End Sub

Private Sub picTransparent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
PicMseColor.BackColor = picTransparent.Point(X, Y)
Else
PicMseColorR.BackColor = picTransparent.Point(X, Y)
End If
End Sub

Private Sub picTransparent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
clr = picTransparent.Point(X, Y)
r = clr Mod 256
g = (clr \ 256) Mod 256
b = clr \ 256 \ 256
lblRGB.Caption = "R " & r & " G " & g & " B " & b & "   -   " & clr
End Sub

Private Sub picMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnuPopUp
Else
    MsgBox "Right click for PopUp menu or use Edit on Menubar."
End If
End Sub

Private Sub PicMseColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
clr = PicMseColor.Point(X, Y)
r = clr Mod 256
g = (clr \ 256) Mod 256
b = clr \ 256 \ 256
lblRGB.Caption = "R " & r & " G " & g & " B " & b & "   -   " & clr
End Sub

Private Sub PicMseColorR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
clr = PicMseColorR.Point(X, Y)
r = clr Mod 256
g = (clr \ 256) Mod 256
b = clr \ 256 \ 256
lblRGB.Caption = "R " & r & " G " & g & " B " & b & "   -   " & clr
End Sub

Private Sub ExtractRequest()
Dim Answ, cdDir, cdIndex, Pos
   Dim hImgLarge As Long
   Dim hImgSmall As Long   'the handle to the system image list
   Dim fName As String     'the file name to get icon from
   Dim r As Long
fName$ = filePath 'Command$
   
   On Local Error GoTo cmdLoadErrorHandler
   
   hImgSmall& = SHGetFileInfo(fName$, 0&, _
                              shinfo, Len(shinfo), _
                              BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)

   hImgLarge& = SHGetFileInfo(fName$, 0&, _
                              shinfo, Len(shinfo), _
                              BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
   
   picTest.Picture = LoadPicture()
   picTest.AutoRedraw = True

   picTest.BackColor = RGB(197, 197, 197)
  'draw the associated icon into the pictureboxes
   Call ImageList_Draw(hImgLarge&, shinfo.iIcon, picReal.hdc, 0, 0, ILD_TRANSPARENT)
PaintDown
Form1.Refresh
MousePointer = 0
UpdateUndo
cmdUndo.Visible = False
lblPath.Caption = filePath 'Command$
Exit Sub

cmdLoadErrorHandler:
MsgBox "Error # " & Err.Number & " - " & Err.Description
Exit Sub
End Sub

Private Sub setSwitchesFalse()
okToMove = False
eraseIt = False
rectDraw = False
lineDraw = False
floodDraw = False
fillBoxDraw = False
circleDraw = False
fillCircleDraw = False
pixelDraw = False
textDraw = False
selRegion = False
Timer1.Enabled = False
shRect.Visible = False
canSelect = False
moveIt = False
selectIt = False
picMove.Visible = False
picMove.Picture = LoadPicture()
picReal16.Picture = LoadPicture()
picContainer.Cls
PaintDown
picContainer.Picture = picContainer.Image
Form1.Refresh
End Sub

Private Sub picReal_DblClick()
picReal16.Picture = LoadPicture()
picReal.Picture = picReal.Image
picReal16.PaintPicture picReal.Picture, 0, 0, 16, 16

End Sub

Private Sub Timer1_Timer()
If shRect.BorderStyle = vbBSDot Then
    shRect.BorderStyle = vbBSDashDot
            Else
                shRect.BorderStyle = vbBSDot
            End If
End Sub
' Return the long file name for a short file name.
Public Function LongFileName(ByVal short_name As String) As String
Dim Pos As Integer
Dim result As String
Dim long_name As String

    ' Start after the drive letter if any.
    If Mid$(short_name, 2, 1) = ":" Then
        result = Left$(short_name, 2)
        Pos = 3
    Else
        result = ""
        Pos = 1
    End If

    ' Consider each section in the file name.
    Do While Pos > 0
        ' Find the next \.
        Pos = InStr(Pos + 1, short_name, "\")

        ' Get the next piece of the path.
        If Pos = 0 Then
            long_name = Dir$(short_name, vbNormal + vbHidden + vbSystem + vbDirectory)
        Else
            long_name = Dir$(Left$(short_name, Pos - 1), vbNormal + vbHidden + vbSystem + vbDirectory)
        End If
        result = result & "\" & long_name
    Loop

    LongFileName = result
End Function

