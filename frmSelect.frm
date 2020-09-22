VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   70
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   448
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Import"
      Height          =   375
      Left            =   5670
      TabIndex        =   6
      Top             =   315
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "Check Image to Import"
      Height          =   780
      Left            =   2070
      TabIndex        =   1
      Top             =   45
      Width           =   3525
      Begin VB.PictureBox picFill 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   2565
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   5
         Top             =   225
         Width           =   510
      End
      Begin VB.PictureBox picRatio 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   1035
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   4
         Top             =   225
         Width           =   510
      End
      Begin VB.OptionButton OptFill 
         Caption         =   "Fill"
         Height          =   195
         Left            =   1980
         TabIndex        =   3
         Top             =   405
         Width           =   510
      End
      Begin VB.OptionButton OptRatio 
         Caption         =   "Aspect Ratio"
         Height          =   420
         Left            =   135
         TabIndex        =   2
         Top             =   270
         Value           =   -1  'True
         Width           =   825
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Color RGB(197,197,197) will be imported as transparent."
      Height          =   195
      Left            =   2070
      TabIndex        =   7
      Top             =   840
      Width           =   4230
   End
   Begin VB.Label Label1 
      Caption         =   "Image is larger than 32X32 pixels. Click and Drag to Select area to import."
      Height          =   870
      Left            =   270
      TabIndex        =   0
      Top             =   90
      Width           =   1410
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim w, h, r, c, pw, ph, pc
If AreaSel = False Then
MsgBox "Select an Area to Import."
Exit Sub
End If
picRatio.Picture = picRatio.Image
picFill.Picture = picFill.Image

Me.MousePointer = 11
If IsCB = True Then
    If OptRatio Then
        For r = 0 To 31
        If w > 0 Then Exit For
            For c = 0 To 31
                If picRatio.Point(r, c) <> RGB(197, 197, 197) Then
                w = w + 1
                
                End If
            Next c
        Next r
        For c = 0 To 31
        If h > 0 Then Exit For
            For r = 0 To 31
                If picRatio.Point(r, c) <> RGB(197, 197, 197) Then
                h = h + 1
                End If
            Next r
        Next c
         Form1.picCB.Picture = LoadPicture()
         Form1.picCB.Width = w
         Form1.picCB.Height = h
         Form1.picMove.Width = w * 10
         Form1.picMove.Height = h * 10
         '=========
        For r = 0 To 31
            For c = 0 To 31
                If picRatio.Point(r, c) <> RGB(197, 197, 197) Then
                pc = picRatio.Point(r, c)
                Form1.picCB.PSet (pw, ph), pc
                pw = pw + 1
                End If
            Next c
            ph = ph + 1
            pw = 0
        Next r
    Else
         Form1.picCB.Width = 32
         Form1.picCB.Height = 32
         Form1.picMove.Height = 32
         Form1.picMove.Width = 32
         Form1.picCB.PaintPicture picFill.Picture, 0, 0, 32, 32
    End If

Else
    If OptRatio Then
         Form1.picReal.PaintPicture picRatio.Picture, 0, 0, 32, 32
    Else
         Form1.picReal.PaintPicture picFill.Picture, 0, 0, 32, 32
    End If
End If
Me.MousePointer = 0
iDone = True
IsCB = False
Unload MDIForm1
End Sub

Private Sub Form_Load()
Me.Move 0, frmScroll.Height
Me.ScaleWidth = frmScroll.ScaleWidth
Me.Width = frmScroll.Width
picRatio.BackColor = RGB(197, 197, 197)
picFill.BackColor = RGB(197, 197, 197)
End Sub
