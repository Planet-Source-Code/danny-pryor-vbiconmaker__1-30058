VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form frmScroll 
   BorderStyle     =   0  'None
   Caption         =   "Click and Drag to Select area to import."
   ClientHeight    =   4665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   MDIChild        =   -1  'True
   ScaleHeight     =   311
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   453
   ShowInTaskbar   =   0   'False
   Begin VB.VScrollBar vbarScroller 
      Height          =   4245
      Left            =   6480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   45
      Width           =   240
   End
   Begin VB.HScrollBar hbarScroller 
      Height          =   240
      Left            =   180
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4320
      Width           =   6225
   End
   Begin VB.PictureBox picOuter 
      Height          =   4185
      Left            =   0
      ScaleHeight     =   275
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   426
      TabIndex        =   0
      Top             =   90
      Width           =   6450
      Begin PicClip.PictureClip PicClip 
         Left            =   195
         Top             =   780
         _ExtentX        =   1111
         _ExtentY        =   661
         _Version        =   393216
      End
      Begin VB.PictureBox picInner 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   8985
         Left            =   -45
         ScaleHeight     =   599
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   798
         TabIndex        =   1
         Top             =   90
         Width           =   11970
         Begin VB.Shape shRect 
            Height          =   465
            Left            =   45
            Top             =   45
            Visible         =   0   'False
            Width           =   645
         End
      End
   End
End
Attribute VB_Name = "frmScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XHi, XLo, YHi, YLo, StopDraw


Private Sub Form_Load()
frmSelect.Show
StopDraw = 1
If IsCB = True Then 'Clipboard Picture
picInner.Picture = Form1.picCB.Picture
Else
picInner.Picture = LoadPicture(Form1.ComDia.FileName)
End If

If picInner.ScaleWidth > picOuter.ScaleWidth Or picInner.ScaleHeight > picOuter.ScaleHeight Then
                picInner.Move 0, 0
Else
        If picInner.ScaleWidth < picOuter.ScaleWidth And picInner.ScaleHeight < picOuter.ScaleHeight Then
            picInner.Move (picOuter.ScaleWidth - picInner.Width) \ 2, (picOuter.ScaleHeight - picInner.Height) \ 2
        Else
            picInner.Move 0, 0
        End If
End If
DoEvents

picInner.Picture = picInner.Image
picClip.Picture = picInner.Picture
End Sub

' Position the controls.
Private Sub Form_Resize()
Dim got_wid As Single
Dim got_hgt As Single
Dim need_wid As Single
Dim need_hgt As Single
Dim need_hbar As Boolean
Dim need_vbar As Boolean

    'If WindowState = vbMinimized Then Exit Sub

    need_wid = picInner.Width + (picOuter.Width - picOuter.ScaleWidth)
    need_hgt = picInner.Height + (picOuter.Height - picOuter.ScaleHeight)
    got_wid = ScaleWidth
    got_hgt = ScaleHeight

    ' See which scroll bars we need.
    need_hbar = (need_wid > got_wid)
    If need_hbar Then got_hgt = got_hgt - hbarScroller.Height

    need_vbar = (need_hgt > got_hgt)
    If need_vbar Then
        got_wid = got_wid - vbarScroller.Width
        If Not need_hbar Then
            need_hbar = (need_wid > got_wid)
            If need_hbar Then got_hgt = got_hgt - hbarScroller.Height
        End If
    End If

    picOuter.Move 0, 0, got_wid, got_hgt

    If need_hbar Then
        hbarScroller.Move 0, got_hgt, got_wid
        hbarScroller.Visible = True
    Else
        hbarScroller.Visible = False
    End If

    If need_vbar Then
        vbarScroller.Move got_wid, 0, vbarScroller.Width, got_hgt
        vbarScroller.Visible = True
    Else
        vbarScroller.Visible = False
    End If

    ' Set the scrollbar properties.
    SetScrollBars
End Sub

Private Sub hbarScroller_Change()
    picInner.Left = hbarScroller.Value
End Sub


Private Sub hbarScroller_Scroll()
    picInner.Left = hbarScroller.Value
End Sub

Private Sub picInner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
StopDraw = 0
XLo = X
YLo = Y
XHi = X
YHi = Y
shRect.Width = Abs(XHi - XLo)
shRect.Height = Abs(YHi - YLo)
End Sub

Private Sub picInner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
XHi = X
YHi = Y
If XHi < 0 Then XHi = 0
If YHi < 0 Then YHi = 0
If XHi > picInner.ScaleWidth Then XHi = picInner.ScaleWidth
If YHi > picInner.ScaleHeight Then YHi = picInner.ScaleHeight
If StopDraw = 0 Then
shRect.Width = Abs(XHi - XLo)
shRect.Height = Abs(YHi - YLo)
'frmContainer.Caption = "Width " & shRect.Width & " Height " & shRect.Height
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
DoEvents
End Sub

Private Sub picInner_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
StopDraw = 1
'Clip Screen Image
    ' Get X and Y coordinates of the clipping region.
        picClip.ClipX = shRect.Left
        picClip.ClipY = shRect.Top
    ' Set the area of the clipping region (in pixels).
    picClip.ClipWidth = shRect.Width
    picClip.ClipHeight = shRect.Height
    If shRect.Width < 2 Or shRect.Height < 2 Then
        Exit Sub
    End If
    ' Assign the clipped bitmap to the form picture.
    frmSelect.picRatio.Picture = LoadPicture()
If Not IsCB Then
    If shRect.Height > shRect.Width Then
        frmSelect.picRatio.PaintPicture picClip.Clip, (32 - Int((shRect.Width / shRect.Height) * 32)) \ 2, 0, Int((shRect.Width / shRect.Height) * 32), 32
    Else
    frmSelect.picRatio.PaintPicture picClip.Clip, 0, (32 - Int((shRect.Height / shRect.Width) * 32)) \ 2, 32, Int((shRect.Height / shRect.Width) * 32)
    End If
End If
If IsCB Then
    frmSelect.picRatio.PaintPicture picClip.Clip, 0, 0
End If
    frmSelect.picFill.PaintPicture picClip.Clip, 0, 0, 32, 32
    AreaSel = True
End Sub

Private Sub vbarScroller_Change()
    picInner.Top = vbarScroller.Value
End Sub

Private Sub vbarScroller_Scroll()
    picInner.Top = vbarScroller.Value
End Sub

' Set scroll bar properties.
Private Sub SetScrollBars()
    vbarScroller.Min = 0
    vbarScroller.Max = picOuter.ScaleHeight - picInner.Height
    vbarScroller.LargeChange = picOuter.ScaleHeight
    vbarScroller.SmallChange = picOuter.ScaleHeight / 5
    
    hbarScroller.Min = 0
    hbarScroller.Max = picOuter.ScaleWidth - picInner.Width
    hbarScroller.LargeChange = picOuter.ScaleWidth
    hbarScroller.SmallChange = picOuter.ScaleWidth / 5
End Sub


