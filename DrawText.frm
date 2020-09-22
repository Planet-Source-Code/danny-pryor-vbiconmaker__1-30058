VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   Caption         =   "Draw Text"
   ClientHeight    =   1620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3930
   Icon            =   "DrawText.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   1620
   ScaleWidth      =   3930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   3150
      TabIndex        =   3
      Top             =   1125
      Width           =   690
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   285
      Left            =   2295
      TabIndex        =   2
      Top             =   1125
      Width           =   690
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "Change Font"
      Height          =   285
      Left            =   2385
      TabIndex        =   1
      Top             =   90
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   45
      TabIndex        =   0
      Text            =   "Text"
      Top             =   45
      Width           =   1995
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2610
      Top             =   495
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Form4.Hide
End Sub

Private Sub cmdFont_Click()
  ' Set Cancel to True
  CommonDialog1.CancelError = True
  On Error GoTo ErrHandler
  ' Set the Flags property
  CommonDialog1.flags = cdlCFEffects Or cdlCFBoth
  ' Display the Font dialog box
  CommonDialog1.ShowFont
  Text1.Font.Name = CommonDialog1.FontName
  Text1.Font.Size = CommonDialog1.FontSize
  Text1.Font.Bold = CommonDialog1.FontBold
  Text1.Font.Italic = CommonDialog1.FontItalic
  Text1.Font.Underline = CommonDialog1.FontUnderline
  Text1.FontStrikethru = CommonDialog1.FontStrikethru
  Text1.ForeColor = Form1.picReal.ForeColor
  Exit Sub
ErrHandler: If Err.Number = 32755 Then Exit Sub 'user pressed cancel
MsgBox "Error # " & Err.Number & " - " & Err.Description
Exit Sub
End Sub

Private Sub cmdOK_Click()
Form1.picReal.CurrentX = curX - 1
Form1.picReal.CurrentY = curY - 2
Form1.picReal.Font.Size = Text1.Font.Size
Form1.picReal.Font.Name = Text1.Font.Name
Form1.picReal.Font.Bold = Text1.Font.Bold
Form1.picReal.Font.Italic = Text1.Font.Italic
Form1.picReal.Font.Underline = Text1.Font.Underline
Form1.picReal.Font.Strikethrough = Text1.FontStrikethru
Form1.picReal.ForeColor = Text1.ForeColor
Form1.picReal.Print Text1.Text
Form4.Hide
End Sub


