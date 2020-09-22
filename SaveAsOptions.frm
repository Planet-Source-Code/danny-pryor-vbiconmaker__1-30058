VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Save As Options"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2670
   Icon            =   "SaveAsOptions.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2610
   ScaleWidth      =   2670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1350
      TabIndex        =   5
      Top             =   2160
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   180
      TabIndex        =   4
      Top             =   2160
      Width           =   780
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1815
      Left            =   315
      TabIndex        =   0
      Top             =   90
      Width           =   1905
      Begin VB.OptionButton opt1Bit 
         Caption         =   "1 Bit - B/W"
         Height          =   240
         Left            =   135
         TabIndex        =   6
         Top             =   360
         Width           =   1320
      End
      Begin VB.OptionButton opt4Bit 
         Caption         =   "4 Bit - 16 Colors"
         Height          =   285
         Left            =   135
         TabIndex        =   3
         Top             =   675
         Width           =   1545
      End
      Begin VB.OptionButton opt8Bit 
         Caption         =   "8 Bit - 256 Colors"
         Height          =   285
         Left            =   135
         TabIndex        =   2
         Top             =   1035
         Width           =   1545
      End
      Begin VB.OptionButton opt24Bit 
         Caption         =   "24 Bit - True Color"
         Height          =   330
         Left            =   135
         TabIndex        =   1
         Top             =   1395
         Value           =   -1  'True
         Width           =   1725
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
CancelIt = True
Unload Form2
End Sub

Private Sub cmdOK_Click()
Form2.Hide
End Sub
