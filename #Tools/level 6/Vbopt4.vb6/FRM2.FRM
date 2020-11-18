VERSION 5.00
Begin VB.Form frmText 
   Caption         =   "*.TXT"
   ClientHeight    =   2430
   ClientLeft      =   885
   ClientTop       =   1515
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2430
   ScaleWidth      =   4260
   Begin VB.TextBox control1 
      Height          =   1392
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Text            =   "FRM2.frx":0000
      Top             =   0
      Width           =   3852
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmText.FRM
Option Explicit

Sub Form_Resize()
  If WindowState = 1 Then Exit Sub
  control1.Width = ScaleWidth
  control1.Height = ScaleHeight
End Sub

Sub control1_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub
