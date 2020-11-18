VERSION 5.00
Begin VB.Form frm17 
   Caption         =   "Labels"
   ClientHeight    =   8550
   ClientLeft      =   9090
   ClientTop       =   2775
   ClientWidth     =   4620
   Height          =   9060
   Left            =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   4620
   Top             =   2325
   Width           =   4740
   Begin VB.ListBox control1 
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Courier New"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   3870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4155
   End
End
' frm17.FRM
Option Explicit

Sub Form_Load ()
  sub1D03 control1
End Sub

Sub Form_Resize ()
  If  WindowState = 1 Then Exit Sub
  control1.Width = ScaleWidth
  control1.Height = ScaleHeight
End Sub
