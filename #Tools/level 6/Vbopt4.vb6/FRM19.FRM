VERSION 5.00
Begin VB.Form frm19 
   Caption         =   "Collection"
   ClientHeight    =   4005
   ClientLeft      =   8190
   ClientTop       =   1590
   ClientWidth     =   5490
   Height          =   4515
   Left            =   8130
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   5490
   Top             =   1140
   Width           =   5610
   Begin VB.ListBox control1 
      Height          =   3360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4688
   End
End
' frm19.FRM
Option Explicit

Sub Form_Resize ()
  If  WindowState = 1 Then Exit Sub
  control1.Width = ScaleWidth
  control1.Height = ScaleHeight
End Sub

Sub control1_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
  Select Case Button
  Case 2
    Set gv02D6 = control1
    PopupMenu frmMain.Menu_popup
  End Select
End Sub
