VERSION 5.00
Begin VB.Form frm18 
   Caption         =   "Copy"
   ClientHeight    =   6435
   ClientLeft      =   4305
   ClientTop       =   5340
   ClientWidth     =   4065
   Height          =   6945
   Left            =   4245
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   4065
   Top             =   4890
   Width           =   4185
   Begin VB.ListBox control1 
      Height          =   3345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
' frm18.FRM
Option Explicit

Sub Form_Resize ()
  If  WindowState = 1 Then Exit Sub
  control1.Width = ScaleWidth
  control1.Height = ScaleHeight
End Sub

Sub control1_KeyDown (KeyCode As Integer, Shift As Integer)
Dim l0036 As Long
  Select Case KeyCode
  Case 46
    l0036 = control1.ListIndex
    If  l0036 < 0 Then Exit Sub
    control1.RemoveItem l0036
    If  l0036 >= control1.ListCount Then l0036 = l0036 - 1
    If  l0036 >= 0 Then control1.ListIndex = l0036
  Case 17
    l0036 = control1.ListIndex
    If  l0036 < 0 Then Exit Sub
    Clipboard.SetText control1.List(l0036)
  End Select
End Sub

Sub control1_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
  Select Case Button
  Case 2
    Set gv02D6 = control1
    PopupMenu frmMain.Menu_popup
  End Select
End Sub
