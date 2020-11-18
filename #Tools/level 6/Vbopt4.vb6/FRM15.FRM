VERSION 5.00
Begin VB.Form frmDecompiler 
   Caption         =   "Decompiler"
   ClientHeight    =   2445
   ClientLeft      =   10515
   ClientTop       =   1920
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   4590
   Begin VB.CommandButton control1 
      Caption         =   "decomp"
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton control1 
      Caption         =   "add"
      Height          =   255
      Index           =   1
      Left            =   1020
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton control1 
      Caption         =   "init"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.ListBox Lst_OutPut 
      Height          =   1425
      Left            =   0
      TabIndex        =   0
      Top             =   180
      Width           =   4395
   End
End
Attribute VB_Name = "frmDecompiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmDecompiler.FRM
Option Explicit

Sub control1_Click(Index%)
  Select Case Index
  Case 0 'Init
    Lst_OutPut.Clear
  
  Case 1 'Add
   'Add Data
    Lst_OutPut.AddItem H32(gVACmdStart) & "  " & fn187B() & " " & fn1820(0)
    
   'Add Offset
    Lst_OutPut.ItemData(Lst_OutPut.NewIndex) = gVACmdStart
    
    If Len(gDisOutputTxt) Then
      Lst_OutPut.AddItem gDisOutputTxt
    End If
  
  Case 2 'decomp
    SetDecompilerListBox Lst_OutPut
    If gv16BC > 0 Then sub028D gv16BC, True
    UnlinkDecompilerListBox
  End Select
  gDisOutputTxt = ""
End Sub

Sub Form_Resize()
  If WindowState = 1 Then Exit Sub
  Lst_OutPut.Width = ScaleWidth
  Lst_OutPut.Height = ScaleHeight - Lst_OutPut.Top
End Sub

Sub Lst_OutPut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    Set gv02D6 = Lst_OutPut
    PopupMenu frmMain.Menu_popup
  End If
End Sub
