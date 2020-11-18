VERSION 5.00
Begin VB.Form frmStatements 
   Caption         =   "Statements"
   ClientHeight    =   3645
   ClientLeft      =   10470
   ClientTop       =   1725
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   4980
   Begin VB.CommandButton cmd_final 
      Caption         =   "final"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   60
      Width           =   675
   End
   Begin VB.CommandButton cmd_pretty 
      Caption         =   "pretty"
      Height          =   255
      Left            =   1020
      TabIndex        =   3
      Top             =   60
      Width           =   735
   End
   Begin VB.CommandButton cmd_raw 
      Caption         =   "raw"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   675
   End
   Begin VB.CommandButton cmd_init 
      Caption         =   "init"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ListBox ListBox 
      Height          =   2595
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4815
   End
End
Attribute VB_Name = "frmStatements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frm16.FRM
Option Explicit
Const mc002C = 4096 ' &H1000%
Const mc0030 = 64 ' &H40%

Sub cmd_final_Click()
Dim l001A
' Const mc002C = 4096 ' &H1000%
' Const mc0030 = 64 ' &H40%
  SetDecompilerListBox ListBox
  For l001A = 1 To gMetaCodeLinesCount
    If gMetaCodeLines(l001A).M36B3_count And mc002C Then
    Else
      DisOutput Hex$(gMetaCodeLines(l001A).VOffset) & fn085C(gMetaCodeLines(l001A).M36B3_count) & " "
      Select Case gMetaCodeLines(l001A).M36B3_count And &HFF
      Case mc0030
        If gMetaCodeLines(l001A).M5742 Then
          DisOutput fn1B19(gMetaCodeLines(l001A).M5742) & " = "
        Else
          DisOutput "?="
        End If
      End Select
      DisOutput gMetaCodeLines(l001A).M574B: WriteOutputTxt
    End If
  Next
  UnlinkDecompilerListBox
End Sub

Sub cmd_init_Click()
Dim l0034
  If Not Me.Visible Then
    Exit Sub
  End If
  SetDecompilerListBox ListBox
  For l0034 = 1 To gMetaCodeLinesCount
    If gMetaCodeLines(l0034).M36B3_count And mc002C Then DisOutput "'"
    DisOutput Hex$(gMetaCodeLines(l0034).VOffset) & fn085C(gMetaCodeLines(l0034).M36B3_count) & " "
    Select Case gMetaCodeLines(l0034).M36B3_count And &HFF
    Case mc0030
      If gMetaCodeLines(l0034).M5742 Then
        DisOutput fn1B19(gMetaCodeLines(l0034).M5742) & " = "
      Else
        DisOutput "?="
      End If
    End Select
    DisOutput gMetaCodeLines(l0034).M574B: WriteOutputTxt
  Next
  UnlinkDecompilerListBox
End Sub

Sub cmd_pretty_Click()
Dim l0038
  SetDecompilerListBox ListBox
  For l0038 = 1 To gMetaCodeLinesCount
    If gMetaCodeLines(l0038).M36B3_count And mc002C Then
    Else
      DisOutput Hex$(gMetaCodeLines(l0038).VOffset) & fn085C(gMetaCodeLines(l0038).M36B3_count) & " "
      Select Case gMetaCodeLines(l0038).M36B3_count And &HFF
      Case mc0030
        If gMetaCodeLines(l0038).M5742 Then
          DisOutput fn1B19(gMetaCodeLines(l0038).M5742) & " = "
        Else
          DisOutput "?="
        End If
      End Select
      DisOutput gMetaCodeLines(l0038).M574B: WriteOutputTxt
    End If
  Next
  UnlinkDecompilerListBox
End Sub

Sub cmd_raw_Click()
Dim l003A
  SetDecompilerListBox ListBox
  For l003A = 1 To gMetaCodeLinesCount
    If gMetaCodeLines(l003A).M36B3_count And mc002C Then DisOutput "'"
    DisOutput Hex$(gMetaCodeLines(l003A).VOffset) & fn085C(gMetaCodeLines(l003A).M36B3_count) & " "
    Select Case gMetaCodeLines(l003A).M36B3_count And &HFF
    Case mc0030
      If gMetaCodeLines(l003A).M5742 Then
        DisOutput fn1B19(gMetaCodeLines(l003A).M5742) & " = "
      Else
        DisOutput "?="
      End If
    End Select
    DisOutput gMetaCodeLines(l003A).M574B: WriteOutputTxt
  Next
  UnlinkDecompilerListBox
End Sub

Sub Form_Resize()
  If WindowState = 1 Then Exit Sub
  ListBox.Width = ScaleWidth - ListBox.Left
  ListBox.Height = ScaleHeight - ListBox.Top
End Sub

Sub ListBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    Set gv02D6 = ListBox
    PopupMenu frmMain.Menu_popup
  End If
End Sub
