VERSION 5.00
Begin VB.Form frm3 
   Caption         =   "*.MAK"
   ClientHeight    =   2955
   ClientLeft      =   8700
   ClientTop       =   1500
   ClientWidth     =   4590
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   2955
   ScaleWidth      =   4590
   Begin VB.TextBox control1 
      Height          =   592
      Left            =   192
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1792
      Visible         =   0   'False
      Width           =   3536
   End
   Begin VB.CommandButton cmdSync 
      Caption         =   "sync"
      Height          =   286
      Left            =   2224
      TabIndex        =   4
      Top             =   64
      Width           =   858
   End
   Begin VB.ListBox control3 
      Height          =   448
      Left            =   0
      MousePointer    =   1  'Pfeil
      TabIndex        =   2
      Top             =   832
      Width           =   3504
   End
   Begin VB.CommandButton control4 
      Caption         =   "init"
      Height          =   252
      Left            =   121
      TabIndex        =   3
      Top             =   2530
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.ComboBox control5 
      Height          =   336
      Left            =   64
      Style           =   2  'Dropdown-Liste
      TabIndex        =   1
      Top             =   416
      Width           =   3008
   End
   Begin VB.ComboBox control6 
      Height          =   336
      Left            =   64
      Style           =   2  'Dropdown-Liste
      TabIndex        =   0
      Top             =   64
      Width           =   2064
   End
End
Attribute VB_Name = "frm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frm3.FRM
Option Explicit
Const mc0044 = 35 ' &H23%
Const mc004E = 34 ' &H22%
Const mc0080 = 2 ' &H2%

Sub cmd_init_Click()
  frmScan.SetFocus
  File_Name gv02EC_MAKName
  Caption = gExeNameOnly & Right$(gv02EC_MAKName, 4)
  sub0AD4
  sub0951 control6
  control6.SetFocus
  control6.ListIndex = control6.ListCount = 0
End Sub

Sub cmdSync_Click()
Dim l0030 As Integer
Dim l0032 As Integer
Dim l0034 As Integer
Dim l0036 As Integer
' Const mc0044 = 35 ' &H23%
' Const mc004E = 34 ' &H22%
  l0036 = g_WorkFrm.control6.ListIndex
  If l0036 < 0 Then Exit Sub
  l0030 = g_WorkFrm.control6.ItemData(l0036)
  l0036 = g_WorkFrm.control5.ListIndex
  If l0036 < 0 Then
    Exit Sub
  End If
  l0032 = g_WorkFrm.control5.ItemData(l0036)
  l0036 = l0030
  control6.ListIndex = l0036 - 1
  DoEvents
  If l0032 < 0 Then
    Select Case gReloc_Arr(-l0032).M36B3_count
    Case mc0044: l0034 = gv094A(l0036).M3F1B
    Case mc004E: l0034 = gv094A(l0036).M3F1B + 1
    Case Else
      Exit Sub
    End Select
  Else
    l0034 = gEXEdef_Subs(l0032).M4289
  End If
  For l0036 = control5.ListCount - 1 To 0 Step -1
    If control5.ItemData(l0036) = l0034 Then
      control5.ListIndex = l0036
      Exit For
    End If
  Next
End Sub

Sub ComboBox2_Click()
Dim l005E As Integer
  l005E = control6.ListIndex + 1
  If l005E <= 0 Then Exit Sub
  control5.Clear
  sub0963 control5, l005E, True
  If ActiveControl Is control6 Then
    If control5.ListCount Then
      control5.ListIndex = 0
      control5.SetFocus
    End If
  End If
End Sub

Sub ComboBox1_Click()
Dim l0066 As Integer
Dim l0068 As String
Dim l006A As Integer
Dim l006C As Long
Dim l006E As Integer
Dim l0070 As Integer
' Const mc0080 = 2 ' &H2%
Static m0084 As Integer
  control3.Clear
  l0066 = control5.ListIndex
  If l0066 < 0 Then Exit Sub
  l006E = control5.ItemData(l0066)
  l0070 = gv0AD6(l006E).M403B
  l006C = gv0AD6(l006E).M4045
  If l006C <= gv0AD6(l006E - 1).M4045 Then ErrBeep: Exit Sub
  l006A = fn171A(l0070, l006E, mc0080)
  Do While Seek(l006A) < l006C
    Line Input #l006A, l0068
    If Len(l0068) Then
      control3.AddItem l0068
      Exit Do
    End If
  Loop
  While Seek(l006A) < l006C
    Line Input #l006A, l0068
    control3.AddItem l0068
  Wend
  l006A = fn171A(l0070, l006E, m0084)
End Sub

Sub control1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Select Case Button
  Case 2
    Set gv02D6 = control3
    PopupMenu frmMain.Menu_popup
  End Select
End Sub

Private Sub control2_Click()

End Sub

Sub Form_Load()
  Icon = frmMain.Icon
End Sub

Sub Form_Resize()
  If WindowState = 1 Then Exit Sub
  control3.Width = ScaleWidth - control3.Left * 2
  control3.Height = ScaleHeight - control3.Top
  control1.Width = ScaleWidth - control1.Left * 2
End Sub

Sub ListBox1_Click()
Dim l00B4 As Integer
  l00B4 = control3.ListIndex
  If l00B4 < 0 Then
    frmMain.CmdOpenProject = ""
  Else
    frmMain.CmdOpenProject = Trim$(control3.List(l00B4))
  End If
End Sub

Sub ListBox1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    Set gv02D6 = control3
    PopupMenu frmMain.Menu_popup
  End If
End Sub
