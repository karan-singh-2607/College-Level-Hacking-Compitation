VERSION 5.00
Begin VB.Form frmScopes 
   Caption         =   "Scopes"
   ClientHeight    =   3525
   ClientLeft      =   9795
   ClientTop       =   1665
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   5355
   Begin VB.CommandButton cmd_temp 
      Caption         =   "temp"
      Height          =   255
      Left            =   4500
      TabIndex        =   15
      Top             =   360
      Width           =   735
   End
   Begin VB.CheckBox chk_raw 
      Caption         =   "raw"
      Height          =   195
      Left            =   3720
      TabIndex        =   13
      Top             =   360
      Width           =   675
   End
   Begin VB.CommandButton cmd_match 
      Caption         =   "match"
      Height          =   255
      Left            =   2100
      TabIndex        =   12
      Top             =   660
      Width           =   795
   End
   Begin VB.ComboBox ComboBox 
      Height          =   315
      Left            =   840
      Style           =   1  'Einfaches Kombinationsfeld
      TabIndex        =   10
      Text            =   "unknown"
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox TextBox 
      Height          =   304
      Left            =   45
      TabIndex        =   6
      Top             =   330
      Width           =   615
   End
   Begin VB.CommandButton SpinButton 
      Height          =   300
      Left            =   660
      TabIndex        =   16
      Top             =   315
      Width           =   195
   End
   Begin VB.ListBox ListBox1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Index           =   1
      Left            =   2460
      TabIndex        =   2
      Top             =   900
      Width           =   2430
   End
   Begin VB.CommandButton cmd_init 
      Caption         =   "init"
      Height          =   231
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Visible         =   0   'False
      Width           =   737
   End
   Begin VB.ListBox ListBox1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   900
      Width           =   2430
   End
   Begin VB.Label ListBox2 
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Lbl_to 
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   60
      Width           =   3015
   End
   Begin VB.Label Lbl_1 
      Caption         =   "to"
      Height          =   255
      Left            =   2700
      TabIndex        =   9
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Lbl_2 
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Lbl_Scope 
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   360
      Width           =   675
   End
   Begin VB.Label Lbl_SourceVars 
      Caption         =   "Scope"
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Lbl_ExeVars 
      Alignment       =   2  'Zentriert
      Caption         =   "Source Vars"
      Height          =   195
      Index           =   1
      Left            =   3240
      TabIndex        =   4
      Top             =   660
      Width           =   1635
   End
   Begin VB.Label Lbl_ExeVars 
      Alignment       =   2  'Zentriert
      Caption         =   "Exe Vars"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   660
      Width           =   1695
   End
End
Attribute VB_Name = "frmScopes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frm11.FRM
Option Explicit
Dim m001A As Integer
Dim m001C As Integer
Dim m001E As Integer
Dim m0020 As Integer
Dim m0022 As Integer
Const mc0024 = 1 ' &H1%
Const mc0026 = 2 ' &H2%
Const mc0028 = 3 ' &H3%
Const mc002A = 4 ' &H4%
Const mc002C = 5 ' &H5%
Const mc0062 = 4096 ' &H1000%
Const mc00BE = -8192 ' &HE000%
Const mc00C0 = 24576 ' &H6000%
Const mc00CA = 16384 ' &H4000%
Const m00CE = 0 ' &H0%
Const mc00D2 = 8192 ' &H2000%
Const mc00E8 = -8192 ' &HE000%
Const mc00F2 = -32768 ' &H8000%

Sub chk_raw_Click()
  sub2C69 m001C
End Sub

Sub cmd_temp_Click()
Dim l005E As Integer
Dim l0060 As Integer
' Const mc0062 = 4096 ' &H1000%
  l005E = ListBox1(0).ListIndex
  If l005E < 0 Then Exit Sub
  l0060 = ListBox1(0).ItemData(l005E)
  If gEXEdefEntries(l0060).M689A >= 0 Then Exit Sub
  gEXEdefEntries(l0060).M6894.M36B3_count = gEXEdefEntries(l0060).M6894.M36B3_count Xor mc0062
  gIs3158 = True
  sub2C69 m001C
End Sub

Sub control1_Click()
Dim l0066 As Integer
  If Not ActiveControl Is ComboBox Then Exit Sub
  l0066 = ComboBox.ListIndex
  Select Case l0066
  Case m0022
    Exit Sub
  Case mc002A
    m001C = gL3Min
  Case mc0028
    m001C = gL1Min
  Case mc0024
    m001C = gL0Min
  Case mc002C
    m001C = fn1C9B(1, False)
  Case mc0026
    m001C = gEXEdef6_Num_scopes
  End Select
  TextBox = Format$(m001C)
End Sub

Sub control3_Click()
Dim l002E As Integer
Dim l0030 As Integer
Dim l0034 As Integer
Dim l003A
  l002E = ListBox1(0).ListIndex
  If l002E < 0 Then ErrBeep2: Exit Sub
  l002E = ListBox1(0).ItemData(l002E)
  l0030 = ListBox1(1).ListIndex
  If l0030 < 0 Then ErrBeep2: Exit Sub
  l0030 = ListBox1(1).ItemData(l0030)
  Assume ComboBox.ListIndex = mc002A, "lcl scope"
  l0034 = gEXEdefEntries(l002E).M689A - gv36AE(l0030).hFile
  If l0034 = 0 Then
    MsgBox "already matched???"
    Exit Sub
  End If
  For l003A = l0030 To gv36E4
    gv36AE(l003A).hFile = gv36AE(l003A).hFile + l0034
  Next
  gv16B8 = True
  sub1740
  sub2C69 m001C
End Sub

Sub control9_SpinDown()
  If m001C > 1 Then
    Select Case g_gbl(m001C).M36B3_count And mc00BE
    Case mc00D2
      m001C = gL2Max
    Case m00CE
      If chk_raw Then
        m001C = m001C - 1
      Else
        m001C = gL2Max
      End If
    Case mc00CA
      If chk_raw Then
        m001C = m001C - 1
      Else
        m001C = 1
      End If
    Case mc00C0
      If chk_raw Then
        m001C = m001C - 1
      Else
ErrBeep2:         Exit Sub
      End If
    Case Else
      m001C = m001C - 1
    End Select
    TextBox = Format$(m001C)
    DoEvents
  Else
    ErrBeep2
  End If
End Sub

Sub control9_SpinUp()
  Select Case g_gbl(m001C).M36B3_count And mc00BE
  Case mc00D2
    m001C = gL4Max + 1
  Case m00CE
    If chk_raw Then
      m001C = m001C + 1
      If m001C > gL3Max Then m001C = gL3Min
    Else
      m001C = gL4Max + 1
    End If
  Case mc00CA
    If chk_raw Then
      m001C = m001C + 1
    Else
      m001C = gL3Min
    End If
  Case mc00C0
    If chk_raw Then
      m001C = m001C + 1
    Else
      m001C = gL1Min
    End If
  Case Else
    m001C = m001C + 1
  End Select
  TextBox = Format$(m001C)
  DoEvents
End Sub

Sub Form_Load()
  ComboBox.AddItem "unknown"
  ComboBox.AddItem "global"
  ComboBox.AddItem "GUID"
  ComboBox.AddItem "static"
  ComboBox.AddItem "local"
  ComboBox.AddItem "UDT"
End Sub

Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Unload Me
End Sub

Sub Form_Resize()
Dim l0082 As Integer
Dim l0084 As Integer
Dim l0086
  If WindowState = 1 Then Exit Sub
  m001A = True
  If gv097C Then
    l0082 = ScaleWidth \ 2
    cmd_match.Left = l0082 - (cmd_match.Width \ 2)
  Else
    l0082 = ScaleWidth
    cmd_match.Left = l0082
  End If
  Lbl_ExeVars(1).Left = l0082
  ListBox1(1).Left = l0082
  l0084 = ScaleHeight - ListBox1(0).Top
  For l0086 = 0 To -(gv097C > 0)
    Lbl_ExeVars(l0086).Width = l0082
    ListBox1(l0086).Width = l0082
    If l0084 > ComboBox.Height Then ListBox1(l0086).Height = l0084
  Next
  l0082 = ScaleWidth - Lbl_to.Left
  If l0082 > 0 Then Lbl_to.Width = l0082
End Sub

Sub sub2C69(p00AE As Integer)
' Const mc00BE =-8192 ' &HE000%
' Const mc00C0 = 24576 ' &H6000%
' Const mc00CA = 16384 ' &H4000%
' Const m00CE = 0 ' &H0%
' Const mc00D2 = 8192 ' &H2000%
' Const mc00E8 =-8192 ' &HE000%
' Const mc00F2 =-32768 ' &H8000%
  If Screen.ActiveForm Is Me Then TextBox.SetFocus
  If p00AE = 0 Then TextBox = 1: Exit Sub
  If p00AE > gEXEdef6_Num_scopes Then TextBox = gEXEdef6_Num_scopes: Exit Sub
  m001E = 0
  ListBox1(1).Clear
  Lbl_to = fn104E(p00AE)
  Select Case g_gbl(p00AE).M36B3_count And mc00BE
  Case mc00C0
    m0020 = mc0024
    m001C = p00AE
    m001E = p00AE - gLWhatEver
    If chk_raw Then
      sub0FC7 m001C, ListBox1(0), True
      sub1AC1 m001E, ListBox1(1)
    Else
      sub0FA2 ListBox1(0)
      sub1AD3 ListBox1(1)
    End If
  Case mc00CA
    m0020 = mc0028
    If chk_raw Then
      m001C = p00AE
    Else
      m001C = gL1Min
    End If
    sub0FC7 m001C, ListBox1(0), True
    If m001C = gL1Min Then
      m001E = p00AE - gL1Max
      sub1AF6 m001E, ListBox1(1)
    End If
  Case m00CE
    m0020 = mc002A
    If g_gbl(p00AE).M36B3_count And &HFFF <> gv16BC Then
      If chk_raw = 0 Then
        TextBox = gL3Min
        Exit Sub
      End If
    End If
    m001C = p00AE
    sub0FC7 m001C, ListBox1(0), True
  Case mc00D2
    m0020 = mc002A
L55E18:
    If p00AE <> gL3Min Then
      TextBox = gL3Min
      Exit Sub
    End If
    Lbl_Scope = toHex16(gv315A): Lbl_2 = toHex16(gv315C)
    sub0FB5 gv16BC, ListBox1(0)
    m001C = gL3Min
    m001E = gv16B6
    sub1AE5 m001E, ListBox1(1)
    GoTo L56090
  Case mc00E8
L55ED4:
    m0020 = mc002C
    m001C = p00AE
    sub0FC7 m001C, ListBox1(0), True
    m001E = fn1C80(m001C)
    If m001E Then
      If m001E <= g_userTypesCount Then
        m001E = fn1CA9(m001E)
      Else
        m001E = 0
      End If
    End If
    If m001E Then
      If m001E <= gv38EC Then sub0F08 m001E, ListBox1(1)
    End If
  Case mc00F2
    m0020 = mc0026
    m001C = p00AE
    DbgMileStone "show GUID"
    sub0FC7 m001C, ListBox1(0), True
  Case Else
    If g_gbl(p00AE).M36B3_count And mc00F2 Then GoTo L55ED4
    m0020 = m0022
    m001C = p00AE
    sub0FC7 m001C, ListBox1(0), True
  End Select
  Lbl_Scope = "0"
  Lbl_2 = g_gbl(m001C).Size
L56090:
  ListBox2 = toHex16(g_gbl(m001C).M36B3_count)
  ComboBox.ListIndex = m0020
End Sub

Sub TextBox_Change()
  If Not m001A Then Exit Sub
  If Not ActiveControl Is TextBox Then Exit Sub
  m001C = Val(TextBox)
  sub2C69 m001C
End Sub

