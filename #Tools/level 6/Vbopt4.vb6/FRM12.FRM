VERSION 5.00
Begin VB.Form frmUserDefinedTypes 
   Caption         =   "User Defined Types"
   ClientHeight    =   4080
   ClientLeft      =   5385
   ClientTop       =   7695
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   5145
   Begin VB.ComboBox control1 
      Height          =   315
      Index           =   1
      Left            =   2580
      Sorted          =   -1  'True
      Style           =   2  'Dropdown-Liste
      TabIndex        =   4
      Top             =   420
      Width           =   2550
   End
   Begin VB.ListBox control2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Index           =   1
      Left            =   2580
      TabIndex        =   3
      Top             =   720
      Width           =   2550
   End
   Begin VB.CommandButton control3 
      Caption         =   "match"
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   60
      Width           =   795
   End
   Begin VB.ListBox control2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   2550
   End
   Begin VB.ComboBox control1 
      Height          =   315
      Index           =   0
      Left            =   0
      Sorted          =   -1  'True
      Style           =   2  'Dropdown-Liste
      TabIndex        =   0
      Top             =   420
      Width           =   2550
   End
   Begin VB.Label control4 
      Alignment       =   2  'Zentriert
      Caption         =   "Exe Types"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   2115
   End
   Begin VB.Label control4 
      Alignment       =   2  'Zentriert
      Caption         =   "Source Types"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Width           =   2115
   End
End
Attribute VB_Name = "frmUserDefinedTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmUserDefinedTypes.FRM
Option Explicit

Sub control3_Click()
Dim l001A As Integer
Dim l001C As Integer
Dim l001E As Integer
  l001A = control1(1).ListIndex
  If l001A < 0 Then Exit Sub
  l001A = control1(1).ItemData(l001A)
  l001C = control1(0).ListIndex
  If l001C < 0 Then Exit Sub
  l001E = control1(0).ItemData(l001C)
  If gEXEdefEntries2(l001E).M56AD = l001A Then Exit Sub
  If gv38BA(l001A).Size <> gEXEdefEntries2(l001E).Size Then
    MsgBox "Types have different size"
    Exit Sub
  End If
  Assume gEXEdefEntries2(l001E).subRef_M3C6A, "udt ref"
  sub0D9B gEXEdefEntries2(l001E).subRef_M3C6A, gv38BA(l001A).Name
  gEXEdefEntries2(l001E).M56AD = l001A
  control1(0).ListIndex = -1
  control1(0).ListIndex = l001C
End Sub

Sub control1_Click(Index%)
Dim l0038 As Integer
Dim l003A As Integer
  l0038 = control1(Index).ListIndex
  If l0038 < 0 Then Exit Sub
  l003A = control1(Index).ItemData(l0038)
  If Index Then
    sub0F08 l003A, control2(Index)
  Else
    sub1BE2 l003A, control2(Index)
  End If
End Sub

Sub Form_Load()
  FillCombo_userTypes control1(0)
  control1(0).ListIndex = control1(0).ListCount = 0
  
  FillCombo_userTypes2 control1(1)
  control1(1).ListIndex = control1(1).ListCount = 0
End Sub

Sub Form_Paint()
Static m0042 As Integer
Dim l0044
  If m0042 <> gMarkerFontSet Then
    For l0044 = 0 To 1
      control2(l0044).FontSize = gFontSize
      control2(l0044).FontBold = gFontBold
    Next
    m0042 = gMarkerFontSet
  End If
End Sub

Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Unload Me
End Sub

Sub Form_Resize()
Dim l0050 As Integer
Dim l0052 As Integer
Dim l0054
  If WindowState = 1 Then Exit Sub
  If ScaleHeight < 2 * control2(0).Top Then Exit Sub
  l0050 = ScaleWidth \ 2
  control1(1).Left = l0050
  control2(1).Left = l0050
  control4(0).Left = l0050
  control3.Left = l0050 - (control3.Width \ 2)
  l0052 = ScaleHeight - control2(0).Top
  For l0054 = 0 To 1
    control4(l0054).Width = l0050
    control1(l0054).Width = l0050
    control2(l0054).Width = l0050
    If l0052 > control1(l0054).Height Then control2(l0054).Height = l0052
  Next
End Sub
