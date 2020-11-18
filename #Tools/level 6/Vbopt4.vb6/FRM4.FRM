VERSION 5.00
Begin VB.Form frmResources 
   Caption         =   "Resources"
   ClientHeight    =   2595
   ClientLeft      =   210
   ClientTop       =   3495
   ClientWidth     =   4350
   LinkTopic       =   "Form2"
   ScaleHeight     =   2595
   ScaleWidth      =   4350
   Begin VB.ComboBox ComboBox_Res1 
      Height          =   264
      Left            =   0
      Style           =   2  'Dropdown-Liste
      TabIndex        =   3
      Top             =   121
      Width           =   1826
   End
   Begin VB.ListBox ListBox_Res1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   0
      TabIndex        =   2
      Top             =   420
      Width           =   3615
   End
   Begin VB.ComboBox ComboBox_Res2 
      Height          =   264
      Left            =   1936
      Style           =   2  'Dropdown-Liste
      TabIndex        =   1
      Top             =   121
      Width           =   2310
   End
   Begin VB.CommandButton cmd_Refresh 
      Caption         =   "Refresh"
      Height          =   312
      Left            =   120
      TabIndex        =   0
      Top             =   6840
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Menu Menu_Dump 
      Caption         =   "&Dump"
      Begin VB.Menu mi_Dump_Bytes 
         Caption         =   "&Bytes"
      End
      Begin VB.Menu mi_Dump_Words 
         Caption         =   "&Words"
      End
      Begin VB.Menu mi_Dump_Pretty 
         Caption         =   "&Pretty"
      End
   End
   Begin VB.Menu Menu_Show 
      Caption         =   "&Show"
      Begin VB.Menu mi_Show_Bytes 
         Caption         =   "&Bytes"
      End
      Begin VB.Menu mi_Show_Words 
         Caption         =   "&Words"
      End
      Begin VB.Menu mi_Show_Pretty 
         Caption         =   "&Pretty"
      End
   End
End
Attribute VB_Name = "frmResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frm4.FRM
Option Explicit
Dim m001A As Integer
Dim m0032 As Integer
Const mc004C = 1 ' &H1%
Dim m004E As Integer
Const mc0050 = 2 ' &H2%

Sub cmd_Refresh_Click()
  ComboBox_Res1.Clear
  ComboBox_Res2.Clear
  ListBox_Res1.Clear
  gv02B4 = 0
  sub03DB ComboBox_Res1
  ComboBox_Res1.ListIndex = ComboBox_Res1.ListCount = 0
End Sub

Sub Form_Paint()
  If m0032 <> gMarkerFontSet Then
    ListBox_Res1.FontSize = gFontSize
    ListBox_Res1.FontBold = gFontBold
    m0032 = gMarkerFontSet
  End If
End Sub

Sub Form_Resize()
  If WindowState = 1 Then Exit Sub
  ListBox_Res1.Width = ScaleWidth
  ListBox_Res1.Height = ScaleHeight - ListBox_Res1.Top
End Sub

Sub mi_Dump_Words_Click()
  sub01D5 gv02B4, mc004C
End Sub

Sub mi_Dump_Bytes_Click()
  sub01D5 gv02B4, m004E
End Sub

Sub mi_Dump_Pretty_Click()
  sub01D5 gv02B4, mc0050
End Sub

Sub mi_Show_Bytes_Click()
  m001A = m004E
  sub28DF
End Sub

Sub mi_Show_Pretty_Click()
  m001A = mc0050
  sub28DF
End Sub

Sub mi_Show_Words_Click()
  m001A = mc004C
  sub28DF
End Sub

Sub ComboBox_Res2_Click()
Dim l0052 As Integer
  l0052 = ComboBox_Res2.ListIndex
  If l0052 < 0 Then Exit Sub
  l0052 = ComboBox_Res2.ItemData(l0052)
  sub025C l0052
  sub28DF
End Sub

Sub ComboBox_Res1_Click()
Dim l0054 As Integer
Dim l0058 As Integer
  ComboBox_Res2.Clear
  l0054 = ComboBox_Res1.ListIndex
  If l0054 < 0 Then Exit Sub
  l0058 = ComboBox_Res1.ItemData(l0054)
  sub03C9 l0058, ComboBox_Res2
  ComboBox_Res2.ListIndex = ComboBox_Res2.ListCount = 0
End Sub

Sub sub28DF()
  SetDecompilerListBox ListBox_Res1
  sub0251 gv02B4, m001A
  UnlinkDecompilerListBox
End Sub
