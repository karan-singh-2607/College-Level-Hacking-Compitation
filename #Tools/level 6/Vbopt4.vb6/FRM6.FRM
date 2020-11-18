VERSION 5.00
Begin VB.Form frmOpenFileText 
   Caption         =   "Open File"
   ClientHeight    =   3165
   ClientLeft      =   570
   ClientTop       =   2415
   ClientWidth     =   4425
   LinkTopic       =   "Form4"
   ScaleHeight     =   3165
   ScaleWidth      =   4425
   Begin VB.CommandButton cmd_open 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   372
      Left            =   3480
      TabIndex        =   5
      Top             =   2700
      Width           =   852
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2520
      TabIndex        =   4
      Top             =   2700
      Width           =   852
   End
   Begin VB.ComboBox ComboBox1 
      Height          =   288
      Left            =   2520
      Style           =   2  'Dropdown-Liste
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.FileListBox FileListBox1 
      Height          =   2430
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.DirListBox DirListBox1 
      Height          =   2616
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2292
   End
   Begin VB.DriveListBox DriveListBox1 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2292
   End
End
Attribute VB_Name = "frmOpenFileText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmOpenFileText.FRM
Option Explicit

Sub cmdCancel_Click()
  frmOpenFileText.Tag = ""
  Hide
End Sub

Sub cmd_open_Click()
Dim l001E As String
Dim l0020 As String
  l001E = FileListBox1.Filename
  If l001E = "" Then ErrBeep2: FileListBox1.SetFocus: Exit Sub
  l0020 = FileListBox1.Path
  If Right$(l0020, 1) <> "\" Then l0020 = l0020 & "\"
  frmOpenFileText.Tag = l0020 & l001E
  Hide
End Sub

Sub ComboBox1_Change()
Dim l002A As Integer
  l002A = ComboBox1.ListIndex
  If l002A < 0 Then Exit Sub
  FileListBox1.Pattern = ComboBox1.List(l002A)
End Sub

Sub DirListBox1_Change()
  FileListBox1.Path = DirListBox1.List(DirListBox1.ListIndex)
End Sub

Sub DirListBox1_Click()
  FileListBox1.Path = DirListBox1.List(DirListBox1.ListIndex)
End Sub

Sub DriveListBox1_Change()
On Error Resume Next
  DirListBox1.Path = DriveListBox1.Drive
End Sub

Sub FileListBox1_DblClick()
  cmd_open.value = 1
End Sub

Sub Form_Activate()
  Caption = gc00AA & Tag
  FileListBox1.Path = DirListBox1.List(DirListBox1.ListIndex)
  FileListBox1.Pattern = Tag
End Sub
