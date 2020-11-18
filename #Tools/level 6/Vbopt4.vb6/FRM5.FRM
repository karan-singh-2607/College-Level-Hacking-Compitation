VERSION 5.00
Begin VB.Form frmProjectDir 
   Caption         =   "Select Project Directory"
   ClientHeight    =   4005
   ClientLeft      =   885
   ClientTop       =   1515
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   4200
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   372
      Left            =   3000
      TabIndex        =   4
      Top             =   1380
      Width           =   1092
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   372
      Left            =   3000
      TabIndex        =   3
      Top             =   600
      Width           =   1092
   End
   Begin VB.TextBox TextBox1 
      Height          =   304
      Left            =   2880
      TabIndex        =   2
      Top             =   180
      Width           =   1212
   End
   Begin VB.DirListBox DirListBox1 
      Height          =   3264
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2652
   End
   Begin VB.DriveListBox DriveListBox1 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   2652
   End
End
Attribute VB_Name = "frmProjectDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frm5.FRM
Option Explicit

Sub cmdOk_Click()
  gProjectDirPath = DirListBox1.Path
  'Append "\"
  If Right$(gProjectDirPath, 1) <> "\" Then gProjectDirPath = gProjectDirPath & "\"
  Hide
End Sub

Sub cmdNew_Click()
Dim l0022 As String
Dim l0024 As String
  l0022 = TextBox1
  If Len(l0022) = 0 Then ErrBeep2: Exit Sub
  l0024 = DirListBox1.List(DirListBox1.ListIndex)
  If Right$(l0024, 1) <> "\" Then l0024 = l0024 & "\"
  l0024 = l0024 & l0022
On Error Resume Next
  MkDir l0024
  DirListBox1.Path = l0024
  If Err Then ErrBeep2: Exit Sub
  gProjectDirPath = l0024 & "\"
End Sub

Sub control6_Change()
  DirListBox1.Path = DriveListBox1.Drive
End Sub
