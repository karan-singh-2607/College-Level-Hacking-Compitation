VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15135
   LinkTopic       =   "Form1"
   ScaleHeight     =   11085
   ScaleWidth      =   15135
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10620
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   List_gEntrys List1
   
   Me.Visible = True
   Do
      DoEvents
   Loop
End Sub
