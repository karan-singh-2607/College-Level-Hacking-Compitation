VERSION 5.00
Begin VB.Form frmFixedFont 
   Caption         =   "Fixed Font"
   ClientHeight    =   1995
   ClientLeft      =   5940
   ClientTop       =   1935
   ClientWidth     =   3000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1995
   ScaleWidth      =   3000
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1500
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox TxtExample 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "FRM10.frx":0000
      Top             =   540
      Width           =   1815
   End
   Begin VB.CheckBox Chk_BoldFont 
      Caption         =   "Bold"
      Height          =   195
      Left            =   1920
      TabIndex        =   3
      Top             =   180
      Width           =   855
   End
   Begin VB.TextBox TextBox2 
      Height          =   304
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Sample"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   540
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Size"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   495
   End
End
Attribute VB_Name = "frmFixedFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmFixedFont.FRM
Option Explicit

Sub cmdCancel_Click()
  Unload Me
End Sub

Sub cmdOk_Click()
  gFontSize = TxtExample.FontSize
  gFontBold = TxtExample.FontBold
  gMarkerFontSet = gMarkerFontSet + 1
  Unload Me
End Sub

Sub TextBox2_Change()
  If Val(TextBox2) > 4 Then TxtExample.FontSize = Val(TextBox2)
End Sub

Sub Form_Load()
  If gFontSize = 0 Then gFontSize = 10
  TextBox2 = Format$(gFontSize)
  Chk_BoldFont = gFontBold And 1
  If gFontSize > 4 Then TxtExample.FontSize = gFontSize
  TxtExample.FontBold = gFontBold
End Sub

Sub Chk_BoldFont_Click()
  TxtExample.FontBold = Chk_BoldFont.Value
End Sub
