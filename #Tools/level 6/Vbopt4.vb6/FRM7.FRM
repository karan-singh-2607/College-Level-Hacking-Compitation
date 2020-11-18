VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "DoDi's VB Tools"
   ClientHeight    =   1320
   ClientLeft      =   1125
   ClientTop       =   2295
   ClientWidth     =   4935
   Icon            =   "FRM7.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4935
   Begin VB.PictureBox control1 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   408
      Left            =   240
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   5
      Top             =   780
      Width           =   408
   End
   Begin VB.CommandButton control2 
      Caption         =   "WOW!"
      Default         =   -1  'True
      Height          =   552
      Left            =   3960
      TabIndex        =   4
      Top             =   180
      Width           =   852
   End
   Begin VB.Label control3 
      Caption         =   "e-mail"
      Height          =   195
      Left            =   960
      TabIndex        =   6
      Top             =   1020
      Width           =   735
   End
   Begin VB.Label control4 
      Caption         =   "DoDi@compuserve.com"
      Height          =   195
      Left            =   1800
      TabIndex        =   7
      Top             =   1020
      Width           =   3075
   End
   Begin VB.Label control5 
      Caption         =   "2.0"
      Height          =   192
      Left            =   1080
      TabIndex        =   3
      Top             =   180
      Width           =   2292
   End
   Begin VB.Label control6 
      Caption         =   "Dr. H.-P. Diettrich"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label control7 
      Caption         =   "Copyright 1997"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label control8 
      AutoSize        =   -1  'True
      Caption         =   "Version"
      Height          =   192
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   852
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmAbout.FRM
Option Explicit

Sub control2_Click()
44000
  Unload Me
End Sub

Sub Form_Load()
Dim l001A As String

44010
  control1.Picture = g_MainForm.Icon
  Caption = "DoDi's " & App.Title
If g_IsDemoVersion Then
  l001A = " Demo"
ElseIf LicEnableProFeatures Then
  l001A = " Pro"
ElseIf LicEnableLiteFeatures Then
  l001A = " Lite"
End If
  control5 = gc02C4 & gc0006 & l001A
End Sub
