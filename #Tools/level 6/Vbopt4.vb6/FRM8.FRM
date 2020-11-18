VERSION 5.00
Begin VB.Form frmScan 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Scanning - Please Wait"
   ClientHeight    =   3510
   ClientLeft      =   1500
   ClientTop       =   1830
   ClientWidth     =   6765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6765
   Begin VB.Label control1 
      Alignment       =   2  'Zentriert
      Caption         =   "DoDi's VB4 Disassembler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   120
      TabIndex        =   4
      Top             =   435
      Width           =   6540
   End
   Begin VB.Label control2 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "MODULE1.BAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2280
      TabIndex        =   3
      Top             =   2520
      Width           =   4260
   End
   Begin VB.Label control3 
      Alignment       =   1  'Rechts
      Caption         =   "Module"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   2535
      Width           =   2055
   End
   Begin VB.Label control4 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Conditions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2280
      TabIndex        =   1
      Top             =   1560
      Width           =   4260
   End
   Begin VB.Label control5 
      Alignment       =   1  'Rechts
      Caption         =   "Scanning"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   2070
   End
End
Attribute VB_Name = "frmScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmScan.FRM
Option Explicit
