VERSION 5.00
Begin VB.Form frm14 
   Caption         =   "Variable"
   ClientHeight    =   3615
   ClientLeft      =   11415
   ClientTop       =   2835
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   4560
   Begin VB.Frame control1 
      Caption         =   "Array"
      Height          =   1215
      Left            =   60
      TabIndex        =   8
      Top             =   2340
      Width           =   4455
      Begin VB.Label control2 
         BorderStyle     =   1  'Fest Einfach
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.Label control3 
         BorderStyle     =   1  'Fest Einfach
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   23
         Top             =   540
         Width           =   735
      End
      Begin VB.Label control4 
         BorderStyle     =   1  'Fest Einfach
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   24
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label control5 
         BorderStyle     =   1  'Fest Einfach
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   25
         Top             =   540
         Width           =   2775
      End
      Begin VB.Label control6 
         BorderStyle     =   1  'Fest Einfach
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   26
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label control7 
         BorderStyle     =   1  'Fest Einfach
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox control8 
      Height          =   304
      Index           =   2
      Left            =   2880
      TabIndex        =   22
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox control8 
      Height          =   304
      Index           =   1
      Left            =   1620
      TabIndex        =   17
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton control9 
      Height          =   255
      Left            =   2820
      TabIndex        =   30
      ToolTipText     =   "SpinButton"
      Top             =   360
      Width           =   195
   End
   Begin VB.TextBox control8 
      Height          =   304
      Index           =   0
      Left            =   780
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox control10 
      Height          =   304
      Left            =   1620
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label control2 
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Index           =   0
      Left            =   780
      TabIndex        =   29
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label control3 
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Index           =   0
      Left            =   780
      TabIndex        =   27
      Top             =   1740
      Width           =   735
   End
   Begin VB.Label control11 
      Caption         =   "max"
      Height          =   255
      Left            =   2400
      TabIndex        =   21
      Top             =   1080
      Width           =   435
   End
   Begin VB.Label control12 
      Caption         =   "Name"
      Height          =   255
      Left            =   60
      TabIndex        =   10
      Top             =   720
      Width           =   615
   End
   Begin VB.Label control13 
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Left            =   3840
      TabIndex        =   20
      Top             =   360
      Width           =   555
   End
   Begin VB.Label control14 
      Caption         =   "owner"
      Height          =   255
      Left            =   3060
      TabIndex        =   19
      Top             =   360
      Width           =   675
   End
   Begin VB.Label control15 
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Left            =   780
      TabIndex        =   18
      Top             =   720
      Width           =   735
   End
   Begin VB.Label control4 
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Index           =   0
      Left            =   1620
      TabIndex        =   16
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label control5 
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Index           =   0
      Left            =   1620
      TabIndex        =   15
      Top             =   1740
      Width           =   2775
   End
   Begin VB.Label control6 
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Index           =   0
      Left            =   1620
      TabIndex        =   14
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label control16 
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Index           =   1
      Left            =   2100
      TabIndex        =   13
      Top             =   360
      Width           =   675
   End
   Begin VB.Label control17 
      Caption         =   "abs"
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   360
      Width           =   435
   End
   Begin VB.Label control16 
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Index           =   0
      Left            =   780
      TabIndex        =   11
      Top             =   360
      Width           =   735
   End
   Begin VB.Label control7 
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Index           =   0
      Left            =   780
      TabIndex        =   7
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label control18 
      Caption         =   "Type"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   1440
      Width           =   675
   End
   Begin VB.Label control19 
      Caption         =   "Size"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   1080
      Width           =   675
   End
   Begin VB.Label control20 
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Left            =   780
      TabIndex        =   2
      Top             =   60
      Width           =   3615
   End
   Begin VB.Label control21 
      Caption         =   "Offset"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   675
   End
   Begin VB.Label control22 
      Caption         =   "Scope"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   675
   End
End
Attribute VB_Name = "frm14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frm14.FRM
Option Explicit
Dim m001A As Integer
Dim m001C As Integer
Const mc0068 = 32 ' &H20%
Const mc0084 = 512 ' &H200%
Const mc0086 = 64 ' &H40%
Const mc0088 = 128 ' &H80%
Const mc008A = 16 ' &H10%
Const mc0098 = 15 ' &HF%
Const mc009A = 1 ' &H1%
Const mc009C = 2 ' &H2%
Const mc00A0 = 3 ' &H3%
Const mc00A2 = 4 ' &H4%
Const mc00A4 = 5 ' &H5%
Const mc00A6 = 6 ' &H6%
Const mc00A8 = 7 ' &H7%
Const mc00AA = 8 ' &H8%
Const mc00AC = 9 ' &H9%
Const mc00AE = 10 ' &HA%
Const mc00B0 = 11 ' &HB%
Const mc00B4 = 12 ' &HC%
Const mc00B6 = 13 ' &HD%
Const mc00C0 = 11 ' &HB%
Const mc00C2 = 1 ' &H1%
Const mc00C4 = 2 ' &H2%
Const mc00C6 = 3 ' &H3%
Const mc00C8 = 4 ' &H4%
Const mc00CA = 5 ' &H5%
Const mc00CC = 261 ' &H105%
Const mc00CE = 9 ' &H9%
Const mc00D0 = 10 ' &HA%

Sub control15_Change()
  sub2DC8
End Sub

Sub control9_SpinDown()
Dim l00D8 As Integer
Dim l00DA As Integer
  l00D8 = g_gbl(m001C).M4950
  If l00D8 <> m001A Then
    Do
      l00DA = gEXEdefEntries(l00D8).Count
      If l00DA = m001A Then
        control15 = l00D8
        Exit Sub
      End If
      l00D8 = l00DA
    Loop While l00D8
  End If
  ErrBeep2
End Sub

Sub control9_SpinUp()
Dim l00DC As Integer
  l00DC = gEXEdefEntries(m001A).Count
  If l00DC Then
    control15 = l00DC
  Else
    ErrBeep2
  End If
End Sub

Sub sub2DC8()
Dim l001E As Long
Dim l0020 As Integer
Dim l0022 As Integer
Dim l0024 As Integer
Dim l0026 As String
Dim l0028 As Integer
Dim l002A As Integer
' Const mc0068 = 32 ' &H20%
' Const mc0084 = 512 ' &H200%
' Const mc0086 = 64 ' &H40%
' Const mc0088 = 128 ' &H80%
' Const mc008A = 16 ' &H10%
' Const mc0098 = 15 ' &HF%
' Const mc009A = 1 ' &H1%
' Const mc009C = 2 ' &H2%
' Const mc00A0 = 3 ' &H3%
' Const mc00A2 = 4 ' &H4%
' Const mc00A4 = 5 ' &H5%
' Const mc00A6 = 6 ' &H6%
' Const mc00A8 = 7 ' &H7%
' Const mc00AA = 8 ' &H8%
' Const mc00AC = 9 ' &H9%
' Const mc00AE = 10 ' &HA%
' Const mc00B0 = 11 ' &HB%
' Const mc00B4 = 12 ' &HC%
' Const mc00B6 = 13 ' &HD%
Static m00BE As Integer
' Const mc00C0 = 11 ' &HB%
' Const mc00C2 = 1 ' &H1%
' Const mc00C4 = 2 ' &H2%
' Const mc00C6 = 3 ' &H3%
' Const mc00C8 = 4 ' &H4%
' Const mc00CA = 5 ' &H5%
' Const mc00CC = 261 ' &H105%
' Const mc00CE = 9 ' &H9%
' Const mc00D0 = 10 ' &HA%
  m001A = Val(control15)
  m001C = gEXEdefEntries(m001A).owningScope
  control20 = fn104E(m001C)
  l001E = gEXEdefEntries(m001A).M689A
  control16(0) = toHex16(l001E)
  control16(1) = toHex16(l001E + g_gbl(m001C).M49E2)
  control10 = fn1B19(m001A)
  l0024 = gEXEdefEntries(m001A).Size
  control8(0) = l0024 & "."
  control8(1) = "&h" & Hex(l0024)
  l002A = fn10AB(m001A)
  If l002A < l0024 Then
    control8(2).ForeColor = QBColor(12)
  Else
    control8(2).ForeColor = control8(1).ForeColor
  End If
  control8(2) = "&h" & Hex$(l002A)
  l002A = 0
  control13 = gEXEdefEntries(m001A).M6877
  l0020 = gEXEdefEntries(m001A).M6894.M36B3_count
  l0022 = gEXEdefEntries(m001A).M6894.M42D1
  l0028 = gEXEdefEntries(m001A).M6894.M42D9
  GoSub L57594
  If l0020 And mc0068 Then
    If l0022 Then
      l0020 = gEXEdef_Arr1(l0022).M6894.M36B3_count
      l0028 = gEXEdef_Arr1(l0022).M6894.M42D9
      l0024 = gEXEdef_Arr1(l0022).M3503
      l0022 = gEXEdef_Arr1(l0022).M6894.M42D1
      l002A = 1
      GoSub L57594
    End If
  End If
Exit Sub

L57594:
  control7(l002A) = toHex16(l0020)
  control3(l002A) = toHex16(l0022)
  control2(l002A) = toHex16(l0028)
  l0026 = ""
  If l0020 And mc0084 Then
    If l0020 And mc0068 Then
      l0026 = l0026 & "ParamArray "
    Else
      l0026 = " Optional "
    End If
  End If
  If l0020 And mc0086 Then
    l0026 = l0026 & "ByRef "
  End If
  If l0020 And mc0088 Then
    l0026 = l0026 & "ByVal "
  End If
  If l0020 And mc008A Then
    If l0028 = 0 Then
      l0026 = l0026 & "ptr to "
    End If
  End If
  If l0020 And mc0068 Then
    If l002A = 0 Then control1.Visible = True
    control5(l002A) = ""
    l0026 = l0026 & "array" & Str$(gEXEdef_Arr1(l0022).M34F5) & "d"
  Else
    If l002A = 0 Then control1.Visible = False
  If l0022 < 0 Then
    control5(l002A) = fn104E(-l0022) & "scope:" & Str$(-l0022)
  ElseIf l0022 > 0 Then
    control5(l002A) = Format$(l0022) & "."
  Else
    control5(l002A) = ""
  End If
    Select Case l0020 And mc0098
    Case 0: l0026 = l0026 & "*" & Format$(l0024)
    Case mc009A: l0026 = l0026 & "Byte"
      DbgVar l0024 = 1, "byte size"
    Case mc009C: l0026 = l0026 & "String *" & Str$(l0022)
If gIsVB_Ver Then
    DbgVar l0024 = l0022 * 2, "unicode"
Else
    DbgVar l0022 = l0024, "ansi"
End If
    Case mc00A0: l0026 = l0026 & "Boolean"
      DbgVar l0024 = 2, "bool size"
    Case mc00A2: l0026 = l0026 & "Integer"
    Case mc00A4: l0026 = l0026 & "Long"
    Case mc00A6: l0026 = l0026 & "Single"
    Case mc00A8: l0026 = l0026 & "Double"
    Case mc00AA: l0026 = l0026 & "Currency"
    Case mc00AC: l0026 = l0026 & "Date"
    Case mc00AE: l0026 = l0026 & "Variant"
    Case mc00B0: l0026 = l0026 & "Type "
      If l0022 < 0 Then
        l0026 = l0026 & fn104E(-l0022)
      ElseIf l0022 > 0 Then
        ErrBeep
        l0026 = l0026 & " " & fn1C33(l0022)
      Else
        l0026 = l0026 & " *" & toHex16(l0024)
      End If
    Case mc00B4: l0026 = l0026 & "String"
    Case mc00B6: l0026 = l0026 & "Object"
      If l0022 Then
        l0026 = l0026 & " " & fn104E(Abs(l0022))
      End If
    Case Else: l0026 = l0026 & "?" & toHex16(l0020)
    End Select
  End If
  control6(l002A) = l0026
  Select Case l0028 And &H11F
  Case m00BE, mc00C0: l0026 = "simple"
  Case mc00C2: l0026 = "var. string"
  Case mc00C4: l0026 = "variant"
  Case mc00C6: l0026 = "object"
  Case mc00C8: l0026 = "long fixed string"
  Case mc00CA: l0026 = "fixed array"
  Case mc00CC: l0026 = "dyn. array"
  Case mc00CE: l0026 = "Type"
  Case mc00D0: l0026 = "fixed string"
  Case Else: l0026 = "type " & toHex16(l0028) & " ???"
  End Select
  control4(l002A) = l0026
Return
End Sub

