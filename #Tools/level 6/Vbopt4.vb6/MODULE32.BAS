Attribute VB_Name = "MODULE32"
' Module32
Option Explicit
Const mc002A = -32768 ' &H8000%
Const mc003A = 11 ' &HB%
Const mc004A = 36 ' &H24%
Const mc0054 = 1 ' &H1%
Const mc005A = 18 ' &H12%
Const mc0064 = 12 ' &HC%
Const mc0066 = 58 ' &H3A%
Const mc009C = 2 ' &H2%
Const mc009E = 3 ' &H3%

Sub sub1F2B(p0006 As Integer, p0008%, p000A As Long)
Dim l000C As Integer
' Const mc002A =-32768 ' &H8000%
  Get g_hFileExe, fn1483(p0008) + p000A, gv2CCC
  Assume gv2CCC.M5A60 <> 0, "gbl layout"
  l000C = fn14AE(p0008, p000A, 0, True)
  Assume l000C, "gbl sym"
  g_decl_objs(p0006).M4131 = l000C
  sub1C42 l000C
If gIsVB5 Then
  DbgMileStone "special implement???"
  Exit Sub
  Get g_hFileExe, fn1483(l000C), gv20E4
  g_decl_objs(p0006).Size = gv20E4.M5EC8
Else
  Get g_hFileExe, fn1483(l000C), gv2DEC
  g_decl_objs(p0006).Size = gv2DEC.M5EC8
End If
  If gv2CCC.M5A68 Then
    l000C = fn14AE(p0008, p000A + 4, 0, True)
    g_decl_objs(p0006).M415B = l000C
    sub1C42 l000C
  End If
  If gv2CCC.M416F < 0 Or gv2CCC.M416F + gv2CCC.M66C4 >= &H8000& Then
    ShowMsgBox mc002A
  End If
  g_decl_objs(p0006).statSize = gv2CCC.statSize
  g_decl_objs(p0006).M416F = gv2CCC.M416F
  g_decl_objs(p0006).gbls_sorted = gv2CCC.M66C4
End Sub

Sub sub0D28()
Dim l002C
Dim l002E%
Dim l0030 As Integer
Dim l0032 As Integer
Dim l0034 As Integer
Dim l0036%
' Const mc003A = 11 ' &HB%
' Const mc004A = 36 ' &H24%
' Const mc0054 = 1 ' &H1%
' Const mc005A = 18 ' &H12%
' Const mc0064 = 12 ' &HC%
' Const mc0066 = 58 ' &H3A%
  If gEXEdef1 Then
    l0032 = gv2EF8(mc003A)
    l0034 = fn14AE(l0032, 0, l0036, True)
    For l002E = 1 To gEXEdef1
      l0034 = gv161C(l0036).subRef_M3C6A
      gReloc_Arr(l0034).M36B3_count = mc004A
      sub1F2B l002E, l0034, 0
      g_decl_objs(l002E).Name_M411A = "Module" & Format$(l002E) & ".BAS"
      g_decl_objs(l002E).M36B3_count = mc0054
      l0036 = fn149E(l0036)
    Next
  End If
  If gFormsCount Then
    l0030 = g_ResObjs(mc005A).M3595
    l0032 = gv2EF8(mc0064)
    l0036 = 0
    l0034 = fn14AE(l0032, 0, l0036, True)
    For l002C = 1 To gFormsCount
      l002E = gEXEdef1 + l002C
      l0034 = gv161C(l0036).subRef_M3C6A
      Assume l0034 > 0, "pat sym"
      gReloc_Arr(l0034).M36B3_count = mc0066
      sub1F2B l002E, l0034, 8
If gIsVB_Ver Then
      Get g_hFileExe, fn1483(l0034), gv33BE
      sub1F3F l002E, gv33BE.M41AD, gv33BE.M41B7, gv33BE.M41CD, l0030
Else
      Get g_hFileExe, fn1483(l0034), gv2D4E
      sub1F3F l002E, gv2D4E.M41AD, gv2D4E.M41B7, gv2D4E.M41CD, l0030
End If
      l0030 = gv026C(l0030).Count
      l0036 = fn149E(l0036)
    Next
  End If
End Sub

Sub sub1F3F(p007A As Integer, p007E As T4938, p0082 As T4938, p0086 As T4938, resource As Integer)
Dim l008A As Integer
Dim l008C As String
Dim l008E As Integer
Dim l0090 As String
Dim l0092 As String
' Const mc009C = 2 ' &H2%
' Const mc009E = 3 ' &H3%
  resource = g_ResObjs(mc005A).M3595
  Do While resource
    sub025C resource
    Get g_hFileExe, g_FileExeOffset + 4, gv0B84
    If gv0B84.M3CE1 = p0086.M3CE1 Then Exit Do
    resource = gv026C(resource).Count
  Loop
  
  Assume resource, "resource"
  l008C = fn040F(resource)
  DebugPrintVar Left$(l008C, 5) = "_IID_", "_IID_"
  l008C = Mid$(l008C, 6)
  l008A = InStr(l008C, ".")
  If l008A Then l008C = Left$(l008C, l008A - 1)
  If Mid$(p007E.M3CE1, 3) <> Mid$(p0082.M3CE1, 3) Then
    l008E = mc009C
    DbgMileStone "baseclass für CLS?"
    l0090 = "Control"
    l0092 = Left$(l008C, 8) & ".CLS"
    l008A = InStr(l008C, "CLASS")
    If l008A Then Mid$(l008C, l008A, 5) = "class"
  Else
    l008E = mc009E
    l0090 = "Form"
    l0092 = Left$(l008C, 8) & ".FRM"
    l008A = InStr(l008C, "FRM")
    If l008A Then Mid$(l008C, l008A, 3) = "frm"
  End If
  g_decl_objs(p007A).M3F23 = l008C
  g_decl_objs(p007A).Name_M411A = l0092
  g_decl_objs(p007A).M36B3_count = l008E
  g_decl_objs(p007A).M41AD = fn11BA(l008C, l0090, p007E)
  g_decl_objs(p007A).M41B7 = fn11BA(l008C, l0090, p0082)
  l008E = Asc(Mid$(p0082.M3CE1, 1, 1))
  Mid$(p0082.M3CE1, 1, 1) = Chr$(l008E + 1)
  g_decl_objs(p007A).M41C1 = fn11BA(l008C, l0090, p0082)
  g_decl_objs(p007A).M41CD = fn11BA(l008C, l0090, p0086)
End Sub
