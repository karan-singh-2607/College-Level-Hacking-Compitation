Attribute VB_Name = "MODULE29"
' Module29
Option Explicit
Const mc005E = 39 ' &H27%
Const mc0060 = 32 ' &H20%
Const mc0078 = 5 ' &H5%
Const mc007A = 9 ' &H9%
Const mc0086 = 3 ' &H3%
Const mc0088 = 15 ' &HF%
Const mc008A = 11 ' &HB%
Const mc008C = 20 ' &H14%
Const mc0090 = -16384 ' &HC000%
Const mc0092 = 16 ' &H10%
Const mc0096 = -20480 ' &HB000%
Const mc009A = 112 ' &H70%
Const mc00C4 = 1 ' &H1%
Const mc00C6 = 2 ' &H2%
Const mc00C8 = 4 ' &H4%
Const mc00CA = 3 ' &H3%
Const mc00CC = 192 ' &HC0%
Const mc0108 = 34 ' &H22%
Const mc0126 = -8192 ' &HE000%
Const mc0128 = 24576 ' &H6000%
Const mc012A = 16384 ' &H4000%
Const mc0132 = 8192 ' &H2000%
Const mc014E = 32 ' &H20%
Const mc0150 = 63 ' &H3F%
Const mc0152 = -8192 ' &HE000%
Const mc0154 = -24576 ' &HA000%
Const mc015E = -32768 ' &H8000%

Function fn1BD4(p0068%, p006C As T4B07, p006E As Integer, p0070%, p0072 As Long) As Integer
Dim l0074 As Integer
Dim l0076 As Integer
' Const mc0078 = 5 ' &H5%
' Const mc007A = 9 ' &H9%
' Const mc0086 = 3 ' &H3%
' Const mc0088 = 15 ' &HF%
' Const mc008A = 11 ' &HB%
' Const mc008C = 20 ' &H14%
' Const mc0090 =-16384 ' &HC000%
' Const mc0092 = 16 ' &H10%
' Const mc0096 =-20480 ' &HB000%
' Const mc009A = 112 ' &H70%
  Assume p0072 >= 0, "member offset >=0"
  l0076 = p006C.M6894.M42D1
  If l0076 < 0 Then
    l0076 = -l0076
    GoTo L3B0EC
  End If
  Select Case p006C.M6894.M42D9
  Case mc0078
    Exit Function
  Case mc007A
    DebugPrintVar ToUnSigned(g_gbl(l0076).Size) >= p0072, "within udt"
  Case mc0086
  End Select
  l0074 = p006C.M6894.M36B3_count
  If (l0074 And mc0088) <> mc008A Then
    If l0074 And mc0088 Then
      DbgMileStone "already typed non-UDT"
      gv34F4 = mc008C
      l0074 = l0074 And Not mc0088
      p006C.M6894.M42D1 = 0
    End If
    If p006E = mc0090 Then
      l0074 = l0074 Or mc0092
    End If
    p006C.M6894.M36B3_count = l0074 Or mc008A
    Assume p006E, "base scope"
    l0076 = fn10F6(p006E, p0070, 0)
    p006C.M6894.M42D1 = -l0076
    p006C.M6894.M42D9 = mc007A
  End If

L3B0EC:
  fn1BD4 = fn18EA(0, mc0096, p0072, l0076, 0, g_gbl(l0076).M4950)
  If p006C.M6894.M36B3_count And mc009A Then
    If p006C.Size <> 4 Then
      If p006C.Size = 0 Then
        p006C.Size = 4
      Else
        DbgMileStone "size mismatch"
      End If
    End If
  End If
  l0076 = p006C.Count
  If fn1C71(p0068) <> l0076 Then DbgMileStone "var removed!"
End Function

Function fn1C06(p00BA As Integer, ByVal pv00BC As Integer, p00C0 As T42C7) As Long
Dim l00C2 As Integer
' Const mc00C4 = 1 ' &H1%
' Const mc00C6 = 2 ' &H2%
' Const mc00C8 = 4 ' &H4%
' Const mc00CA = 3 ' &H3%
' Const mc00CC = 192 ' &HC0%
If gIsVB_Ver Then
  If p00C0.M36B3_count And mc0096 Then
    Select Case p00C0.M36B3_count And mc0088
    Case mc00C4
      l00C2 = 1
    Case mc00C6
      l00C2 = 2
      If pv00BC And 1 Then ErrBeep
    Case mc00C8, mc00CA
      l00C2 = 2
    Case mc008A
      l00C2 = gEXEdefEntries2(p00C0.M42D1).Size
    Case Else
      l00C2 = 4
    End Select
  ElseIf p00C0.M36B3_count And mc00CC Then
    l00C2 = 4
  Else
    l00C2 = fn102F(p00C0)
  End If
  If pv00BC < 0 Then
    p00BA = p00BA + pv00BC: pv00BC = 0
    If l00C2 >= 4 Then
      If p00BA And 3 Then
        p00BA = p00BA And -4
      End If
    ElseIf l00C2 >= 2 Then
      If p00BA And 1 Then
        p00BA = p00BA And -2
      End If
    End If
  Else
    If l00C2 >= 4 Then
      If p00BA And 3 Then
        p00BA = (p00BA + 3) And -4
      End If
    ElseIf l00C2 >= 2 Then
      If p00BA And 1 Then
        p00BA = (p00BA + 1) And -2
      End If
    End If
  End If
Else
  If pv00BC < 0 Then
    p00BA = p00BA + pv00BC: pv00BC = 0
  End If
End If
  fn1C06 = p00BA
  p00BA = p00BA + pv00BC
End Function

Function fn1C12(p00D2 As Long, p00D4 As Integer) As String
Dim l00D6 As Long
Dim l00D8 As Integer
Dim l00DA As Integer
  If p00D4 Then
    l00DA = g_gbl(p00D4).M4950
    Do While l00DA
      If p00D2 <= gEXEdefEntries(l00DA).M689A Then
        If p00D2 = gEXEdefEntries(l00DA).M689A Then
          fn1C12 = fn1C21(l00DA)
          Exit Function
        End If
        Exit Do
      End If
      l00D8 = l00DA
      l00DA = gEXEdefEntries(l00D8).Count
    Loop
    If l00D8 Then
      If gEXEdefEntries(l00D8).M6894.M42D9 = 9 Then
        l00D6 = p00D2 - gEXEdefEntries(l00D8).M689A
        fn1C12 = "." & fn1B19(l00D8) & fn1C12(l00D6, gEXEdefEntries(l00D8).M6894.M42D1)
        Exit Function
      End If
    Else
      DbgMileStone "add new member?"
    End If
  Else
  End If
  fn1C12 = ".M" & toHex16(p00D2)
End Function

Function fn1C21(p00E8%) As String
Dim l00EA As String
  l00EA = fn1B19(p00E8)
  If Left$(l00EA, 1) = "." Then
    fn1C21 = l00EA
  Else
    fn1C21 = "." & l00EA
  End If
End Function

Function fn1C33(p00EE%) As String
Dim l00F0 As Integer
  l00F0 = fn1CA9(p00EE)
  If l00F0 Then
    fn1C33 = gv38BA(l00F0).Name
    Exit Function
  End If
  l00F0 = gEXEdefEntries2(p00EE).subRef_M3C6A
  If l00F0 Then
    fn1C33 = fn2023(l00F0)
  Else
    DbgMileStone "exetyp ohne sym"
    fn1C33 = "T" & toHex16(p00EE)
  End If
End Function

Function fn1C51(p010C() As Integer, p0112 As Integer, p0114 As Integer) As Integer
Dim l0116 As Long
Dim l0118 As Integer
Dim l011A As Long
  fn1C51 = gEXEdefEntries(p0112).Count
  Assume gEXEdefEntries(p0112).owningScope = p0114, "owning scope"
  
  l0118 = gEXEdefEntries(p0112).Size
  If l0118 <= 0 Then
    If gEXEdefEntries(p0112).M6894.M36B3_count And mc009A Then
      DbgMileStone "ptr without size"
      l0118 = 4
      gEXEdefEntries(p0112).Size = l0118
    Else
      l0118 = 1
    End If
  End If
  l0116 = gEXEdefEntries(p0112).M689A
  If l0116 + ToUnSigned(l0118) > ToUnSigned(g_gbl(p0114).Size) Then
    Exit Function
  End If
  Assume l0116 + ToUnSigned(l0118) <= ToUnSigned(g_gbl(p0114).Size), "within udt"
  l0116 = l0116 + g_gbl(p0114).M49E2
  For l011A = l0116 To l0116 + l0118 - 1
    l0118 = p010C(l011A)
    If l0118 Then
      If l0118 <> p0112 Then
        If l0118 < 0 Then
          DbgMileStone "check remove scope ref"
          p010C(l011A) = p0112
        ElseIf gEXEdefEntries(l0118).owningScope <> p0114 Then
          p010C(l011A) = p0112
        Else
          If gEXEdefEntries(l0118).M689A > gEXEdefEntries(p0112).M689A Then
            sub1C62 l0118, p0112
            p010C(l011A) = p0112
          Else
            Assume l011A = l0116, "first ref"
            sub1C62 p0112, l0118
            Exit Function
          End If
        End If
      End If
    Else
      p010C(l011A) = p0112
    End If
  Next
End Function

Function fn1C71(p011E As Integer) As Integer
Dim l0120 As Integer
Dim l0124 As Integer
' Const mc0126 =-8192 ' &HE000%
' Const mc0128 = 24576 ' &H6000%
' Const mc012A = 16384 ' &H4000%
' Const mc0132 = 8192 ' &H2000%
Static m0134 As Integer
  fn1C71 = gEXEdefEntries(p011E).Count
  l0124 = gEXEdefEntries(p011E).Size
  Assume l0124 >= 0, "var size>0"
  If l0124 = 0 Then
    If gEXEdefEntries(p011E).M6894.M36B3_count And mc009A Then
      l0124 = 4
      gEXEdefEntries(p011E).Size = l0124
    End If
  End If
  l0120 = gEXEdefEntries(p011E).owningScope
  Select Case g_gbl(l0120).M36B3_count And mc0126
  Case mc0128, mc012A
    fn1C71 = fn1C51(gv30F0(), p011E, l0120)
  Case mc0132, m0134
    fn1C71 = fn1C51(gv3126(), p011E, l0120)
  End Select
End Function

Function fn1C80(p015A As Integer) As Integer
Dim l015C As Integer
' Const mc015E =-32768 ' &H8000%
  DebugPrintVar g_gbl(p015A).M36B3_count And mc015E, "guid scope"
  l015C = g_gbl(p015A).M36B3_count And &HFFF
  fn1C80 = l015C
End Function

Function fn1C8E(p0162 As Integer) As Integer
Dim l0164 As Integer
Dim l0166 As Integer
Dim l0168 As Integer
  l0164 = g_gbl(p0162).M36B3_count
  l0166 = l0164 And &HFFF
  l0168 = fn1CA9(l0166)
  fn1C8E = l0168
End Function

Function fn1C9B(Entry As Integer, Bool As Integer) As Integer
Dim scopeEntries As Integer
  
  If Bool Then
    
    scopeEntries = gEXEdefEntries2(Entry).M56B4
    If scopeEntries = 0 Then
      
      DebugPrintVar gEXEdefEntries2(Entry).Size > 0, "scopeEntries size"
      scopeEntries = fn10F6(mc0152 Or &H1000, Entry, gEXEdefEntries2(Entry).subRef_M3C6A)
      
      gEXEdefEntries2(Entry).M56B4 = scopeEntries
    End If
  
  Else
    
    scopeEntries = gEXEdefEntries2(Entry).M56A4
    If scopeEntries = 0 Then
    
    'New
      scopeEntries = fn10F6(mc0152, Entry, gEXEdefEntries2(Entry).subRef_M3C6A)
      
      gEXEdefEntries2(Entry).M56A4 = scopeEntries
    End If
  End If
  
  Assume scopeEntries <= gEXEdef6_Num_scopes, "valid scopeEntries"
  
  fn1C9B = scopeEntries
End Function

Function fn1CA9(p0176 As Integer) As Integer
  fn1CA9 = gEXEdefEntries2(p0176).M56AD
End Function

Sub sub0C60(p005A As Integer)
Dim l005C As Long
' Const mc005E = 39 ' &H27%
' Const mc0060 = 32 ' &H20%
  Select Case gReloc_Arr(p005A).M36B3_count
  Case mc005E
    If gReloc_Arr(p005A).subRef_M3C6A Then
      ShowMsgBox mc0060
      Exit Sub
    End If
  End Select
  g_userTypesCount = g_userTypesCount + 1
  If g_userTypesCount > UBound(gEXEdefEntries2) Then ReDim Preserve gEXEdefEntries2(g_userTypesCount + 10)
  gEXEdefEntries2(g_userTypesCount).subRef_M3C6A = p005A
  l005C = Seek(g_hFileExe)
  Get g_hFileExe, fn1483(p005A) + 2, gEXEdefEntries2(g_userTypesCount).Size
  Seek g_hFileExe, l005C
  gReloc_Arr(p005A).M36B3_count = mc005E
  Assume MakeSubRef(p005A, "Type" & toHex16(g_userTypesCount), -g_userTypesCount), "sym created"
End Sub

Sub sub0C6E(p003C As Integer, p003E As Integer)
Dim l0040 As Integer
Dim l0044 As T66F3
  Assume p003E > 0, "off>0"
If gIsVB_Ver Then
  If p003E And 2 Then p003E = p003E + 2
End If
  l0040 = gReloc_Arr(p003C).Count
  If fn202F(l0040) = fn202F(p003C) + p003E Then Exit Sub
  l0040 = CreateSym(fn202F(p003C) + p003E)
If gIsVB5 Then
Else
  Get g_hFileExe, fn1483(l0040), l0044
  If l0044.M5EBD = gv3754 And l0044.Entries = 0 And l0044.Flags = &H4000 Then
    sub0C60 l0040
  End If
End If
End Sub

Sub sub0CDD()
  g_userTypesCount = 0
  ReDim gEXEdefEntries2(10)
  ReDim gEXEdef_Arr1(10): gEXEdef_Arr1Count = 0
  ReDim gEXEdef_Arr2(0): gEXEdef_Arr2Count = 0
If gIsVB_Ver = 4 Then
  gv3754 = &HC
ElseIf gIsVB_Ver = 5 Then
  gv3754 = Len(gv20E4)
Else
  gv3754 = Len(gv2DEC)
End If
End Sub

Sub sub0D1C()
Dim l009E As T66F3
Dim l00A0 As Integer
If gIsVB5 Then
  DbgMileStone "assume unused!"
  Exit Sub
End If
  DbgVar g_userTypesCount = 0, "no types"
  l00A0 = gv2EF8(3)
  Do While l00A0
    Get g_hFileExe, fn1483(l00A0), l009E
    If l009E.M5EBD <> gv3754 Or l009E.Entries Then
      If fn202F(gReloc_Arr(l00A0).Count) <> fn202F(l00A0) + gv3754 Then
        Exit Do
      End If
      l00A0 = gReloc_Arr(l00A0).Count
    Else
      Select Case l009E.Flags
      Case &H4000
        sub0C60 l00A0
      Case 0
      Case Else
        Exit Do
      End Select
      l00A0 = CreateSym(fn202F(l00A0) + gv3754)
    End If
  Loop
End Sub

Sub sub1BE2(p00A8%, p00AC As Control)
  sub0FC7 fn1C9B(p00A8, False), p00AC, True
End Sub

Sub FillCombo_userTypes(ComboBox As ComboBox)
Dim i%
Dim l00B6 As Integer
  ComboBox.Clear
  ComboBox.ListIndex = -1
  For i = 1 To g_userTypesCount
    l00B6 = fn1C9B(i, False)
    ComboBox.AddItem toHex16(g_gbl(l00B6).Size) & " " & fn1C33(i)
    ComboBox.ItemData(ComboBox.NewIndex) = i
  Next
End Sub

Sub sub1C42(p0104 As Integer)
Dim l0106%
' Const mc0108 = 34 ' &H22%
  gReloc_Arr(p0104).M36B3_count = mc0108
  Get g_hFileExe, fn1483(p0104), l0106
If gIsVB_Ver Then
  If l0106 And 2 Then
    DbgMileStone "really align???"
    l0106 = l0106 + 2
  End If
End If
  sub0C6E p0104, l0106
End Sub

Sub sub1C62(p013C As Integer, p013E As Integer)
Dim l0140 As Integer
Dim l0142 As Integer
Dim l0144 As Integer
Dim l0146 As Integer
Dim l0148 As Long
Dim l014C As Integer
' Const mc014E = 32 ' &H20%
' Const mc0150 = 63 ' &H3F%
' Const mc0152 =-8192 ' &HE000%
' Const mc0154 =-24576 ' &HA000%
  Assume gEXEdefEntries(p013C).owningScope = gEXEdefEntries(p013E).owningScope, "scope match"
  l0140 = gEXEdefEntries(p013C).owningScope
  If gEXEdefEntries(p013E).M6894.M36B3_count And mc014E Then
    l014C = gEXEdefEntries(p013E).M6894.M42D1
    Assume l014C > 0, "dst scope"
    DebugPrintVar gEXEdef_Arr1(l014C).M6894.M42D9 = mc007A, "udt scope"
    l014C = -gEXEdef_Arr1(l014C).M6894.M42D1
  Else
    DebugPrintVar gEXEdefEntries(p013E).M6894.M42D9 = mc007A, "udt scope"
    l014C = -gEXEdefEntries(p013E).M6894.M42D1
    If l014C <= 0 Then Exit Sub
  End If
  DbgVar l014C > 0, "scoped"
  If l014C <= 0 Then Exit Sub
  DbgVar g_gbl(l014C).M36B3_count And mc0126, "typed scope"
  l0142 = g_gbl(l0140).M4950
  l0144 = gEXEdefEntries(l0142).Count
  If p013C = l0142 Then
    g_gbl(l0140).M4950 = l0144
  Else
    Do While l0144 <> p013C
      l0142 = l0144
      l0144 = gEXEdefEntries(l0142).Count
      If l0144 = 0 Then Exit Do
    Loop
    If l0144 Then
      gEXEdefEntries(l0142).Count = gEXEdefEntries(l0144).Count
    Else
      DbgMileStone "var not linked!?"
    End If
  End If
  l0148 = gEXEdefEntries(p013C).M689A - gEXEdefEntries(p013E).M689A
  DebugPrintVar l0148 > 0, "src off>0"
  gEXEdefEntries(p013C).M689A = l0148
  gEXEdefEntries(p013C).M68A4 = l0148
  gEXEdefEntries(p013C).owningScope = l014C
  gEXEdefEntries(p013C).M6877 = 0
  l0146 = gEXEdefEntries(p013C).M6894.M36B3_count And mc0150
  l0146 = l0146 Or mc0096
  gEXEdefEntries(p013C).M6894.M36B3_count = l0146
  l0146 = g_gbl(l014C).M36B3_count
  Select Case l0146 And mc0126
  Case mc0152, mc0090, mc0154
    l0142 = 0
    l0144 = g_gbl(l014C).M4950
    If l0144 = 0 Then
      DbgMileStone "enter first"
    Else
      If gEXEdefEntries(l0144).M689A > l0148 Then
      Else
        Do
          l0144 = gEXEdefEntries(l0142).Count
          If l0144 = 0 Then Exit Do
          If gEXEdefEntries(l0144).M689A <= l0148 Then Exit Do
          l0142 = l0144
        Loop
      End If
    End If
    If l0142 Then
      If gEXEdefEntries(l0142).M689A + ToUnSigned(gEXEdefEntries(l0142).Size) <= l0148 Then
        gEXEdefEntries(l0142).Count = p013C
      Else
        ShowMsgBox mc0060
      End If
    Else
      g_gbl(l014C).M4950 = p013C
    End If
    If l0144 = 0 Then
      gEXEdefEntries(p013C).Count = l0144
    ElseIf gEXEdefEntries(l0144).M689A = l0148 Then
      sub10D6 gEXEdefEntries(p013C), gEXEdefEntries(l0144)
      gEXEdefEntries(p013C).Count = gEXEdefEntriesCount2
      gEXEdefEntriesCount2 = p013C
      Exit Sub
    Else
      DbgVar gEXEdefEntries(l0144).M689A >= l0148 + ToUnSigned(gEXEdefEntries(p013C).Size), "ref within src"
      gEXEdefEntries(p013C).Count = l0144
    End If
  End Select
End Sub

