Attribute VB_Name = "MODULE15"
' Module15
Option Explicit
Const mc008E = 112 ' &H70%
Const mc0090 = 15 ' &HF%
Const mc0092 = 1 ' &H1%
Const mc0094 = 4 ' &H4%
Const mc0096 = 3 ' &H3%
Const mc0098 = 5 ' &H5%
Const mc009A = 6 ' &H6%
Const mc009C = 7 ' &H7%
Const mc009E = 8 ' &H8%
Const mc00A0 = 9 ' &H9%
Const mc00A2 = 10 ' &HA%
Const mc00A4 = 11 ' &HB%
Const mc00A6 = 2 ' &H2%
Const mc00CA = 64 ' &H40%
Const mc00FE = 11 ' &HB%
Const mc0100 = 63 ' &H3F%
Const mc0102 = 257 ' &H101%
Const mc0104 = 514 ' &H202%
Const mc0106 = 258 ' &H102%
Const mc0108 = 260 ' &H104%
Const mc010A = 516 ' &H204%
Const mc010C = 264 ' &H108%
Const mc010E = 1032 ' &H408%
Const mc0110 = 520 ' &H208%
Const mc0112 = 272 ' &H110%
Const mc0114 = 28 ' &H1C%
Const mc0116 = 4356 ' &H1104%
Const mc0118 = 29 ' &H1D%
Const mc011A = 4612 ' &H1204%
Const mc011C = 5124 ' &H1404%
Const mc011E = 1 ' &H1%
Const mc0120 = 2 ' &H2%
Const mc0122 = 4 ' &H4%
Const mc0124 = 8 ' &H8%
Const mc0126 = 16 ' &H10%
Const mc0128 = 32 ' &H20%
Const mc012C = 4096 ' &H1000%
Const mc017C = 1 ' &H1%
Const mc017E = 2 ' &H2%
Const mc0180 = 10 ' &HA%
Const mc0182 = 4 ' &H4%
Const mc0184 = 3 ' &H3%
Const mc0186 = 5 ' &H5%
Const mc0188 = 9 ' &H9%
Const mc018A = 32 ' &H20%
Const mc0224 = 1 ' &H1%
Const mc022E = -32768 ' &H8000%
Const mc023E = 14 ' &HE%
Const mc0246 = 13 ' &HD%
Const mc0248 = 38 ' &H26%

Function fn0F99(p01E8 As Integer, p01EA As Long) As Integer
Dim l01EC As Integer
Dim l01EE As Integer
Dim l01F2 As Integer
On Error Resume Next
  l01EE = g_gbl(p01E8).M36B3_count
  If l01EE < 0 Then
    l01F2 = g_gbl(p01E8).M4950
    Do While gEXEdefEntries(l01F2).M689A <= p01EA
      l01EC = l01F2
      l01F2 = gEXEdefEntries(l01F2).Count
      If l01F2 = 0 Then Exit Do
    Loop
  Else
    p01EA = p01EA + g_gbl(p01E8).M49E2
    Select Case l01EE And gc3028
    Case gc3018
      l01EC = gv30F0(p01EA)
    Case gc301A
      l01EC = gv30F0(p01EA)
    Case gc301C, gv301E
      l01EC = gv3126(p01EA)
    End Select
  End If
  fn0F99 = l01EC
End Function

Function fn0FD9(ByVal p0232 As Long) As String
  gv2FDE = gv30F0(p0232)
  If gv2FDE < 0 Then
    If gv2FDE = &H8000 Then
      fn0FD9 = "GUID" & toHex16(p0232)
    Else
      fn0FD9 = fn104E(-gv2FDE)
    End If
  ElseIf gv2FDC Then
    fn0FD9 = fn1B34(p0232)
  ElseIf p0232 < &H10000 Then
    fn0FD9 = "g" & toHex16(p0232)
  Else
    fn0FD9 = "g" & Hex$(p0232)
  End If
End Function

Function fn100C(p0256 As Integer, p0258 As Integer) As String
On Error GoTo 0
  If p0256 = gv2FDA Then
    fn100C = "me"
    Exit Function
  End If
  If gv2FDC = p0258 Then
    fn100C = fn1B28(p0256)
  ElseIf p0256 > 0 Then
    fn100C = "p" & toHex16(CLng(p0256))
  Else
    fn100C = "l" & toHex16(CLng(p0256))
  End If
End Function

Function fn102F(p008A As T42C7) As Integer
Dim l008C As Integer
' Const mc008E = 112 ' &H70%
' Const mc0090 = 15 ' &HF%
' Const mc0092 = 1 ' &H1%
' Const mc0094 = 4 ' &H4%
' Const mc0096 = 3 ' &H3%
' Const mc0098 = 5 ' &H5%
' Const mc009A = 6 ' &H6%
' Const mc009C = 7 ' &H7%
' Const mc009E = 8 ' &H8%
' Const mc00A0 = 9 ' &H9%
' Const mc00A2 = 10 ' &HA%
' Const mc00A4 = 11 ' &HB%
' Const mc00A6 = 2 ' &H2%
  If p008A.M36B3_count And mc008E Then
    l008C = 4
  Else
    Select Case p008A.M36B3_count And mc0090
    Case mc0092: l008C = 1
    Case mc0094, mc0096: l008C = 2
    Case mc0098, mc009A: l008C = 4
    Case mc009C, mc009E, mc00A0: l008C = 8
    Case mc00A2: l008C = 16
    Case mc00A4: l008C = g_gbl(p008A.M42D1).Size
      Assume l008C > 0, "UDT<32K"
    Case mc00A6
If gIsVB_Ver Then
      l008C = p008A.M42D1 * 2
Else
      l008C = p008A.M42D1
End If
    Case Else
      ErrBeep
    End Select
  End If
  fn102F = l008C
End Function

Function fn103E(p00AC As Integer, ByVal p00AE As Long, p00B0%) As String
Dim l00B2 As Integer
Dim l00B4 As Integer
Dim l00B6 As String
' Const mc00CA = 64 ' &H40%
On Error GoTo 0
    If p00AC > g_gbl(gL2Min).Size Then
      Assume p00AE = 0, "locals<64K"
      fn103E = fn0FD9(p00AC) & l00B6
      Exit Function
    End If
  If p00AC = gv2FDA Then
If gIsVB5 Then
    fn103E = fn1B34(g_decl_objs(gv16BA).M4150 + p00AE)
    Exit Function
Else
    If g_decl_objs(gv16BA).statSize Then
      If p00AE = g_decl_objs(gv16BA).M416F Then
        fn103E = "Static"
        gv2FDE = -(gL1Max + gv16BA)
        Exit Function
      End If
    End If
    If gv16BA > gEXEdef1 Then
      If p00AE = 0 Then fn103E = "me": Exit Function
    End If
      l00B6 = fn1B34(g_decl_objs(gv16BA).M4126 + p00AE)
      fn103E = l00B6
      Exit Function
End If
  End If
  l00B2 = gv3126(p00AC)
  If l00B2 = 0 Then
    If p00AC > g_gbl(gL2Min).Size Then
      Assume p00AE = 0, "locals<64K"
      fn103E = fn0FD9(p00AC) & l00B6
      Exit Function
    Else
      l00B2 = fn18EA(0, mc00CA, p00AC, gL2Min, gv16BC, g_gbl(gL2Min).M4950)
      If (gEXEdefEntries(l00B2).M6894.M36B3_count And mc0090) <> mc00A4 Then
        gEXEdefEntries(l00B2).M6894.M36B3_count = mc00CA Or mc00A4
      End If
    End If
  ElseIf (gEXEdefEntries(l00B2).M6894.M36B3_count And mc0090) <> mc00A4 Then
    gEXEdefEntries(l00B2).M6894.M36B3_count = mc00CA Or mc00A4
  End If
  l00B4 = fn1175(l00B2)
  If l00B4 > 0 Then
    l00B6 = fn1C12(p00AE, l00B4)
  ElseIf p00AE Then
    l00B6 = " +" & toHex16(p00AE)
  End If
  fn103E = fn100C(l00B2, p00B0) & l00B6
End Function

Function fn104E(p00D6 As Integer) As String
Dim l00D8 As Integer
  l00D8 = g_gbl(p00D6).M49D8
  If l00D8 Then
    fn104E = fn2023(l00D8)
  Else
    fn104E = "scope#" & Format$(p00D6)
  End If
End Function

Function fn105F(ByVal pv00DE As Integer) As String
  If gv2FDC Then
    fn105F = fn1182(gv2FDE, pv00DE)
  Else
    fn105F = "ptr+" & toHex16(CLng(pv00DE))
  End If
End Function

Function fn106F(p00E2 As Integer, p00E4 As Long, p00E6 As Long) As Integer
Dim l00E8 As Integer
Dim l00EA As Integer
Dim l00EC As Integer
Dim l00EE As Integer
Dim l00F0 As Integer
Dim l00F2 As Integer
Dim l00F4 As Integer
' Const mc00FE = 11 ' &HB%
' Const mc0100 = 63 ' &H3F%
' Const mc0102 = 257 ' &H101%
' Const mc0104 = 514 ' &H202%
' Const mc0106 = 258 ' &H102%
' Const mc0108 = 260 ' &H104%
' Const mc010A = 516 ' &H204%
' Const mc010C = 264 ' &H108%
' Const mc010E = 1032 ' &H408%
' Const mc0110 = 520 ' &H208%
' Const mc0112 = 272 ' &H110%
' Const mc0114 = 28 ' &H1C%
' Const mc0116 = 4356 ' &H1104%
' Const mc0118 = 29 ' &H1D%
' Const mc011A = 4612 ' &H1204%
' Const mc011C = 5124 ' &H1404%
' Const mc011E = 1 ' &H1%
' Const mc0120 = 2 ' &H2%
' Const mc0122 = 4 ' &H4%
' Const mc0124 = 8 ' &H8%
' Const mc0126 = 16 ' &H10%
' Const mc0128 = 32 ' &H20%
  l00E8 = gEXEdef_Subs(p00E2).M36B3_count
  l00EC = gEXEdef_Subs(p00E2).M4277
  l00EA = l00EC And &H1F
  If l00EA = 1 Then l00EA = 2
  If l00EC Then
    If l00E8 = mc00FE Then
      l00F2 = g_gbl(gL3Max + p00E2).M4950
      If l00F2 Then
        If gEXEdefEntries(l00F2).M689A + l00EA = gv2FD8 Then
          l00EE = gEXEdefEntries(l00F2).M6894.M36B3_count And mc0100
          Select Case l00EE
          Case mc0092: l00F4 = mc0102
          Case mc0096: l00F4 = mc0104
          Case mc0094: l00F4 = mc0106
          Case mc0098: l00F4 = mc0108
          Case mc009A: l00F4 = mc010A
          Case mc009C: l00F4 = mc010C
          Case mc00A0: l00F4 = mc010E
          Case mc009E: l00F4 = mc0110
          Case mc00A2: l00F4 = mc0112
          Case mc0114: l00F4 = mc0116
          Case mc0118: l00F4 = mc011A
          Case mc00A4: l00F4 = mc011C
          Case 0
          End Select
          If l00F4 Then
            If l00F4 <> l00EC Then
              l00EC = l00F4
              gEXEdef_Subs(p00E2).M4277 = l00F4
            End If
            l00EA = l00EC And &H1F
            If l00EA = 1 Then l00EA = 2
          End If
        End If
      End If
    End If
  End If
  If l00F4 = 0 Then
    Select Case l00EC
    Case 0: Exit Function
    Case mc0102: l00EE = mc0092
    Case mc0104: l00EE = mc0096
    Case mc0106: l00EE = mc0094
    Case mc0116: l00EE = mc0114
    Case mc011A: l00EE = mc0118
    Case mc011C: l00EE = mc00A4
    Case mc010A: l00EE = mc009A
    Case mc0108: l00EE = mc0098
    Case mc0110: l00EE = mc009E
    Case mc010E: l00EE = mc00A0
    Case mc010C: l00EE = mc009C
    Case mc0112: l00EE = mc00A2
    Case Else
      l00F0 = True
      Select Case l00EA
      Case mc011E: l00EE = mc0092
      Case mc0120: l00EE = mc0094
      Case mc0122: l00EE = mc0098
      Case mc0124: l00EE = mc009C
      Case mc0126: l00EE = mc00A2
      Case mc0128: l00EE = mc00A4: l00EA = 0
      Case Else
        Exit Function
      End Select
    End Select
  End If
  If l00E8 = mc00FE Then
    If l00EA Then
      p00E6 = gv2FD8 - l00EA
    End If
    Select Case l00EE
    Case 0
    Case 0 To mc0098, mc0114
    Case mc0118
      p00E4 = gv2FDA
    Case mc00A2
      If gIsVB_Ver Then
        p00E4 = gv2FDA + 4
      Else
        p00E4 = gv2FDA + 2
      End If
    Case Else
      p00E4 = gv2FDA + 2
    End Select
  Else
    If l00EC Then
      p00E4 = gv315C - 4
    End If
  End If
  fn106F = l00EE
End Function

Function fn108E(p0130 As Integer, p0132 As Integer) As Integer
  If p0130 <> p0132 Then
    If p0132 = 0 Then
      p0132 = p0130
    ElseIf p0130 Then
      MsgBox "incompatible info"
      fn108E = -1
    End If
  End If
End Function

Function fn10AB(p0156 As Integer) As Integer
Dim l0158 As Integer
Dim l015A As Integer
Dim l015C As Integer
Dim l015E As Long
Dim l0160 As Long
  l0158 = gEXEdefEntries(p0156).M6894.M36B3_count
  l0160 = gEXEdefEntries(p0156).M689A
  l015A = gEXEdefEntries(p0156).owningScope
  If g_gbl(l015A).M36B3_count And gc302A Then
    l015C = g_gbl(l015A).M4950
    If l015C = p0156 Then
      If l0160 < 0 Then
        l015E = gv2FD8 - l0160
      Else
        l015E = ToUnSigned(g_gbl(l015A).Size) - l0160
      End If
    Else
      Do While gEXEdefEntries(l015C).Count <> p0156
        If gEXEdefEntries(l015C).M689A = l0160 Then
          If l015C <> p0156 Then
            Exit Do
          End If
        End If
        l015C = gEXEdefEntries(l015C).Count
      Loop
      l015E = gEXEdefEntries(l015C).M689A - l0160
    End If
  Else
    l015C = gEXEdefEntries(p0156).Count
    If l015C Then
      l015E = gEXEdefEntries(l015C).M689A - l0160
    Else
      l015E = ToUnSigned(g_gbl(l015A).Size) - l0160
    End If
  End If
  DbgVar l015E > 0, "var size"
  fn10AB = l015E
End Function

Function fn10F6(p018E As Integer, p0190 As Integer, p0192 As Integer) As Integer
Dim l0194 As Integer
  l0194 = gEXEdef6_Num_scopes + 1
  gEXEdef6_Num_scopes = l0194
  If l0194 > UBound(g_gbl) Then ReDim Preserve g_gbl(l0194 + 10)
  g_gbl(l0194).M36B3_count = p018E Or p0190
  g_gbl(l0194).M49D8 = p0192
  fn10F6 = l0194
  gIs3124 = True
  Select Case p018E And gc3028
  Case gc3022, gc3024
  Case gc3026: gIs3158 = True
  End Select
End Function

Function fn1103(p0198 As Integer) As Integer
Dim l019C As T4B07
Dim l019E As Integer
  l019E = gEXEdefEntriesCount2
  If l019E Then
    gEXEdefEntriesCount2 = gEXEdefEntries(l019E).Count
    gEXEdefEntries(l019E) = l019C
  Else
    l019E = gEXEdefEntriesCount1 + 1
    gEXEdefEntriesCount1 = l019E
    If gEXEdefEntriesCount1 > UBound(gEXEdefEntries) Then ReDim Preserve gEXEdefEntries(gEXEdefEntriesCount1 + 100)
  End If
  gEXEdefEntries(l019E).Count = p0198
  p0198 = l019E
  fn1103 = l019E
End Function

Function fn111E(p01A6 As Integer, p01A8&) As String
Dim l01AA As String
Dim l01AE As Integer
Dim l01B2 As Long
  l01AE = g_gbl(p01A6).M36B3_count
  If l01AE < 0 Then
    l01AA = fn1C12(p01A8, p01A6)
  Else
    Select Case l01AE And gc3028
    Case gc3018
      l01B2 = p01A8 + g_gbl(p01A6).M49E2
      l01AA = fn1B34(l01B2)
    Case gc301A
      l01AA = fn1B41(p01A8)
    Case gc301C, gv301E
      l01AA = fn1B28(p01A8)
    End Select
  End If
  fn111E = l01AA
End Function

Function fn1175(p01D8 As Integer) As Integer
Dim l01DA As Integer
  If p01D8 > 0 Then
    l01DA = gEXEdefEntries(p01D8).M6894.M42D1
  Else
    l01DA = p01D8
  End If
  DebugPrintVar l01DA, "scoped var"
  If l01DA < 0 Then
    fn1175 = -l01DA
  End If
End Function

Function fn1182(p01DE As Integer, ByVal p01E0&) As String
Dim l01E2 As String
Dim l01E4 As Integer
  If p01DE Then
    If gv2FDC Then
      If p01DE < 0 Then
        If p01DE = &H8000 Then
          DbgMileStone "find GUID ref'd"
          fn1182 = "GUID+" & toHex16(p01E0)
          Exit Function
        End If
        l01E4 = -p01DE
      ElseIf p01DE Then
        l01E4 = fn1175(p01DE)
      End If
      If l01E4 > 0 Then
        l01E2 = fn111E(l01E4, p01E0)
      End If
      If Len(l01E2) Then
        fn1182 = l01E2
        Exit Function
      End If
    End If
  Else
    DbgMileStone "missing basevar"
  End If
  fn1182 = "p+" & toHex16(p01E0)
End Function

Sub sub0CCE()
If gIsVB_Ver Then
  gv2FDA = 8
  gv2FD8 = &HFF7C
Else
  gv2FDA = 6
  gv2FD8 = &HFFB0
End If
End Sub

Sub sub0D39()
Dim l0144%
  If gv097C = 0 Then Exit Sub
  If g_decl_num = gv097C And gFormsCount = gv0982 Then
    For l0144 = 1 To g_decl_num
      g_decl_objs(l0144).Name_M411A = fn0B9D(l0144)
    Next
  End If
End Sub

Sub sub0DB8()
Dim l0218
Dim l021A As Long
Dim l0220 As Integer
' Const mc0224 = 1 ' &H1%
' Const mc022E =-32768 ' &H8000%
On Error GoTo L21A9E
  If g_decl_num = 0 Then ErrBeep: Exit Sub
If gIsVB_Ver = 5 Then
  gblsCountVars = 0
  For l0218 = 1 To g_decl_num
    If g_decl_objs(l0218).M36B3_count = mc0224 Then
      Assume g_decl_objs(l0218).M4126 > 0, "mod always Gbl"
      l021A = g_decl_objs(l0218).M4126 + g_decl_objs(l0218).Size
      If g_decl_objs(l0218).statSize Then
        Assume (g_decl_objs(l0218).M4178 > 0), "pStat"
        If g_decl_objs(l0218).M4178 > l021A Then
          l021A = g_decl_objs(l0218).M4178 + g_decl_objs(l0218).statSize
        End If
      Else
        Assume g_decl_objs(l0218).M4178 = 0, "no statics"
      End If
      If l021A And 2 Then l021A = l021A + 2
    Else
      l021A = g_decl_objs(l0218).M4126 + &H14
    End If
    If l021A > gblsCountVars Then gblsCountVars = l021A
  Next
  For l0218 = 1 To g_decl_num
    If g_decl_objs(l0218).M36B3_count = mc0224 Then
    Else
      g_decl_objs(l0218).M4126 = l021A
      l021A = l021A + gc2FD6
      g_decl_objs(l0218).M4150 = l021A
      l021A = l021A + g_decl_objs(l0218).Size
      If g_decl_objs(l0218).statSize Then
        Assume g_decl_objs(l0218).M4178 = 0, "no pStat"
        g_decl_objs(l0218).M4178 = l021A
        l021A = l021A + g_decl_objs(l0218).statSize
      End If
    End If
  Next
  gblsCountVars = l021A
Else
If gIsVB_Ver Then
  l021A = gv3318.M337F
Else
  l021A = gv2EAC.M337F
End If
  For l0218 = 1 To gEXEdef1
    g_decl_objs(l0218).M4126 = g_decl_objs(l0218).gbls_sorted
    g_decl_objs(l0218).M4150 = g_decl_objs(l0218).gbls_sorted
  Next
  For l0218 = gEXEdef1 + 1 To g_decl_num
    g_decl_objs(l0218).M4126 = l021A
    l021A = l021A + gc2FD6
    g_decl_objs(l0218).M4150 = l021A
    l021A = l021A + g_decl_objs(l0218).Size
  Next
  For l0218 = 1 To g_decl_num
    l0220 = g_decl_objs(l0218).statSize
    If l0220 Then
      g_decl_objs(l0218).M4178 = l021A
      l021A = l021A + l0220
    End If
  Next
  gblsCountVars = l021A
End If
Exit Sub

L21A9E:
  ShowMsgBox mc022E
End Sub

Sub sub0FA2(p01F6 As Control)
Dim l01F8
Dim l01FA As Long
Dim l01FC As String
Dim l01FE As Integer
Dim l0200 As Long
Dim l0202 As Integer
  p01F6.Clear
  p01F6.ListIndex = -1
  For l01F8 = 1 To g_decl_num
    p01F6.AddItem toHex16(g_decl_objs(l01F8).M4126) & " " & g_decl_objs(l01F8).Name_M411A
    l0202 = gLWhatEver + l01F8
    l01FA = g_gbl(l0202).M49E2
    p01F6.AddItem toHex16(l01FA) & " gbl " & g_decl_objs(l01F8).M3F23 & "[" & Hex$(g_decl_objs(l01F8).Size) & "]"
    l01FE = g_gbl(l0202).M4950
    While l01FE
      l0200 = gEXEdefEntries(l01FE).M689A + l01FA
      l01FC = toHex16(l0200) & " " & fn1B19(l01FE)
      p01F6.AddItem l01FC
      p01F6.ItemData(p01F6.NewIndex) = l01FE
      l01FE = gEXEdefEntries(l01FE).Count
    Wend
  Next
End Sub

Sub sub0FB5(p0206 As Integer, p020A As ListBox)
  DbgVar p0206 = gv16BC, "current sub"
  sub0FC7 gL2Min, p020A, True
  sub0FC7 gL3Min, p020A, False
End Sub

Sub sub0FC7(p020C As Integer, p0210 As ListBox, p0212 As Integer)
Dim l0214 As Integer
  If p0212 Then p0210.Clear: p0210.ListIndex = -1
  l0214 = g_gbl(p020C).M49D8
  If l0214 Then p0210.AddItem fn2023(l0214) & " @" & fn0226(l0214)
  l0214 = g_gbl(p020C).M4950
  While l0214 > 0
    Stop
    p0210.AddItem toHex16(CLng(gEXEdefEntries(l0214).M689A)) & " " & fn1B19(l0214)
    p0210.ItemData(p0210.NewIndex) = l0214
    l0214 = gEXEdefEntries(l0214).Count
  Wend
End Sub

Sub sub0FE8()
Dim l0234 As Integer
Dim l0238
Dim l023A As Long
Dim l023C As Integer
' Const mc023E = 14 ' &HE%
' Const mc0246 = 13 ' &HD%
' Const mc0248 = 38 ' &H26%
  DbgVar gblsCountVars, "gbls count"
  ReDim gv30F0(gblsCountVars)
If gIsVB5 Then
  DbgMileStone "exegbl offtables"
Else
  If gEXEdef2 > gFormsCount Then
    Seek g_hFileExe, fn1483(gv2EF8(mc023E))
    l023C = gv2EF8(mc0246)
    For l0238 = gEXEdef2 To gFormsCount + 1 Step -1
      Get g_hFileExe, , l023A
      gv30F0(l023A) = -(l0238 + gL0Max)
      g_gbl(l0238 + gL0Max).M49D8 = l023C
      gReloc_Arr(l023C).M36B3_count = mc0248
      l023C = CreateSym(fn202F(l023C) + &H10)
    Next
  End If
  For l0238 = 1 To gFormsCount
    l0234 = l0238 + gEXEdef1
    g_gbl(gL0Max + l0238).M49D8 = g_decl_objs(l0234).M41D7
    DbgVar gv30F0(g_decl_objs(l0234).gbls_sorted) = 0, "gbls sorted"
    gv30F0(g_decl_objs(l0234).gbls_sorted) = -(gLWhatEver + l0234)
    gv30F0(g_decl_objs(l0234).M4126) = -(gL0Max + l0238)
  Next
End If
  For l0234 = 1 To g_decl_num
    sub0FFC gLWhatEver + l0234, gv30F0()
    If g_decl_objs(l0234).statSize Then
      l023A = g_decl_objs(l0234).M4126 + g_decl_objs(l0234).M416F
      l023A = l023A + g_decl_objs(l0234).statSize
      If l023A > gblsCountVars Then
        DbgMileStone "missing global space"
        gblsCountVars = l023A
        ReDim Preserve gv30F0(gblsCountVars)
      End If
      gv30F0(g_decl_objs(l0234).M4126 + g_decl_objs(l0234).M416F) = -(gL1Max + l0234)
      sub0FFC gL1Max + l0234, gv30F0()
    End If
  Next
End Sub

Sub sub0FFC(p0136 As Integer, p0138() As Integer)
Dim l013E As Integer
  l013E = g_gbl(p0136).M4950
  Do While l013E > 0
    l013E = fn1C51(p0138(), l013E, p0136)
  Loop
End Sub

Sub sub101B()
Dim l025E As Integer
  Assume gL3Min, "lcl scope"
  Assume gL2Min, "para scope"
  Assume gv315A, "lcls"
  Assume gv315C, "params"
  ReDim gv3126(gv315A To gv315C)
  sub0FFC gL3Min, gv3126()
  sub0FFC gL2Min, gv3126()
  l025E = gEXEdef_Subs(gv16BC).M4289
  If l025E Then sub0A5E l025E
End Sub

Sub sub107E(p012A As Integer)
' Const mc012C = 4096 ' &H1000%
  If p012A <= 0 Then Exit Sub
  If gEXEdefEntries(p012A).M689A >= 0 Then Exit Sub
  If gEXEdefEntries(p012A).M6894.M36B3_count And mc012C Then
    gEXEdefEntries(p012A).M6894.M36B3_count = gEXEdefEntries(p012A).M6894.M36B3_count Xor mc012C
    gIs3158 = True
  End If
End Sub

Sub sub109D(ByVal pv0134 As Integer)
  If pv0134 Then
    gv2FDC = gEXEdef_Subs(gv16BC).lOffsetSym_M3C6A
  Else
    gv2FDC = 0
  End If
End Sub

Sub sub10BA(ByVal pv0164 As Integer, p0166 As Integer, p0168 As Integer)
Dim l016A As Integer
  pv0164 = pv0164 And p0168
  l016A = p0166 And p0168
  If pv0164 <> l016A Then
    If l016A = 0 Then
      p0166 = pv0164 Or (p0166 And Not p0168)
    End If
  End If
End Sub

Sub sub10C9(p016C As Integer, p016E As Integer)
  If p016C <> p016E Then
    If p016E = 0 Then
      p016E = p016C
    End If
  End If
End Sub

Sub sub10D6(p0172 As T4B07, p0176 As T4B07)
Dim l0178 As Integer
Dim l017A As Integer
' Const mc017C = 1 ' &H1%
' Const mc017E = 2 ' &H2%
' Const mc0180 = 10 ' &HA%
' Const mc0182 = 4 ' &H4%
' Const mc0184 = 3 ' &H3%
' Const mc0186 = 5 ' &H5%
' Const mc0188 = 9 ' &H9%
' Const mc018A = 32 ' &H20%
  l0178 = p0172.M6894.M42D9
  l017A = p0176.M6894.M42D9
  If l0178 <> l017A Then
    If l0178 = 0 Then
    ElseIf l017A = 0 Then
      ErrBeep
      sub10BA p0172.M6894.M36B3_count, p0176.M6894.M36B3_count, mc0100
      sub10C9 p0172.M6894.M42D1, p0176.M6894.M42D1
      p0176.M6894.M42D9 = l0178
    End If
  ElseIf l017A = 0 Then
    sub10C9 p0172.Size, p0176.Size
    sub10C9 p0172.M6894.M42D1, p0176.M6894.M42D1
    sub10BA p0172.M6894.M36B3_count, p0176.M6894.M36B3_count, mc0100
  Else
    sub10BA p0172.M6894.M36B3_count, p0176.M6894.M36B3_count, mc0100
    Select Case l0178
    Case mc017C, mc017E
    Case mc0180, mc0182
      sub10C9 p0172.M6894.M42D1, p0176.M6894.M42D1
    Case mc0184
      If fn108E(p0172.M6894.M42D1, p0176.M6894.M42D1) Then
        DbgMileStone "match objects"
      End If
    Case mc0186
      If fn108E(p0172.M6894.M42D1, p0176.M6894.M42D1) Then
        DbgMileStone "match arrays"
      End If
    Case mc0188
      If fn108E(p0172.M6894.M42D1, p0176.M6894.M42D1) Then
        DbgMileStone "match types"
      End If
      ShowMsgBox mc018A
    End Select
  End If
End Sub

Sub sub10E5()
  Assume gv16BA, "current module"
  gL1Min = gL1Max + gv16BA
  sub0FFC gL1Min, gv30F0()
End Sub

Sub sub110E()
  frmScopes.TextBox = gL3Min
End Sub

Sub sub112F()
Dim i%
Dim l01BA As Integer
Dim decl As Integer
Dim subIndex As Integer
  ReDim g_gbl(gEXEdef6_Num_scopes)
  
  For i = 1 To g_decl_num
    decl = i + gLWhatEver
    g_gbl(decl).M36B3_count = gc3018 Or i
    g_gbl(decl).M49D8 = g_decl_objs(i).M4131
    g_gbl(decl).M49E2 = g_decl_objs(i).M4126
    g_gbl(decl).Size = g_decl_objs(i).Size
    decl = i + gL1Max
    g_gbl(decl).M36B3_count = gc301A Or i
    g_gbl(decl).M49D8 = g_decl_objs(i).M415B
    g_gbl(decl).M49E2 = g_decl_objs(i).M4178
    g_gbl(decl).Size = g_decl_objs(i).statSize
  Next
  
  For i = 1 To gEXEdef2
    decl = i + gL0Max
    l01BA = i + gEXEdef1
    g_gbl(decl).M36B3_count = gc3020 Or i
  Next
  
  For i = 1 To gEXEdef_Arr5Count
: subIndex = gEXEdef_Subs(i).lOffsetSym_M3C6A
    If fn202F(subIndex) Then
If gIsVB5 < 0 Then
      DbgMileStone "scopes_alloc native"
      Exit Sub
ElseIf gIsVB5 > 0 Then
      Get g_hFileExe, fn1483(subIndex) + 4, gv2F50
Else
      Get g_hFileExe, fn1483(subIndex), gv2F50
End If
    Else
      DbgMileStone "Sub ohne Loc?"
    End If
    decl = i + gL3Max
    g_gbl(decl).M36B3_count = gc301C Or i Or gc302A
    g_gbl(decl).M49D8 = subIndex
    g_gbl(decl).Size = gv2F50.M5E8D - gv2FD8
    decl = i + gL2Max
    
    If gIsVB_Ver Or gEXEdef_Subs(i).ModuleIndex > gEXEdef1 Then
      g_gbl(decl).M36B3_count = gv301E Or i
    Else
      g_gbl(decl).M36B3_count = gv301E Or i Or gc302A
    End If
    
    g_gbl(decl).Size = gv2F50.M5E83 + gv2FDA
  
  Next
  
  sub113F
  sub0FE8
End Sub

Sub ScanModule()
  InitSth2
  sub112F
End Sub

Sub InitSth2()
Dim inivar As Integer
  gEXEdefEntriesCount1 = 0: ReDim gEXEdefEntries(100): gEXEdefEntriesCount2 = 0
  gIs3124 = True
  gIs3158 = False
  gLWhatEver = 0:      gL0Min = 0:   inivar = inivar + g_decl_num
  gL0Max = inivar: inivar = inivar + gEXEdef2
  gL1Max = inivar: gL1Min = 0: inivar = inivar + g_decl_num
  gL2Max = inivar: gL2Min = 0: inivar = inivar + gEXEdef_Arr5Count
  gL3Max = inivar: gL3Min = 0: inivar = inivar + gEXEdef_Arr5Count
  gL4Max = inivar
  gEXEdef6_Num_scopes = inivar
End Sub

