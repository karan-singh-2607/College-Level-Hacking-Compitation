Attribute VB_Name = "MODULE28"
' Module28
Option Explicit
Const mc001C = -4096 ' &HF000%
Const mc001E = 4095 ' &HFFF%
Const mc0020 = 20480 ' &H5000%
Const mc0022 = 28672 ' &H7000%
Const mc0024 = 12288 ' &H3000%
Const mc00A6 = 192 ' &HC0%
Const mc00A8 = 12288 ' &H3000%
Const mc00D0 = -832 ' &HFCC0%
Const mc00D2 = 20480 ' &H5000%
Const mc00D6 = 28672 ' &H7000%
Const mc00DA = 64 ' &H40%
Const mc00DC = 128 ' &H80%
Const mc00DE = -20480 ' &HB000%
Const mc00FE = -8192 ' &HE000%
Const mc0100 = 16384 ' &H4000%
Const mc0102 = 24576 ' &H6000%
Const mc0104 = 8192 ' &H2000%
Const mc0106 = 4096 ' &H1000%
Const mc011A = 112 ' &H70%
Const mc011C = 1 ' &H1%
Const mc0122 = 2 ' &H2%
Const mc0124 = 16384 ' &H4000%
Const mc0128 = 3 ' &H3%
Const mc0144 = 32 ' &H20%
Const mc018E = 63 ' &H3F%
Const mc0190 = 11 ' &HB%
Const mc0192 = -32768 ' &H8000%
Const mc01B2 = 1 ' &H1%
Const mc01B4 = -4096 ' &HF000%
Const mc01B6 = 8192 ' &H2000%
Const mc01FC = 15 ' &HF%

Function fn1B08(p00B6 As Integer) As String
  fn1B08 = fn1B6C(gv36E6(gEXEdefEntries(p00B6).M689A) And mc001E)
End Function

Function fn1B19(variable As Integer) As String
Dim l00C4 As Long
Dim l00C6 As String
Dim l00C8 As Integer
Dim l00CA As Integer
' Const mc00D0 =-832 ' &HFCC0%
' Const mc00D2 = 20480 ' &H5000%
' Const mc00D6 = 28672 ' &H7000%
' Const mc00DA = 64 ' &H40%
' Const mc00DC = 128 ' &H80%
' Const mc00DE =-20480 ' &HB000%
Dim l00E0 As Integer
Dim l00E2 As Integer
Dim l00E4 As Integer
Dim l00E6%
' Const mc00FE =-8192 ' &HE000%
' Const mc0100 = 16384 ' &H4000%
' Const mc0102 = 24576 ' &H6000%
' Const mc0104 = 8192 ' &H2000%
' Const mc0106 = 4096 ' &H1000%
  If variable < 0 Then fn1B19 = fn104E(-variable): Exit Function
  DbgVar variable, "variable"
  If variable = 0 Then
    fn1B19 = "?v:0?"
    Exit Function
  End If
  l00C4 = gEXEdefEntries(variable).M68A4
  If gv097C Then
    Select Case gEXEdefEntries(variable).M6894.M36B3_count And mc00D0
    Case mc00D2
If gIsVB5 Then
  DbgMileStone "implement VB vars"
Else
      l00CA = gv360C(l00C4)
      l00C8 = l00CA And mc001E
      If l00C8 Then
        fn1B19 = fn1B9B(l00C8, gv35D4(l00C8))
        Exit Function
      End If
End If
    Case mc00D6
      If l00C4 <= UBound(gv3678) Then
        l00CA = gv3678(l00C4)
        l00C8 = l00CA And mc001E
        If l00C8 Then
          fn1B19 = fn1B9B(l00C8, gv3640(l00C8))
          Exit Function
        End If
      End If
    Case mc00A8
      If l00C4 >= gv3718 And l00C4 < gv371A Then
        l00CA = gv36E6(l00C4)
        l00C8 = l00CA And mc001E
        If l00C8 Then
          Select Case l00CA And mc001C
          Case mc0024
            fn1B19 = fn1B6C(l00C8)
            Exit Function
          Case gc35D0
            fn1B19 = fn0F4A(l00C8)
            Exit Function
          End Select
        End If
      End If
    Case mc00DA, mc00DC, mc00A6
      l00CA = gv36E6(l00C4)
      l00C8 = l00CA And mc001E
      If l00C8 Then
        fn1B19 = fn1B9B(l00C8, gv36AE(l00C8))
        Exit Function
      End If
    Case mc00DE
      l00E0 = gEXEdefEntries(variable).owningScope
      DbgVar l00E0 = gEXEdefEntries(variable).owningScope, "owned var"
      l00CA = fn1C80(l00E0)
      Assume l00CA > 0, "udt scope"
      If l00CA <= g_userTypesCount Then
        l00E2 = fn1CA9(l00CA)
        If l00E2 Then
          l00E4 = gv38BA(l00E2).M4950
          For l00E6 = l00E4 + 1 To l00E4 + gv38BA(l00E2).M4959
            If gv38F0(l00E6).hFile = l00C4 Then
              fn1B19 = fn0F4A(l00E6)
              Exit Function
            End If
          Next
          DbgMileStone "no matching member in VB-Type?"
        End If
      End If
      l00C6 = "M"
    End Select
  End If
  Select Case gEXEdefEntries(variable).M6894.M36B3_count And mc00FE
  Case mc0100
    l00C6 = "g"
  Case mc0102
    l00C6 = "s"
  Case mc0104
    If l00C4 > 0 Then
      l00C6 = "p"
    ElseIf gEXEdefEntries(variable).M6894.M36B3_count And mc0106 Then
      l00C6 = "l"
    Else
      l00C6 = "t"
    End If
  Case mc00DA, mc00DC, mc00A6
    DbgMileStone "expect never reached!"
    l00C6 = "p"
  Case mc00DE, Is < 0
    l00C6 = "M"
  End Select
  fn1B19 = l00C6 & toHex16(l00C4)
End Function

Function fn1B28(ByVal pv013C As Integer) As String
Dim l013E As Integer
Dim l0140 As Integer
Dim l0142 As String
' Const mc0144 = 32 ' &H20%
  If gv097C Then
    If pv013C >= LBound(gv36E6) And pv013C < UBound(gv36E6) Then
      l0140 = gv36E6(pv013C)
      l013E = l0140 And mc001E
    End If
  End If
  If l013E Then
    Select Case l0140 And mc001C
    Case gc35D0
      l0142 = fn1B28(pv013C - gv38F0(l013E).hFile)
      l0142 = l0142 & fn0F4A(l013E)
      fn1B28 = l0142
      Exit Function
    Case mc0024
      fn1B28 = fn1B6C(l013E)
      Exit Function
    Case 0
    Case Else
      ShowMsgBox mc0144
    End Select
  End If
  If pv013C > 0 Then
    fn1B28 = "p" & toHex16(pv013C)
  ElseIf gEXEdefEntries(l013E).M6894.M36B3_count And mc0106 Then
    fn1B28 = "l" & toHex16(pv013C)
  Else
    fn1B28 = "t" & toHex16(pv013C)
  End If
End Function

Function fn1B34(p0148 As Long) As String
Dim l014A As Integer
Dim l014C As Integer
Dim l014E As String
  If gv097C Then
    l014C = gv360C(p0148)
    l014A = l014C And mc001E
  End If
  If l014A Then
    Select Case l014C And mc001C
    Case mc0020
      fn1B34 = fn1B5B(l014A)
    Case gc35D0
      l014E = fn1B34(p0148 - gv38F0(l014A).hFile)
      l014E = l014E & fn0F4A(l014A)
      fn1B34 = l014E
    Case 0
    Case Else
      ShowMsgBox mc0144
    End Select
  Else
    fn1B34 = "g" & toHex16(p0148)
  End If
End Function

Function fn1B41(ByVal p0152 As Long) As String
Dim l0154 As Integer
Dim l0156 As Integer
Dim l0158 As String
  If gv097C Then
    l0156 = gv3678(p0152)
    l0154 = l0156 And mc001E
  End If
  If l0154 Then
    Select Case l0156 And mc001C
    Case mc0022
      fn1B41 = fn1B8A(l0154)
    Case gc35D0
      l0158 = fn1B41(p0152 - gv38F0(l0154).hFile)
      l0158 = l0158 & fn0F4A(l0154)
      fn1B41 = l0158
    End Select
  Else
    fn1B41 = "s" & toHex16(p0152)
  End If
End Function

Function fn1B4E(p015C As Integer, p0160() As T5679, p0168 As T5679) As Integer
Dim l016A
  If gv097C = 0 Then Exit Function
  For l016A = 1 To UBound(p0160)
    If p0160(l016A).hFile = p015C Then
      fn1B4E = l016A
      p0168 = p0160(l016A)
      Exit Function
    End If
  Next
End Function

Function fn1B5B(p0196 As Integer) As String
  If p0196 > gv360A Then
    ErrBeep2
    fn1B5B = "g#" & toHex16(p0196)
  Else
    fn1B5B = ExtractKeyWord(gv3606, gv35D4(p0196).StringBuffOffs)
  End If
End Function

Function fn1B6C(p01CE As Integer) As String
Dim l01D0 As Integer
  If p01CE <= UBound(gv36AE) Then
    l01D0 = gv36AE(p01CE).StringBuffOffs
    If l01D0 Then
      fn1B6C = ExtractKeyWord(gv36E0, l01D0)
      Exit Function
    End If
  Else
    DbgMileStone "ill. lcl var#"
    fn1B6C = "l?#" & Format$(p01CE)
    Exit Function
  End If
  DbgMileStone "unnamed VB local"
  If l01D0 <= 0 Then
    fn1B6C = "l" & toHex16(l01D0)
  Else
    fn1B6C = "p" & toHex16(l01D0)
  End If
End Function

Function fn1B7C(p01DE As Integer, p01E0%) As String
Dim l01E2 As String
  Select Case p01DE And mc00FE
  Case mc0100
    l01E2 = fn1B34(ToUnSigned(p01E0) + g_gbl(p01DE).M49E2)
  Case mc0102
    l01E2 = fn1B41(p01E0)
  Case mc0104
    l01E2 = fn1B28(p01E0)
  Case Else
    l01E2 = ".M" & toHex16(p01E0)
  End Select
  fn1B7C = l01E2
End Function

Function fn1B8A(ImportDllIndex As Integer) As String
  fn1B8A = ExtractKeyWord(gv3672, gv3640(ImportDllIndex).StringBuffOffs)
End Function

Function fn1B9B(p01F2%, p01F6 As T5679) As String
Dim l01F8 As String
Dim l01FA As Integer
' Const mc01FC = 15 ' &HF%
' Const mc01FE = 2 ' &H2%
  l01FA = p01F6.M6894.M36B3_count
  If l01FA And (mc00DA Or mc00DC) Then
    l01F8 = fn1B6C(p01F2)
  Else
    Select Case l01FA And mc00FE
    Case mc0100
      l01F8 = fn1B5B(p01F2)
    Case mc0102
      l01F8 = fn1B8A(p01F2)
    Case mc0104
      l01F8 = fn1B6C(p01F2)
    End Select
  End If
  Select Case l01FA And mc01FC
  Case 2
    l01F8 = l01F8 & " *" & p01F6.M6894.M42D1
  Case mc0190
    l01F8 = l01F8 & " Type " & gv38BA(p01F6.M6894.M42D1).Name
  End Select
  l01F8 = l01F8 & " [" & Hex$(p01F6.size) & "]"
  fn1B9B = l01F8
End Function

Sub sub0A3E(p010A As Integer, p010C As String, p0110 As T42C7, p0112 As Integer, p0114 As Integer)
Dim l0116 As Integer
Dim l0118 As Long
' Const mc011A = 112 ' &H70%
' Const mc011C = 1 ' &H1%
' Const mc0122 = 2 ' &H2%
' Const mc0124 = 16384 ' &H4000%
' Const mc0128 = 3 ' &H3%
  If p0110.M36B3_count And mc011A Then
    If Left$(p010C, 1) <> "^" Then p010C = "^" & p010C
  Else
    If Left$(p010C, 1) = "^" Then p010C = Mid$(p010C, 2)
  End If
  Select Case p0114
  Case mc011C
    l0116 = mc00D2
    l0118 = fn0B84(gv0A88, p0112, p0110)
  Case mc0122
    l0116 = mc0124
    l0118 = fn0B84(gv0A8A, p0112, p0110)
  Case mc0128
    l0116 = mc00D6
    l0118 = fn0B84(gv36AA, p0112, p0110)
  End Select
  p0110.M36B3_count = p0110.M36B3_count Or l0116
  Select Case l0116 And mc00FE
  Case mc0100
    gv360A = gv360A + 1
    If gv360A > UBound(gv35D4) Then ReDim Preserve gv35D4(gv360A + 10)
    QuoteWithParag p010C, gv3606, gv35D4(gv360A).StringBuffOffs
    gv35D4(gv360A).M6894 = p0110
    gv35D4(gv360A).size = p0112
    gv35D4(gv360A).M6877 = p010A
    gv35D4(gv360A).hFile = l0118
  Case mc0102
    gv3676 = gv3676 + 1
    If gv3676 > UBound(gv3640) Then ReDim Preserve gv3640(gv3676 + 10)
    QuoteWithParag p010C, gv3672, gv3640(gv3676).StringBuffOffs
    gv3640(gv3676).M6894 = p0110
    gv3640(gv3676).size = p0112
    gv3640(gv3676).M6877 = p010A
    gv3640(gv3676).hFile = l0118
  End Select
End Sub

Sub sub0A7A(p012A As Integer, p012C As String, p0130 As T42C7, p0132 As Integer)
Dim l0134 As Long
  gv36E4 = gv36E4 + 1
  If gv36E4 > UBound(gv36AE) Then ReDim Preserve gv36AE(gv36E4 + 10)
  If p0130.M36B3_count And mc011A Then
    If Left$(p012C, 1) <> "^" Then p012C = "^" & p012C
  Else
    If Left$(p012C, 1) = "^" Then p012C = Mid$(p012C, 2)
  End If
  p0132 = (p0132 + 1) And -2
  QuoteWithParag p012C, gv36E0, gv36AE(gv36E4).StringBuffOffs
  gv36AE(gv36E4).M6894 = p0130
  gv36AE(gv36E4).size = p0132
  If p0130.M36B3_count And mc00A6 Then
If gIsVB_Ver Then
    l0134 = fn0B84(gv371A, 4, p0130)
    If gv371A And 2 Then ErrBeep
Else
    l0134 = fn0B84(gv371A, p0132, p0130)
End If
  ElseIf p0130.M36B3_count And mc0100 Then
    ErrBeep
    l0134 = fn0B84(gv36AA, p0132, p0130)
  ElseIf p0130.M36B3_count And mc0106 Then
    l0134 = fn0B84(gv3718, -p0132, p0130)
  End If
  gv36AE(gv36E4).M6877 = p012A
End Sub

Sub sub0BB9(Dummy%)
   Stop
  ReDim gv36AE(20): gv36E4 = 0
If gIsVB_Ver Then
  gv2FDA = 8
  gv2FD8 = &HFF7C
Else
  gv2FDA = 6
  gv2FD8 = &HFFB0
End If
  gv371E = 0
  gv3718 = 0
  gv371C = 0
  gv371A = 0
End Sub

Sub sub0BC7(p01A2 As Integer, p2%)
Dim l01A4
Dim l01A6 As Integer
Dim l01A8 As Integer
Dim l01AE As Integer
Dim l01B0 As Integer
' Const mc01B2 = 1 ' &H1%
' Const mc01B4 =-4096 ' &HF000%
' Const mc01B6 = 8192 ' &H2000%
  l01A6 = gv371E
  If gv36E4 Then
    If p01A2 <> mc01B2 Or gIsVB_Ver Then
      l01A8 = gv371C
      For l01A4 = 1 To gv36E4
        l01AE = gv36AE(l01A4).M6894.M36B3_count
        l01B0 = (gv36AE(l01A4).size + 1) And -2
        If l01AE And mc00A6 Then
If gIsVB_Ver Then
          If l01B0 < 4 Then l01B0 = 4
End If
          gv36AE(l01A4).hFile = fn0B84(l01A8, l01B0, gv36AE(l01A4).M6894)
        ElseIf l01AE And mc00A8 Then
          gv36AE(l01A4).hFile = fn0B84(l01A6, -l01B0, gv36AE(l01A4).M6894)
        End If
      Next
      gv371A = l01A8
    Else
      l01A8 = gv371A
      For l01A4 = 1 To gv36E4
        l01AE = gv36AE(l01A4).M6894.M36B3_count
        l01B0 = (gv36AE(l01A4).size + 1) And -2
        DbgVar l01B0 > 0, "var size"
        If l01AE And mc00A6 Then
          l01A8 = l01A8 - l01B0
          gv36AE(l01A4).hFile = l01A8
        Else
          Select Case l01AE And mc01B4
          Case mc00A8
            gv36AE(l01A4).hFile = fn0B84(l01A6, -l01B0, gv36AE(l01A4).M6894)
          Case mc01B6
            DbgVar gv36AE(l01A4).hFile = 0, "dyn var"
          End Select
        End If
      Next
    End If
    If l01A6 < gv3718 Then gv3718 = l01A6
  End If
  sub1740
End Sub

Sub sub1740()
Dim l01B8
Dim l01BA As Long
Dim l01BC As Integer
  If gv16BC > 0 Then
    DebugPrintVar gv16B6 = gEXEdef_Subs(gv16BC).M4289, "sec loaded"
If gIsVB5 < 0 Then
    DbgMileStone "scopes"
    Exit Sub
Else
    DbgVar gv3718 >= gv315A, "lcl alloc"
    DbgVar gv371A <= gv315C, "para alloc"
End If
    If gv371A > gv315C Then gv315C = gv371A
    If gv3718 < gv315A Then gv315A = gv3718
    ReDim gv36E6(gv315A To gv315C)
  Else
    ReDim gv36E6(gv3718 To gv371A)
  End If
  For l01B8 = 1 To gv36E4
    l01BC = gv36AE(l01B8).M6894.M36B3_count
    If l01BC And (mc00A6 Or mc00A8) Then
      l01BA = gv36AE(l01B8).hFile
      gv36E6(l01BA) = l01B8 Or mc0024
      If l01BC And mc011A Then
      ElseIf (l01BC And mc018E) = mc0190 Then
        sub0F2A gv36E6(), l01BA, gv36AE(l01B8).M6894.M42D1
      End If
    End If
  Next
End Sub

Sub sub19DD()
Dim l016C
Dim l0172 As Long
Dim l0178 As Integer
Dim l017A As Long
' Const mc018E = 63 ' &H3F%
' Const mc0190 = 11 ' &HB%
' Const mc0192 =-32768 ' &H8000%
On Error GoTo L3A02E
  If gIsVB_Ver = -5 Then
    Exit Sub
  End If
  If gblsCountVars = 0 Then
    Exit Sub
  End If
  If gv097C = 0 Then Exit Sub
  Assume gEXEdef6_Num_scopes, "init scopes"
  Assume gblsCountVars, "init vars"
  ReDim gv360C(gblsCountVars)
  l0178 = 0
  
  For l016C = 1 To gv360A
    If gv35D4(l016C).M6877 <> l0178 Then
      l0178 = gv35D4(l016C).M6877
      Assume l0178, "prj match"
      l0172 = g_decl_objs(l0178).M4150
    End If
    l017A = ToUnSigned(gv35D4(l016C).hFile) + l0172
    GoSub L39FBC
  Next
  sub0F5B
Exit Sub

L39FBC:
    gv360C(l017A) = l016C Or mc0020
    If (gv35D4(l016C).M6894.M36B3_count And mc018E) = mc0190 Then
      sub0F2A gv360C(), l017A, gv35D4(l016C).M6894.M42D1
    End If
Return

L3A02E:
If gIsVB_Ver Then
  ErrBeep
  Resume
End If
  ShowMsgBox mc0192
End Sub

Sub sub19FF()
Dim l01D2
Dim l01D4 As Long
Dim l01D6 As Integer
Dim l01D8 As Long
  DebugPrintVar gv16B2 = gv16BA, "mod loaded"
  ReDim gv3678(g_decl_objs(gv16BA).statSize)
  If g_decl_objs(gv16BA).statSize = 0 Then
    DbgMileStone "statics"
    Exit Sub
  End If
  l01D8 = g_decl_objs(gv16BA).M4178
  For l01D2 = 1 To gv3676
    l01D6 = gv3640(l01D2).M6894.M36B3_count
    l01D4 = ToUnSigned(gv3640(l01D2).hFile)
    gv3678(l01D4) = l01D2 Or mc0022
    If l01D6 And mc011A Then
    ElseIf (l01D6 And mc018E) = mc0190 Then
      sub0F2A gv3678(), l01D4, gv3640(l01D2).M6894.M42D1
    End If
  Next
End Sub

Sub sub1AC1(p007A As Integer, p007E As Control)
Dim l0080 As Integer
Dim l0082
Dim l0084 As Integer
Dim l0086 As Long
  p007E.Clear
  p007E.ListIndex = -1
  If gv097E = 0 Then Exit Sub
  l0086 = ToUnSigned(gv094A(p007A).M3F46)
  p007E.AddItem "' Publics"
  l0084 = gv094A(p007A).M3F30
  For l0082 = 1 To gv094A(p007A).M3F3B
    GoSub L38A52
  Next
  p007E.AddItem "' Privates"
  l0084 = gv094A(p007A).M3F5B
  For l0082 = 1 To gv094A(p007A).M3F67
    GoSub L38A52
  Next
Exit Sub

L38A52:
  l0080 = l0082 + l0084
  p007E.AddItem fn1B5B(l0080)
  p007E.ItemData(p007E.NewIndex) = l0080
Return
End Sub

Sub sub1AD3(p0096 As Control)
Dim l0098%
  p0096.Clear
  p0096.ListIndex = -1
  For l0098 = 1 To gv360A
    p0096.AddItem fn1B5B(l0098)
    p0096.ItemData(p0096.NewIndex) = l0098
  Next
End Sub

Sub sub1AE5(p009A As Integer, p009E As Control)
Dim l00A0%
Dim l00A2 As Integer
' Const mc00A6 = 192 ' &HC0%
' Const mc00A8 = 12288 ' &H3000%
  p009E.Clear
  p009E.ListIndex = -1
  DebugPrintVar p009A = gv16B6, "sec current"
  l00A2 = mc00A6 Or mc00A8
  For l00A0 = 1 To gv36E4
    If gv36AE(l00A0).M6894.M36B3_count And l00A2 Then
      p009E.AddItem fn1B6C(l00A0)
      p009E.ItemData(p009E.NewIndex) = l00A0
    End If
  Next

L38C50:
  If p009E.ListCount = gv36E4 Then Exit Sub
  p009E.AddItem "--------------"
  For l00A0 = 1 To gv36E4
    If (gv36AE(l00A0).M6894.M36B3_count And l00A2) = 0 Then
      p009E.AddItem fn1B6C(l00A0)
      p009E.ItemData(p009E.NewIndex) = l00A0
    End If
  Next
End Sub

Sub sub1AF6(p00AA As Integer, p00AE As Control)
Dim l00B0%
  p00AE.Clear
  p00AE.ListIndex = -1
  If p00AA <> gv16B2 Then Exit Sub
  For l00B0 = 1 To gv3676
    p00AE.AddItem fn1B8A(l00B0)
    p00AE.ItemData(p00AE.NewIndex) = l00B0
  Next
End Sub

