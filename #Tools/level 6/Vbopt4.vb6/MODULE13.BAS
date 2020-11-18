Attribute VB_Name = "MODULE13"
' Module13
Option Explicit
Const mc018E = 17 ' &H11%
Const mc0248 = 32 ' &H20%
Const mc026A = 1 ' &H1%
Const mc02E0 = 4096 ' &H1000%
Const mc02E2 = 8192 ' &H2000%
Const mc0302 = 8 ' &H8%
Const mc0360 = 2 ' &H2%
Const mc03A2 = 2 ' &H2%

Function fn0C32(p0314%, p0316 As Long) As Integer
Dim l0318 As Integer
  If p0316 <= 0 Then Exit Function
  l0318 = CreateSym(p0316)
  sub0C21 p0314, l0318
  fn0C32 = l0318
End Function

Function fn0C85(p032C As Integer) As String
Dim l032E As Integer
Dim l0330 As String
Dim l0334 As T6859
  l032E = gEXEdef_Subs(p032C).lOffsetSym_M3C6A
  l032E = gReloc_Arr(l032E).subRef_M3C6A
  While l032E < 0
    l032E = gReloc_Arr(-l032E).subRef_M3C6A
  Wend
  l0334.subRef_M3C6A = l032E
  l0330 = gEXEdef_Arr4_imports(l032E).Name
  fn0C85 = l0330
  l032E = InStr(l0330, ":")
  If l032E Then
    l0334.M6866 = Left$(l0330, l032E - 1)
    l0330 = Mid$(l0330, l032E + 1)
  End If
  l032E = InStr(l0330, Chr$(0))
  If l032E = 0 Then
    l0334.Name = l0330
  Else
    l0334.Name = Left$(l0330, l032E - 1)
    l0330 = Mid$(l0330, l032E + 1)
    l0334.M6426 = Asc(Mid$(l0330, 1, 1))
    l0334.M36B3_count = Asc(Mid$(l0330, 2, 1))
    l0334.M4CCD = Asc(Mid$(l0330, 3, 1))
    l0334.M388F = Asc(Mid$(l0330, 4, 1))
    l0334.M574B = Mid$(l0330, 5)
  End If
  gv2FBE = l0334
End Function

Function fn0C98(p0338 As Integer) As String
Dim l033A As String
Dim l033C As Integer
  l033A = fn2023(gEXEdef_Subs(p0338).lOffsetSym_M3C6A)
  l033C = InStr(l033A, Chr$(0))
  If l033C Then l033A = Left$(l033A, l033C - 1)
  fn0C98 = l033A
End Function

Function fn0DD8(p02F2 As Integer) As String
Dim l02F6 As String * 4
  Select Case p02F2
  Case gv2A44: l02F6 = "???"
  Case gc2A46: l02F6 = "b/w"
  Case gc2A48: l02F6 = "$-0"
  Case gc2A4A: l02F6 = "A$-0"
  Case gc2A4C: l02F6 = "U$-0"
  Case gc2A4E: l02F6 = "byte"
  Case gc2A50: l02F6 = "word"
  Case gc2A52: l02F6 = "long"
  Case gc2A54: l02F6 = "asm0"
  Case gc2A5C: l02F6 = "meth"
  Case gc2A5A: l02F6 = "proc"
  Case gc2A5E: l02F6 = "sub"
  Case gc2A56: l02F6 = "dll"
  Case gc2A58: l02F6 = "rec"
  Case gc2A60 To gc2A62
    l02F6 = "var" & Hex$(p02F2 - gc2A60)
  Case gc2A64: l02F6 = "Hdr"
  Case gc2A66: l02F6 = "Sub"
  Case gc2A68: l02F6 = "Mem"
  Case gc2A72: l02F6 = "Type"
  Case gc2A6A: l02F6 = "Form"
  Case gc2A70: l02F6 = "Clas"
  Case Is > gc2A78 + mc018E: Exit Function
  Case Is >= gc2A78: l02F6 = "T" & Format$(p02F2 - gc2A78)
  Case Else: l02F6 = Str$(p02F2)
  End Select
  fn0DD8 = l02F6
End Function

Function fn0EC1(p0308 As Integer) As String
  Select Case p0308
  Case gc2EDC
    fn0EC1 = "Sub Main"
  Case gc2EDE
    fn0EC1 = "Forms"
  Case gc2EE0
    fn0EC1 = "VBP"
  Case gc2EE2
    fn0EC1 = "Project"
  Case gc2EE4
    fn0EC1 = "PrjPtrs"
  Case gc2EEA
    fn0EC1 = "OlbRec"
  Case gc2EEC
    fn0EC1 = "Modules"
  Case gc2EEE
    fn0EC1 = "Classes"
  Case gc2EF0
    fn0EC1 = "GUIDs"
  Case gc2EF2
    fn0EC1 = "ClassOffsets"
  Case gc2EF4
    fn0EC1 = "Declares"
  Case gc2EF6
    fn0EC1 = "^data"
  Case Else: fn0EC1 = "Tab" & Format$(p0308)
  End Select
End Function

Function fn0DE7(p1%, p01DC$, p01DE%) As Integer
' called (p1%, p2$, p3%)
Stop
Dim l01E0 As Integer
Dim l01E2 As Integer
  l01E0 = MakeSubRef(p01DE, p01DC, gc2A84)
  fn0DE7 = l01E0
  gReloc_Arr(p01DE).M36B3_count = gc2A56
  l01E2 = fn0ECF(p01DE, l01E0, gc2A56, 0)
End Function

Function fn0DF5(p01E8%, p01EA&) As Integer
Dim l01EC As Integer
Dim l01EE As String
Dim l01F0 As Integer
  gEXEdef5 = gEXEdef5 + 1
  l01EE$ = "sub" & CStr(gEXEdef5)
  l01EC = MakeSubRef(p01E8, l01EE, gEXEdef5)
  gReloc_Arr(p01E8).M36B3_count = gc2A5E
  l01F0 = fn0ECF(p01E8, l01EC, gc2A5E, p01EA)
  fn0DF5 = l01F0
End Function

Function fn0E06(p01F4%, p01F6%, p01F8&, p01FA%) As Integer
Dim l01FC As Integer
  l01FC = fn14AE(p01F6, p01F8, p01FA, False)
  fn0E06 = l01FC
  If l01FC Then sub0C21 p01F4, l01FC
End Function

Function fn0ECF(p01A4 As Integer, p01A6 As Integer, p01A8 As Integer, p01AA As Long) As Integer
Dim l01AC As Integer
  gEXEdef_Arr5Count = gEXEdef_Arr5Count + 1
  l01AC = gEXEdef_Arr5Count
  If l01AC <= 1 Then
    ReDim gEXEdef_Subs(50)
  End If
  If l01AC > UBound(gEXEdef_Subs) Then ReDim Preserve gEXEdef_Subs(l01AC + 50)
  gEXEdef_Subs(l01AC).lOffsetSym_M3C6A = p01A4
  gEXEdef_Subs(l01AC).M36B3_count = p01A8
  gEXEdef_Subs(l01AC).M428F = p01AA
  gEXEdef_Arr4_imports(p01A6).TypeCode_M5F96 = -l01AC
  fn0ECF = l01AC
End Function

Function sub09DB(p0202%, ByVal p0204 As Long, p0206 As Integer, p0208 As Integer, p020A As Long) As Integer
Dim l020C As Integer
Dim l020E As Integer
Dim l0210 As String
Dim l0212 As Integer
Dim l0216
Dim l0218 As Long
Dim l021A As Long
  l0218 = fn202F(p0202)
If gIsVB_Ver Then
Else
  If p0204 Then
    If l0218 And &HFFFF& Then
      l0218 = l0218 And &HFFFF0000
    End If
    l0218 = l0218 Or (p0204 And &HFFFF&)
  End If
End If
  Select Case p0208
  Case gc2A5C
    l0210 = "meth"
    gEXEdef4 = gEXEdef4 + 1
  Case gc2A5A
    l0210 = "proc"
    gEXEdef5 = gEXEdef5 + 1
If gIsVB5 Then
    DbgMileStone "must create modules"
Else
    For l0216 = gEXEdef1 To 1 Step -1
      If p020A = g_decl_objs(l0216).gbls_sorted Then Exit For
    Next
    Assume l0216, "module assigned"
End If
  Case gc2A60 To gc2A62
    l0210 = gEXEdef_Arr4_imports(p0206).Name
    gReloc_Arr(p0202).M36B3_count = p0208
    l020C = MakeSubRef(p0202, l0210 & Hex$(p020A), (p020A))
    Exit Function
  Case Else
    ErrBeep
  End Select
  l020C = RetOffsetSym(l0218, l0210, 0)
  gReloc_Arr(p0206).M36B3_count = p0208
  If l020C = 0 Then ErrBeep: Exit Function
  gReloc_Arr(p0206).subRef_M3C6A = gReloc_Arr(l020C).subRef_M3C6A
  If p0208 >= gc2A60 Then
    gReloc_Arr(l020C).M36B3_count = p0208
  Else
    gReloc_Arr(l020C).M36B3_count = gc2A64
  End If
  l020E = gReloc_Arr(l020C).subRef_M3C6A
  l0212 = gEXEdef_Arr4_imports(l020E).TypeCode_M5F96
  If l0212 >= 0 Then
    If l0212 > 0 And l0212 <> p0208 Then
      MsgBox "type mismatch", , "add_vbsub"
      Exit Function
    End If
    l0212 = fn0ECF(l020C, l020E, p0208, p020A)
    If l0216 Then
      gEXEdef_Subs(l0212).ModuleIndex = l0216
      l020C = g_decl_objs(l0216).M3F13
      If l020C = 0 Then
        g_decl_objs(l0216).M3F1B = l0212 - 1
      Else
        DebugPrintVar g_decl_objs(l0216).M3F1B + g_decl_objs(l0216).M3F13 = l0212 - 1, "sub count"
      End If
      g_decl_objs(l0216).M3F13 = l020C + 1
    End If
    l0210 = l0210 & CStr(l0212)
    gEXEdef_Arr4_imports(l020E).Name = l0210
  Else
    MsgBox "typ " & Str(l0212), , "add_vbsub"
  End If
  sub09DB = l020E
  l021A = VAtoRaw(l0218)
If gIsVB5 Then
  Get g_hFileExe, l021A + 8, l020C
  l020C = CreateSym(l0218 - ToUnSigned(l020C))
  If l020C Then
    gReloc_Arr(l020C).M36B3_count = gc2A66
    gReloc_Arr(l020C).subRef_M3C6A = gReloc_Arr(p0206).subRef_M3C6A
  Else
    ErrBeep
  End If
  Exit Function
End If
  Get g_hFileExe, l021A + 4, l020C
  l020C = CreateSym(l0218 - ToUnSigned(l020C))
  If l020C Then
    gReloc_Arr(l020C).M36B3_count = gc2A66
    gReloc_Arr(l020C).subRef_M3C6A = gReloc_Arr(p0206).subRef_M3C6A
  Else
    ErrBeep
  End If
  If gIsVB_Ver Then
    sub0E13 l0212, l021A
  Else
    sub0E24 l0212, l021A
  End If
End Function

Sub sub0C11(p1%, p2%)
L1C69E:
'Stop
End Sub

Sub sub0C21(p030E As Integer, p0310 As Integer)
  If p030E > 0 Then
    If p0310 = gv2EF8(p030E - 1) Then
      gv2EF8(p030E) = 0
      Exit Sub
    End If
  End If
  gv2EF8(p030E) = p0310
End Sub

Sub sub0C42(p031A%, ByVal p031C As Long, p031E%)
Dim l0320 As Integer
Dim l0322 As Long
  l0322 = fn202F(p031A)
  l0322 = l0322 Or (p031C And &HFFFF&)
  l0320 = CreateSym(l0322)
  If l0320 Then sub0C51 l0320, p031E
End Sub

Sub sub0C51(p0324 As Integer, p0326 As Integer)
Dim l0328%
  gReloc_Arr(p0326).M36B3_count = gc2A58
  gReloc_Arr(p0324).M36B3_count = gc2A72
  If gReloc_Arr(p0324).subRef_M3C6A = 0 Then
    sub0C60 p0324
  End If
  gReloc_Arr(p0326).subRef_M3C6A = -p0324
  Get g_hFileExe, fn1483(p0324), l0328
  sub0C6E p0324, l0328
End Sub

Sub sub0CA7()
Dim l033E%
Dim l0340 As Integer
Dim l0342 As String
Dim l0344%
Dim l0346 As Long
Dim l0348 As Long
Dim l034A As String
Dim l034C As Integer
Dim l034E As Long
Dim l0350 As Integer
Dim l0352 As Integer
Dim l0354%
Dim l0356 As Integer
Dim l0358 As Integer
Dim l035A As Integer
Dim l035C%
' Const mc0360 = 2 ' &H2%
Dim l039C As Integer
Dim l039E As Integer
' Const mc03A2 = 2 ' &H2%
Dim l03A4%

'Not a VBProgram
If gIsVB_Ver > 0 Then
ElseIf g_IsNE Then
Else
  ShowMsgBox mc0360
  Exit Sub
End If
  
  LoadVBDat
  sub0CB5
  sub0CC1
  frmScan.control5 = "Scanning"
  frmScan.control4 = "Locations"
  frmScan.control3 = "Segment"
  frmScan.control2 = "Tables"
  DoEvents
  gEXEdef_Arr5Count = 0: gEXEdef5 = 0: gEXEdef4 = 0
  sub0CCE
  sub0CDD
  l0358 = gEntryPoint_Page
  Assume l0358 > 0, "WinEntry"
If gIsVB_Ver Then
  gReloc_Arr(l0358).M36B3_count = gc2A54
  l0348 = fn202F(l0358)
  l0340 = fn1624(gPESections3(1).VOffset)
  Do While l0340
    l0346 = gv161C(l0340).VOffset - 2
    If l0346 >= l0348 Then Exit Do
    Get g_hFileExe, fn15A9(l0346), l0344
    If l0344 = &H25FF Then
      l0358 = CreateSym(l0346)
      gReloc_Arr(l0358).M36B3_count = gc2A54
      l035A = gv161C(l0340).subRef_M3C6A
      If gReloc_Arr(l035A).M36B3_count < gc2A54 Then
        gReloc_Arr(l035A).M36B3_count = gc2A54
      End If
      l0358 = CreateSym(l0346 + 6)
    End If
    l0340 = fn164F(l0340)
  Loop
  l0358 = gEntryPoint_Page
  l035A = fn0E06(0, gEntryPoint_Page, 1, l0340)
  gv2EA8 = gv2EF8(0)
  Assume gv2EA8, "tab0"
If gIsVB5 Then
    sub0CEE
    Exit Sub
End If
  Get g_hFileExe, fn1483(gv2EA8), gv323E
  l0340 = 0
  l035A = fn0E06(1, gv2EA8, &H30, l0340)
  l035A = fn0E06(2, gv2EA8, &H50, l0340)
  l035A = fn0E06(3, gv2EA8, &H54, l0340)
  l035A = fn0E06(4, gv2EA8, &H58, l0340)
  l035A = fn0E06(5, gv2EA8, &H5C, l0340)
  gv2ED8 = gv2EF8(1)
  Assume gv2ED8, "tab1"
  Get g_hFileExe, fn1483(gv2ED8), gv3318
  l0340 = 0
  l035A = fn0E06(11, gv2ED8, &HC, l0340)
  l035A = fn0E06(12, gv2ED8, &H10, l0340)
  l035A = fn0E06(13, gv2ED8, &H14, l0340)
  l035A = fn0E06(14, gv2ED8, &H18, l0340)
  l035A = fn0E06(15, gv2ED8, &H20, l0340)
  l035A = fn0E06(16, gv2ED8, &H28, l0340)
  l035A = fn0E06(17, gv2ED8, &H30, l0340)
  gv315E = gv2EF8(17)
  If gv315E Then
    gv3160 = CreateSym(fn202F(gv315E) + 4)
  Else
    MsgBox "no .data?"
  End If
  gFormsCount = gv3318.MFormsCount
  gEXEdef1 = gv3318.ModuleTableEntryCount
  gEXEdef2 = gv3318.M33C1
Else
  gReloc_Arr(l0358).M36B3_count = gc2A54
  l0340 = fn1624(fn202F(l0358))
  l035A = CreateSym(gv161C(l0340).VOffset + 4)
  l035A = fn0E06(0, l035A, 0, l0340)
  gv2EA8 = gv2EF8(0)
  Assume gv2EA8, "tab0"
  Get g_hFileExe, fn1483(gv2EA8), gv2E4C
  l0340 = 0
  l035A = fn0E06(-2, gv2EA8, &H28, l0340)
  l035A = fn0E06(-1, gv2EA8, &H2C, l0340)
  l035A = fn0E06(1, gv2EA8, &H30, l0340)
  l035A = fn0E06(2, gv2EA8, &H4C, l0340)
  l035A = fn0E06(3, gv2EA8, &H50, l0340)
  l035A = fn0E06(4, gv2EA8, &H54, l0340)
  l035A = fn0E06(5, gv2EA8, &H58, l0340)
  gv2ED8 = gv2EF8(1)
  Assume gv2ED8, "tab1"
  Get g_hFileExe, fn1483(gv2ED8), gv2EAC
  l0340 = 0
  l035A = fn0E06(11, gv2ED8, &HC, l0340)
  l035A = fn0E06(12, gv2ED8, &H10, l0340)
  l035A = fn0E06(13, gv2ED8, &H14, l0340)
  l035A = fn0E06(14, gv2ED8, &H18, l0340)
  l035A = fn0E06(15, gv2ED8, &H1E, l0340)
  l035A = fn0E06(16, gv2ED8, &H24, l0340)
  gFormsCount = gv2EAC.MFormsCount
  gEXEdef1 = gv2EAC.ModuleTableEntryCount
  gEXEdef2 = gv2EAC.M33C1
End If
  l0358 = gv2EF8(gc2EF4)
  If l0358 Then
    gv0456 = fn202F(l0358)
  Else
    gv0456 = 0
  End If
  g_decl_num = gFormsCount + gEXEdef1
  ReDim g_decl_objs(g_decl_num), gv28BA(gEXEdef2)
  For l0344 = 1 To gFormsCount
    gv28BA(l0344) = l0344 + gEXEdef1
  Next
  load_400ss
If gIsVB5 Then
  sub0D0B
Else
  sub0D1C
  sub0D28
End If
  sub0D39
If gIsVB_Ver Then
    l0346 = fn202F(gEntryPoint_Page)
    sub0115 l0346
    l0340 = fn1624(l0346)
    While l0340
      l0346 = gv161C(l0340).VOffset
      Select Case gv161C(l0340).subRef_M3C6A
      Case gv315E
        Get g_hFileExe, gVAToRaw2 + l0346 + 4, gSingleChar
        l0344 = Asc(gSingleChar)
        Select Case l0344
        Case &HE9
          Get g_hFileExe, gVAToRaw2 + l0346 - 8, l0344
          Assume l0344 = &HC933, "xor"
          Assume gv161C(gv161C(l0340).M54BB).VOffset = l0346 - 5, "method"
          Get g_hFileExe, gVAToRaw2 + l0346 - 14, l035C
          l0358 = CreateSym(l0346 - 8)
          l0352 = gv161C(gv161C(l0340).M54BB).subRef_M3C6A
          l0356 = sub09DB(l0352, 0, l0358, gc2A5C, ToUnSigned(l035C))
          l035A = CreateSym(l0346 - 15)
          gReloc_Arr(l035A).M36B3_count = gReloc_Arr(l0358).M36B3_count
          gReloc_Arr(l035A).subRef_M3C6A = gReloc_Arr(l0358).subRef_M3C6A
        Case &H8D
          Get g_hFileExe, gVAToRaw2 + l0346 + 4, l0344
          Assume l0344 = &H888D, "lea"
          Assume gv161C(gv161C(l0340).Count).VOffset = l0346 + 13, "proc"
          Get g_hFileExe, gVAToRaw2 + l0346 + 6, l035C
          l0358 = CreateSym(l0346 - 2)
          l0352 = gv161C(gv161C(l0340).Count).subRef_M3C6A
          l0356 = sub09DB(l0352, 0, l0358, gc2A5A, ToUnSigned(l035C))
        End Select
      Case gv3160
        Get g_hFileExe, gVAToRaw2 + l0346 - 2, l0344
        Select Case l0344
        Case &H158B
          l0358 = CreateSym(l0346 - 2)
          gReloc_Arr(l0358).M36B3_count = gc2A54
          Get g_hFileExe, gVAToRaw2 + l0346 + 19, gSingleChar
          l0344 = Asc(gSingleChar)
          Assume l0344 = &HE8, "rel call"
          Get g_hFileExe, , l0348
          l0348 = l0348 + l0346 + 24
          l035A = CreateSym(l0348)
          Assume l035A, "pointer loc"
          While gReloc_Arr(l035A).subRef_M3C6A < 0
            l035A = -gReloc_Arr(l035A).subRef_M3C6A
          Wend
          If fn2017(l035A) Then
            l035A = gReloc_Arr(l035A).subRef_M3C6A
            Select Case gEXEdef_Arr4_imports(l035A).TypeCode_M5F96
            Case gc2A84
              Get g_hFileExe, gVAToRaw2 + l0346 + 6, l0354
              l0356 = l0354 \ 4
              l0354 = l0356 * 12
              l0350 = fn1631(gv0456 + l0354, False)
              l035A = gv161C(l0350).subRef_M3C6A
              l0342 = fn1561(fn202F(l035A))
              l0350 = fn164F(l0350)
              l035A = gv161C(l0350).subRef_M3C6A
              l034A = fn0218(fn202F(l035A))
              l0342 = l0342 & ":" & l034A
              l0356 = fn0DE7(l0356, l0342, l0358)
            End Select
          End If
        End Select
      End Select
      l0340 = gv161C(l0340).Count
    Wend
Else
    For l033E = 3 To gSegCount
      l0340 = fn13E0(l033E, 0)
      If l0340 Then Exit For
    Next
    While l0340
      If fn087D(gv161C(l0340).VOffset) <> l033E Then
        l033E = fn087D(gv161C(l0340).VOffset)
        frmScan.control2 = "#" & Str$(l033E)
        DoEvents
        sub0D4A l033E
      End If
      l0358 = gv161C(l0340).subRef_M3C6A
      If fn2017(l0358) Then
        l034E = gv161C(l0340).VOffset - gv1062
        l0358 = gReloc_Arr(l0358).subRef_M3C6A
        Select Case gEXEdef_Arr4_imports(l0358).TypeCode_M5F96
        Case gc2A88
          Get g_hFileExe, gv1056 + l034E - 11, l035C
          Get g_hFileExe, gv1056 + l034E - 6, l0354
          l0352 = fn13E0(l033E, l034E - 3)
          l0352 = gv161C(l0352).subRef_M3C6A
          l0358 = fn06CC(l033E, l034E - 22)
          l0356 = sub09DB(l0352, l0354, l0358, gc2A5A, ToUnSigned(l035C))
        Case gc2A8A
          Get g_hFileExe, gv1056 + l034E - 12, l035C
          Get g_hFileExe, gv1056 + l034E - 6, l0354
          l0352 = fn13E0(l033E, l034E - 3)
          l0352 = gv161C(l0352).subRef_M3C6A
          l0358 = fn06CC(l033E, l034E - 13)
          l0356 = sub09DB(l0352, l0354, l0358, gc2A5C, ToUnSigned(l035C))
          l035A = fn06CC(l033E, l034E - 9)
          gReloc_Arr(l035A).M36B3_count = gReloc_Arr(l0358).M36B3_count
          gReloc_Arr(l035A).subRef_M3C6A = gReloc_Arr(l0358).subRef_M3C6A
        Case gc2A84
          Get g_hFileExe, gv1056 + l034E - 15, l0354
          l0356 = l0354 \ 4
          l0354 = l0356 * 10
          l0350 = fn1631(gv0456 + l0354, False)
          l0358 = gv161C(l0350).subRef_M3C6A
          l0342 = fn0218(fn202F(l0358))
          l0350 = fn149E(l0350)
          l0358 = gv161C(l0350).subRef_M3C6A
          l0342 = l0342 & ":" & fn0218(fn202F(l0358))
          l0358 = fn06CC(l033E, l034E - 22)
          l0356 = fn0DE7(l0356, l0342, l0358)
        Case gc2A86
          Get g_hFileExe, gv1056 + l034E - 5, l0354
          l0352 = fn13E0(l033E, l034E - 8)
          If gv161C(l0352).VOffset - gv1062 = l034E - 8 Then
            l0352 = gv161C(l0352).subRef_M3C6A
            l0358 = fn06CC(l033E, l034E - 11)
            sub0C42 l0352, l0354, l0358
          Else
            ErrBeep
          End If
        Case gc2A8E To gc2A94
          l0356 = fn06CC(l033E, l034E - 9)
          Get g_hFileExe, gv1056 + l034E - 3, l035C
          l0356 = sub09DB(l0356, 0, l0358, gc2A60 + (gEXEdef_Arr4_imports(l0358).TypeCode_M5F96 - gc2A8E), ToUnSigned(l035C))
        End Select
      End If
      l0340 = fn164F(l0340)
    Wend
End If
If gIsVB_Ver Then
Else
  frmScan.control4 = "Fixups"
  l033E = 0
  l0340 = fn13E0(2, 0)
  While l0340
    If fn087D(gv161C(l0340).VOffset) <> l033E Then
      l033E = fn087D(gv161C(l0340).VOffset)
      frmScan.control2 = "#" & Str$(l033E)
      DoEvents
      sub0D4A l033E
      l0358 = SomeAdrConv(gv1062)
      GoSub L1EA6C
    End If
    While gv161C(l0340).VOffset > l034E
      GoSub L1EA6C
    Wend
    If l039C = 0 Then
      If gv161C(l0340).M36B3_count = mc03A2 Then
        sub00B9 gv161C(l0340)
      End If
    End If
    l0340 = fn164F(l0340)
  Wend
End If
  frmScan.control4 = "Forms"
  DoEvents
  l0358 = gv2EF8(gc2EDE)
  For l03A4 = 1 To gFormsCount
    If l0358 = 0 Then Exit For
    Get g_hFileExe, fn1483(l0358), gv2C22
    l0340 = fn124C(gv2C22.M41AD, False)
    If l0340 = 0 Then Exit For
      For l0344 = gFormsCount To 1 Step -1
        l034C = gv28BA(l0344)
        If g_decl_objs(l034C).M41B7 = l0340 Then Exit For
      Next
      If l0344 = 0 Then Exit For
    If gv2C22.M5A03 <> l0344 - 1 Then
      Exit For
    End If
    If gv2C22.M34FD Then
      l035A = fn14AE(l0358, &H2C, 0, True)
      Assume l035A, "form RC ptr"
      If l035A = 0 Then Exit For
      gReloc_Arr(l0358).M36B3_count = gc2A78 + 2
      sub0D5B l0344, l034C, l035A
      frmScan.control2 = g_decl_objs(l034C).M3F23
      DoEvents
      sub0D6B l034C, VAtoRaw(fn202F(l035A)), 0
    Else
      Assume gv097C, "Classes not decompiled"
      Assume gv2C22.M6692 = 0, "no RCptr"
      gReloc_Arr(l0358).M36B3_count = gc2A78 + 2
      frmScan.control2 = g_decl_objs(l034C).M3F23
    End If
    l0358 = CreateSym(fn202F(l0358) + &H30)
  Next
  For l0344 = 1 To gFormsCount
    l034C = gv28BA(l0344)
    If g_decl_objs(l034C).M41D7 Then
    Else
If g_IsDemoVersion Then
      Assume gv097C, "no Classes decompiled"
End If
      sub0D79 l034C
      l0342 = g_decl_objs(l034C).Name_M411A
      l03A4 = InStr(l0342, ".")
      If l03A4 Then l0342 = Left$(l0342, l03A4 - 1)
      If Len(g_decl_objs(l034C).M3F23) = 0 Then
        g_decl_objs(l034C).M3F23 = l0342
      End If
      g_decl_objs(l034C).Name_M411A = l0342 & ".CLS"
      frmScan.control2 = g_decl_objs(l034C).M3F23
      DoEvents
    End If
  Next
  sub0D8B
  l0358 = gv2EF8(gc2EF0)
  If l0358 Then
    Seek g_hFileExe, fn1483(l0358)
    gReloc_Arr(l0358).M36B3_count = 1
    While l0358 And gReloc_Arr(l0358).M36B3_count <= 1
      Get g_hFileExe, , gv0B84
      l0344 = fn124C(gv0B84, True)
      If l0344 Then
        l0342 = fn120C(l0344)
        l0344 = Len(l0342)
      End If
      If l0344 > 1 Then
        sub0D9B l0358, l0342
      End If
      gReloc_Arr(l0358).M36B3_count = gc2A70
      l0358 = CreateSym(fn202F(l0358) + &H10)
    Wend
  End If
If gIsVB_Ver Then
  frmScan.control4 = "Labels"
  l0358 = gEntryPoint_Page
  While l0358
    l035A = gReloc_Arr(l0358).Count
    If gReloc_Arr(l0358).M36B3_count = gc2A46 Then
      If fn202F(l035A) - fn202F(l0358) >= 14 Then
        l0340 = fn1624(fn202F(l0358))
        Select Case gv161C(l0340).VOffset - fn202F(l0358)
        Case 3
          l034E = fn1483(l0358)
          Get g_hFileExe, l034E + 8, l0344
          If l0344 = &HE950 Then
            sub0C51 gv161C(l0340).subRef_M3C6A, l0358
          Else
            ErrBeep
          End If
        Case 4
          l034E = fn1483(l0358)
          Get g_hFileExe, l034E + 10, l0344
          If l0344 = &HE950 Then
            sub0C51 gv161C(l0340).subRef_M3C6A, l0358
          Else
            ErrBeep
          End If
        End Select
       End If
    End If
    l0358 = gReloc_Arr(l0358).Count
  Wend
End If
  If gv097C Then
    If gFormsCount <> gv0982 + gv0980 Or gEXEdef5 <> gv0B0A Or gEXEdef4 <> gv0B0C Then
      ShowMsgBox mc0302
    End If
  Else
    gExeNameOnly = gExeFileNameOnly2
    g_ExeFilePath = gv_ExeFilePath
    If gFormsCount > 8 Then
      ShowMsgBox "Too many forms for this version - aborting"
      Stop
    End If
  End If
  sub0DAA
  sub0DB8
  sub0DC9
Exit Sub

L1EA6C:
  l039E = gReloc_Arr(l0358).M36B3_count
  Select Case l039E
  Case gc2A54 To gc2A62
    l039C = True
  Case Else
    l039C = False
  End Select
  l0358 = gReloc_Arr(l0358).Count
  If l0358 Then
    l034E = fn202F(l0358)
  Else
    l034E = gv1066
  End If
Return
End Sub

Sub sub0D5B(p01B8 As Integer, p01BA As Integer, p01BC As Integer)
Dim l01C0 As String
Dim l01C4 As T60A3
If gIsVB5 Then
  DbgMileStone "implement add_exeform5"
  Exit Sub
End If
  Assume p01B8 > 0, "Form#"
  Assume gv28BA(p01B8) = p01BA, "f2m set"
  DebugPrintVar p01BC, "Form with RC"
  gReloc_Arr(p01BC).M36B3_count = gc2A6A
  Get g_hFileExe, fn1483(p01BC), l01C4
  Get g_hFileExe, , gv0120
  Get g_hFileExe, , gFormMainStruc
  l01C0 = Space$(gFormMainStruc.NameSize): Get g_hFileExe, , l01C0
  g_decl_objs(p01BA).M3F23 = l01C0
  g_decl_objs(p01BA).M41D7 = p01BC
  g_decl_objs(p01BA).M41E8 = Asc(l01C4.M606F)
  sub0D79 p01BA
End Sub

Private Sub sub0D8B()
Dim l0292
Dim l0294%
Dim l0296 As Integer
Dim l0298 As Integer
Dim l029A As Long
Dim l029C
Dim l029E As Integer
Dim l02A0 As Integer
Dim l02A2 As Integer
Dim l02A4 As Integer
Dim l02A6 As Integer
Dim l02A8 As Integer
Dim l02AA As Integer
Dim l02AC As Integer
Dim l02AE As Integer
Dim l02B0 As Integer
Dim l02B2 As Integer
Dim l02B4 As Integer
Dim l02B6 As Integer
Dim l02B8 As Integer
Dim l02D2 As Integer
Dim l02D4 As Integer
Dim l02D6 As Integer
' Const mc02E0 = 4096 ' &H1000%
' Const mc02E2 = 8192 ' &H2000%
  If gEXEdef4 = 0 Then Exit Sub
  l029E = gv2EF8(gc2EEE)
  Assume l029E, "tab sym"
  l02A4 = fn14AE(l029E, 0, l0294, True)
  For l0292 = 1 To gFormsCount
    l02A6 = gv28BA(l0292)
    Assume l02A6 = gEXEdef1 + l0292, "f2m set"
    If gv097C Then
      l02AC = l02A6
      l02A8 = gv094A(l02AC).M3F1B + 1
    End If
    l02A4 = gv161C(l0294).subRef_M3C6A
    Assume l02A4 > 0, "sym8"
    l02A2 = fn14AE(l02A4, 4, 0, True)
    Assume l02A2 > 0, "sym3"
sub0C11 l02A6, l02A2
If gIsVB_Ver Then
    Get g_hFileExe, fn1483(l02A2), gv3450
    l0296 = 0
    l02A0 = fn14AE(l02A2, &H1C, l0296, True)
    Assume l02A0, "ptr"
    gReloc_Arr(l02A0).M36B3_count = gc2EE6
sub0C11 l02A6, l02A0
    l029A = gv3450.M6720 + gv161C(l0296).VOffset + 4
Else
    Get g_hFileExe, fn1483(l02A2), gv2E2E
    l0296 = 0
    l02A0 = fn14AE(l02A2, &H18, l0296, True)
    Assume l02A0, "ptr3"
    gReloc_Arr(l02A0).M36B3_count = gc2EE6
sub0C11 l02A6, l02A0
    l029A = gv2E2E.M6720 + gv161C(l0296).VOffset + 4
End If
    l0296 = fn164F(l0296)
    If l0296 Then
      l0298 = (l029A - gv161C(l0296).VOffset) \ 4
      l02B8 = 0
      Do While l0298 > 0
        l02B6 = gv161C(l0296).subRef_M3C6A
If gIsVB_Ver Then
        If l02B6 Then
          Select Case gReloc_Arr(l02B6).M36B3_count
          Case gc2A60 To gc2A62
            l0298 = l0298 - 1
            l02B8 = l02B8 + 1
          Case Else
            Exit Do
          End Select
        End If
Else
        If l02B6 Then
          Select Case gReloc_Arr(l02B6).M36B3_count
          Case gc2A60 To gc2A62
            l0298 = l0298 - 1
            l02B8 = l02B8 + 1
          Case Else
            Exit Do
          End Select
        End If
End If
        l0296 = fn164F(l0296)
      Loop
      If l0298 > 0 Then
        l02B2 = gReloc_Arr(l02B6).subRef_M3C6A
        If l02B2 > 0 Then
          Assume gReloc_Arr(gEXEdef_Arr4_imports(l02B2).lOffsetSym_M3C6A).M36B3_count = gc2A64, "sub hdr"
          l02B4 = -gEXEdef_Arr4_imports(l02B2).TypeCode_M5F96
          Assume l02B4 > 0, "sub ref"
          l02AE = l02B4
          l02B0 = l02B4
          If gv097C Then
            l02D6 = l02A8 + l0298
            l02D2 = l02A8
          Else
            l02D2 = 0
          End If
          l02D4 = 0
          For l029C = 1 To l0298
            l02B6 = gv161C(l0296).subRef_M3C6A
            l02B2 = gReloc_Arr(l02B6).subRef_M3C6A
            Assume gReloc_Arr(gEXEdef_Arr4_imports(l02B2).lOffsetSym_M3C6A).M36B3_count = gc2A64, "hdr ref"
            l02B4 = -gEXEdef_Arr4_imports(l02B2).TypeCode_M5F96
            If l02B4 < l02AE Then l02AE = l02B4
            If l02B4 > l02B0 Then l02B0 = l02B4
            l02AA = 0
            If gv097C Then
              If l02D2 Then
                Do While l02D2 < l02D6
                  l02D2 = l02D2 + 1
                  If gv0AD6(l02D2).Flags And mc02E0 Then
                    l02AA = l02D2
                    Exit Do
                  End If
                Loop
                If l02AA = 0 Then
                  l02D2 = 0
                  l02D4 = l02A8
                End If
              End If
              If l02AA Then
              ElseIf l02D4 Then
                Do While l02D4 < l02D6
                  l02D4 = l02D4 + 1
                  If gv0AD6(l02D4).Flags And mc02E2 Then
                    l02AA = l02D4
                    Exit Do
                  End If
                Loop
              End If
              Assume l02AA, "match methods"
            End If
            Select Case gEXEdef_Subs(l02B4).M36B3_count
            Case gc2A5C
            Case gc2A60 To gc2A62
              Assume False, "more vars"
            Case gc2A5A
              Assume False, "Procs in BAS"
            Case Else
              ShowMsgBox mc0248
              ErrBeep
            End Select
            gEXEdef_Subs(l02B4).ModuleIndex = gEXEdef1 + l0292
            If l02AA Then
              gEXEdef_Subs(l02B4).M4289 = l02AA
              gv0AD6(l02AA).M404D = l02B4
            End If
            l0296 = fn164F(l0296)
          Next
          l02AE = l02AE - 1
          If l02B0 - l02AE <> l0298 Then
            ErrBeep
            Exit Sub
          End If
          Assume l02B0 - l02AE = l0298, "sub count"
          g_decl_objs(l0292 + gEXEdef1).M3F1B = l02AE
          g_decl_objs(l0292 + gEXEdef1).M3F13 = l0298
        Else
          MsgBox "loc ohne symbol"
        End If
      End If
    End If
If gIsVB_Ver Then
    If gv3450.M6729 Then sub0EA2 l02A6, l02A2
Else
    If gv2E2E.M6729 Then sub0EA2 l02A6, l02A2
End If
    l0294 = fn164F(l0294)
  Next
End Sub

Sub sub0DAA()
Dim l02F8 As Integer
Dim l02FA As Integer
Dim l02FC As Integer
Dim l02FE
Dim l0300 As Integer
' Const mc0302 = 8 ' &H8%
  For l02FE = 1 To gv097C
    l02F8 = l02FE
    l02FC = g_decl_objs(l02F8).M3F1B
    l0300 = g_decl_objs(l02F8).M3F13
    Assume l02FC >= 0, "module nr"
    l02FA = gv094A(l02FE).M3F1B + 1
If gIsVB5 Then
    If gv094A(l02FE).M3F13 - 1 < l0300 Then
      DbgMileStone "strip assumed declares"
      l0300 = gv094A(l02FE).M3F13 - 1
      g_decl_objs(l02F8).M3F13 = l0300
    End If
End If
    If gv094A(l02FE).M3F13 - 1 <> l0300 Then
      ShowMsgBox mc0302
    Else
      gv0AD6(l02FA - 1).M404D = -g_decl_objs(l02F8).M41D7
      gv0AD6(l02FA).M404D = -g_decl_objs(l02F8).M4131
      If gIsVB5 Or l02FE <= gv097E Then
        For l02F8 = 1 To l0300
          If gEXEdef_Subs(l02FC + l02F8).ModuleIndex Then
            If gIsVB5 < 0 Then
            Else
              Assume gEXEdef_Subs(l02FC + l02F8).ModuleIndex = l02FE, "wrong mod for sub"
            End If
          ElseIf gIsVB5 Then
            gEXEdef_Subs(l02FC + l02F8).ModuleIndex = l02FE
          End If
          gv0AD6(l02FA + l02F8).M404D = l02FC + l02F8
          gEXEdef_Subs(l02FC + l02F8).M4289 = l02FA + l02F8
        Next
      End If
    End If
  Next
End Sub

Sub sub0E13(p023C As Integer, p023E As Long)
Dim l0240%
Dim l0242%
Dim l0244 As Integer
Dim l0246%
' Const mc0248 = 32 ' &H20%
Dim l024A
Dim l024C
  Get g_hFileExe, p023E - 2, l0242
  Select Case l0242
  Case &H378: gEXEdef_Subs(p023C).M4277 = gc298A
  Case &H37A: gEXEdef_Subs(p023C).M4277 = gc298E Or gc2990
  Case &H37C: gEXEdef_Subs(p023C).M4277 = gc2994
  Case &H37E: gEXEdef_Subs(p023C).M4277 = gc2996
  Case &H380: gEXEdef_Subs(p023C).M4277 = gc29A0 Or gc29A4
  Case &H382: gEXEdef_Subs(p023C).M4277 = gc29A2
  Case &H386: gEXEdef_Subs(p023C).M4277 = gc2998
  Case &H388: gEXEdef_Subs(p023C).M4277 = gc299A
  Case &H38A
  Case &H2BC
    Get g_hFileExe, p023E, l0246
    l0246 = l0246 + 4
    Get g_hFileExe, p023E + 4, l0240
    Seek g_hFileExe, p023E - ToUnSigned(l0240)
    Do
      Get g_hFileExe, , l0242
      Select Case l0242
      Case &H704, &H706
        Get g_hFileExe, , gv0120
      Case Else
        Exit Do
      End Select
    Loop
    Select Case l0242
    Case &H422
      l0244 = gc298E Or gc2990 Or gc29A8
    Case &H292
      Get g_hFileExe, , l0242
      Select Case l0242
      Case &H1EE
        Get g_hFileExe, , l0242
        If l0242 = &H42A Then l0244 = gc29A2
      Case &H1FE
        Get g_hFileExe, , l0242
        If l0242 = &H420 Then l0244 = gc298A
      Case &H22E
        Get g_hFileExe, , l0242
        If l0242 = &H432 Then l0244 = gc2996
      Case &H23E
        Get g_hFileExe, , l0242
        If l0242 = &H434 Then l0244 = gc29A4 Or gc29A0
      End Select
    Case &H750
      Get g_hFileExe, , gv0120
      If gv0120 Then Exit Sub
      Get g_hFileExe, , l0242
      If l0242 = &H424 Then l0244 = gc2994 Or gc2998 Or gc299A
    End Select
    If l0244 = 0 Then Exit Sub
    Get g_hFileExe, , l0240
    If l0240 = l0246 Then
      gEXEdef_Subs(p023C).M4277 = l0244
    End If
  Case Else
    Get g_hFileExe, p023E - 4, l0242
    If l0242 = &H50A Then
      Get g_hFileExe, , l0240
      If l0240 = &H10 Then
        gEXEdef_Subs(p023C).M4277 = gc29A8
      Else
        ShowMsgBox mc0248
      End If
    Else
      Get g_hFileExe, p023E - 6, l0242
      Get g_hFileExe, , l024A
      Get g_hFileExe, , l024C
      Select Case l0242
      Case &H71C, &H71E
        If l024A + l024C = gv2FD8 Then
          gEXEdef_Subs(p023C).M4277 = gc29AA
          DbgMileStone "def. retvar"
        End If
      End Select
    End If
  End Select
End Sub

Sub sub0E24(p0230 As Integer, p0232 As Long)
Dim l0234%
Dim l0236
Dim l0238 As Integer
Dim l023A
  Get g_hFileExe, p0232 - 2, l0236
  Select Case l0236
  Case &H1032
    gEXEdef_Subs(p0230).M4277 = gc2994 Or gc2998 Or gc299A
  Case &H103F
    gEXEdef_Subs(p0230).M4277 = gc298A
  Case &H104A
    gEXEdef_Subs(p0230).M4277 = gc298E Or gc2990
  Case &H1069
    GoTo L1A66A
  Case Else
    Get g_hFileExe, p0232 - 4, l0236
    If l0236 = &H1012 Then
      Get g_hFileExe, , l0234
      Select Case l0234
      Case &H10
        l0238 = gc29A8
      Case 8
        l0238 = gc29A0 Or gc29A2 Or gc29A4
      Case 4
        l0238 = gc2996
      Case Else
        DbgMileStone "detect var Type"
      End Select
      gEXEdef_Subs(p0230).M4277 = gc29AA Or l0238
    End If
  End Select
Exit Sub

L1A66A:
  Get g_hFileExe, p0232, l023A
  DbgMileStone "retoff16=?"
  l023A = l023A + 2
  Get g_hFileExe, p0232 + 4, l0234
  Seek g_hFileExe, p0232 - ToUnSigned(l0234)
  Do
    Get g_hFileExe, , l0236
    Select Case l0236
    Case &H451B, &H3F6F
      Get g_hFileExe, , gv0120
    Case Else
      Exit Do
    End Select
  Loop
  Select Case l0236
  Case &H321B
    Get g_hFileExe, , l0236
    Select Case l0236
    Case &H19A6
      Get g_hFileExe, , l0236
      If l0236 = &H2A3B Then l0238 = gc29A2
    Case &H17C3
      Get g_hFileExe, , l0236
      If l0236 = &H2A0F Then l0238 = gc298A
    Case &H1869
      Get g_hFileExe, , l0236
      Select Case l0236
      Case &H2A54
        l0238 = gc2996
      Case &H2A65
        l0238 = gc29A0 Or gc29A4
      End Select
    Case &H2A1D
      l0238 = gc298E Or gc2990 Or gc29A6
    End Select
  Case &H3283
    Get g_hFileExe, , gv0120
    If gv0120 Then Exit Sub
    Get g_hFileExe, , l0236
    Select Case l0236
    Case &H2A2A
      l0238 = gc2994 Or gc2998 Or gc299A
    End Select
  Case Else
    l0238 = gc29AA
    DbgMileStone "form: Type retoff?"
  End Select
  If l0238 = 0 Then Exit Sub
  Get g_hFileExe, , l0234
  If l0234 = l023A Then
    gEXEdef_Subs(p0230).M4277 = l0238
  ElseIf l0238 = gc29AA Then
    gEXEdef_Subs(p0230).M4277 = l0238
  End If
End Sub

Sub sub0E35(p022A As T6859)
Dim l022C As String
  l022C = p022A.M6866
  If Len(l022C) Then l022C = l022C & ":"
  l022C = l022C & p022A.Name & Chr$(0)
  l022C = l022C & Chr$(p022A.M6426) & Chr$(p022A.M36B3_count) & Chr$(p022A.M4CCD)
  l022C = l022C & Chr$(p022A.M388F) & p022A.M574B
  gEXEdef_Arr4_imports(p022A.subRef_M3C6A).Name = l022C
  gIs3124 = True
End Sub

Sub sub0E40(p0250 As Integer, p0254 As Control)
Dim l0256
Dim l0258 As Integer
  p0254.Clear
  l0258 = g_decl_objs(p0250).M41E0
  For l0256 = -1 To g_decl_objs(p0250).M41E8
    p0254.AddItem gv0D3E(l0258 + l0256).M3F23
    p0254.ItemData(p0254.NewIndex) = l0258 + l0256
  Next
End Sub

Sub sub0E51(p0264 As Control)
Dim l0268
' Const mc026A = 1 ' &H1%
  p0264.Clear
  For l0268 = 1 To g_decl_num
    If g_decl_objs(l0268).M36B3_count = mc026A Then
      p0264.AddItem g_decl_objs(l0268).M3F23
      p0264.ItemData(p0264.NewIndex) = l0268
    End If
  Next
End Sub

Sub sub0E62(p026E As Control)
Dim l0270
  p026E.Clear
  p026E.ListIndex = -1
  For l0270 = 1 To g_decl_num
    p026E.AddItem g_decl_objs(l0270).Name_M411A
    p026E.ItemData(p026E.NewIndex) = l0270
  Next
End Sub

Sub sub0E72(p0274 As Control, p0276 As Integer, p3%)
' called (p1 As object, p2%, p3%)
Dim l027A As Integer
Dim l027C As Integer
Dim l0280 As String
  If p0276 Then
    l027A = -g_decl_objs(p0276).M41D7
    If l027A Then
      l0280 = "(form)": GoSub L1B224
    End If
    l027C = g_decl_objs(p0276).M3F1B
    For l027A = l027C + 1 To l027C + g_decl_objs(p0276).M3F13
      l0280 = fn2023(gEXEdef_Subs(l027A).lOffsetSym_M3C6A)
      GoSub L1B224
    Next
  End If
  p0274.ListIndex = -1
Exit Sub

L1B224:
  p0274.AddItem l0280
  p0274.ItemData(p0274.NewIndex) = l027A
Return
End Sub

Sub sub0E82(p0286 As Control)
Dim l028A%
Dim l028C As String
If g_IsDemoVersion Then
Else
  For l028A = 0 To gc2A78 + mc018E
    l028C = fn0DD8(l028A)
    If Len(l028C) = 0 Then Exit For
    p0286.AddItem l028C
    p0286.ItemData(p0286.NewIndex) = l028A
  Next
End If
End Sub

Sub sub0E93()
End Sub

Sub sub0EB3(p02E4 As Integer)
Dim l02EA As Integer
Dim l02EC As Long
  l02EA = gEXEdef_Subs(p02E4).lOffsetSym_M3C6A
  l02EC = fn202F(l02EA)
  sub0115 l02EC
  gModuleVAEnd = l02EC
  If gIsVB_Ver < 0 Then
    DbgMileStone "eval native hdr"
  Else
    If gIsVB_Ver = 5 Then
      Get g_hFileExe, gVAToRaw2 + l02EC + 4, gv2F50
    Else
      Get g_hFileExe, gVAToRaw2 + l02EC, gv2F50
    End If
    gModuleVAStart = gModuleVAEnd - ToUnSigned(gv2F50.M5E97)
  End If
End Sub

