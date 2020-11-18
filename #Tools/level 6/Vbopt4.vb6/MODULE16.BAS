Attribute VB_Name = "MODULE16"
' Module16
Option Explicit
Dim m001E() As T4BA0
Dim m0050 As Integer
Dim m0052 As String
Const mc006A = 255 ' &HFF%
Dim m006C() As Integer
Const mc0164 = 16384 ' &H4000%
Const mc01A8 = 57 ' &H39%
Const mc01B2 = 2 ' &H2%

Function fn11BA(p012C As String, p012E As String, p0132 As T4938) As Integer
Dim ownClassID As Integer
Dim ClassID As Integer
Dim l0138 As Integer
  If Len(p012C) = 0 Then
    DbgMileStone "name class"
    p012C = fn1285(p0132)
  End If
  ownClassID = fn124C(p0132, False)
  If ownClassID Then
    DbgMileStone "dup appclass"
    l0138 = gClassNamesPos(ownClassID).StringBuffOffs
    If l0138 Then
      DbgMileStone "check legal rename"
    End If
    If gClassNamesPos(ownClassID).M4B72 Then
      DbgMileStone "replace parent"
    End If
  Else
    ownClassID = fn11CA(p0132)
    Assume ownClassID, "class added"
  End If
  If gClassNamesPos(ownClassID).M4B5E Then ErrBeep
  
  ClassID = Findclass(p012E)
  If ClassID = 0 Then
    MsgBox p012E, , "unknown class"
  ElseIf ClassID = ownClassID Then
    DbgMileStone "ill. parent class (self)"
  Else
    l0138 = gClassNamesPos(ownClassID).Count
    LSet gClassNamesPos(ownClassID) = gClassNamesPos(ClassID)
    
    gClassNamesPos(ownClassID).M4B72 = 0
    LSet gClassNamesPos(ownClassID).ResObjName = p0132
    
    gClassNamesPos(ownClassID).Count = l0138
  End If
  
  QuoteWithParag p012C, gClassNamesStr, gClassNamesPos(ownClassID).StringBuffOffs
  fn11BA = ownClassID
End Function

Function fn11CA(p013E As T4938) As Integer
Dim l0140 As Integer
  l0140 = fn124C(p013E, False)
  If l0140 = 0 Then
    gClassNamesCount = gClassNamesCount + 1
    l0140 = gClassNamesCount
    If gClassNamesCount > UBound(gClassNamesPos) Then
      ReDim Preserve gClassNamesPos(l0140 + 10)
    End If
    gClassNamesPos(l0140).ResObjName = p013E
    sub11DD l0140
  End If
  fn11CA = l0140
End Function

Function fn120C(p016E As Integer) As String
  fn120C = ExtractKeyWord(gClassNamesStr, gClassNamesPos(p016E).StringBuffOffs)
End Function

Function fn123C(p00C0 As T4938, p00C2%) As String
Dim l00C4 As Integer
Dim l00C6 As String
  l00C4 = m006C(Asc(p00C0.M3CE1))
  Do While l00C4
    If gClassNamesPos(l00C4).ResObjName.M3CE1 = p00C0.M3CE1 Then Exit Do
    l00C4 = gClassNamesPos(l00C4).Count
  Loop
  If l00C4 Then
    l00C6 = fn12C6(l00C4, p00C2)
  Else
    MsgBox fn1285(p00C0), , "unknown Class"
  End If
  If Len(l00C6) = 0 Then l00C6 = "_" & toHex16(CLng(p00C2))
  fn123C = l00C6
End Function

Function fn124C(p00CE As T4938, p00D0 As Integer) As Integer
Dim l00D2 As Integer
Dim l00D4 As Integer
On Error Resume Next
  l00D4 = Asc(p00CE.M3CE1)
  l00D2 = m006C(l00D4)
  Do While l00D2
    If gClassNamesPos(l00D2).ResObjName.M3CE1 = p00CE.M3CE1 Then
      fn124C = l00D2
      Exit Function
    End If
    If gClassNamesPos(l00D2).Count = l00D2 Then
      DbgMileStone "self-ref in AppClasses"
      sub1260 l00D4
      l00D2 = m006C(l00D4)
    Else
      l00D2 = gClassNamesPos(l00D2).Count
    End If
  Loop
  If p00D0 Then
    If p00CE.M3CE1 <> String$(16, 0) Then
    End If
  End If
End Function

Function Findclass(ByVal ClassName As String) As Integer
Dim gClassNamePos As Integer
Dim i
  gClassNamePos = InStr(gClassNamesStr, ClassName & "§")
  If gClassNamePos = 0 Then Exit Function
  
  For i = 1 To gClassNamesCount
    If gClassNamesPos(i).StringBuffOffs = gClassNamePos Then
      Findclass = i
      Exit Function
    End If
  Next
  
  MsgBox ClassName, , "class not found"
  
End Function

Function fn1285(p00E2 As T4938) As String
Dim l00E4 As String
Dim l00E6
Dim l00EA As T4B14
  LSet l00EA = p00E2
  l00E4 = "(" & toHex16(CLng(l00EA.M3CF5)) & toHex16(CLng(l00EA.M34FD)) & "-" & toHex16(CLng(l00EA.M4B20)) & "-" & toHex16(CLng(l00EA.M3D3B)) & "-"
  For l00E6 = 1 To 8
    l00E4 = l00E4 & fn0842(Asc(Mid$(l00EA.M4B26, l00E6, 1)))
    If l00E6 = 2 Then l00E4 = l00E4 & "-"
  Next
  fn1285 = l00E4 & ")"
End Function

Function fn1291(p00F0 As Integer) As String
  fn1291 = fn1285(gClassNamesPos(p00F0).ResObjName)
End Function

Function fn12C6(p011C As Integer, p011E As Integer) As String
Dim l0120 As Integer
Dim l0122 As Integer
Dim l0124 As String
Dim l0126 As Integer
  If p011C Then
    l0120 = p011E \ 4
    If l0120 > gClassNamesPos(p011C).M4B85 Then
      l0124 = "?"
    Else
      l0126 = gClassNamesPos(p011C).M4B7C + l0120
      l0122 = m001E(l0126).StringBuffOffs
      If l0122 Then
        l0124 = ExtractKeyWord(m0052, l0122)
      End If
    End If
  End If
  fn12C6 = l0124
End Function

Sub load_400ss()
  load_400s
End Sub

Sub sub0D79(p010E As Integer)
Dim l0110
Dim l0112 As Integer
Dim l0114 As Integer
  l0114 = g_decl_objs(p010E).M41E8
  l0112 = gv0D70 + 2
  gv0D70 = l0112 + l0114
  If gv0D70 > UBound(gv0D3E) Then ReDim Preserve gv0D3E(gv0D70)
  For l0110 = -1 To l0114
    gv0D3E(l0112 + l0110).ModuleIndex = p010E
  Next
  g_decl_objs(p010E).M41E0 = l0112
End Sub

Sub sub0EA2(p0188 As Integer, ByVal pv018A As Integer)
Dim l018C As Integer
Dim l018E As Integer
Dim Offset As Integer
Dim l0192 As Long
Dim l0196 As Integer
Dim l019C As String
Dim l019E As String
Dim l01A0 As Integer
Dim l01A2 As Integer
Dim l01A4 As Integer
Dim l01A6 As Integer
' Const mc01A8 = 57 ' &H39%
' Const mc01B2 = 2 ' &H2%
  gReloc_Arr(pv018A).M36B3_count = mc01A8
  l019C = g_decl_objs(p0188).M3F23
  l019E = l019C & ".class"
  If Len(l019C) = 0 Then
    DebugPrintVar False, "form# valid"
  Else
    sub0D9B pv018A, l019E & Format$(l01A2)
  End If
  l018C = g_decl_objs(p0188).M41E0
  If l018C = 0 Then
    If g_decl_objs(p0188).M36B3_count <> mc01B2 Then
      DbgMileStone "init *.CLS!"
    End If
    Exit Sub
  End If
  gv0D3E(l018C).M3F23 = l019C
If gIsVB_Ver Then
  gv0D3E(l018C).M41AD = fn124C(gv3450.M41AD, True)
  l01A4 = gv0D3E(l018C).M41AD
  l0192 = fn202F(pv018A) + &H20
  l0196 = gv3450.M6720 \ 4
Else
  gv0D3E(l018C).M41AD = fn124C(gv2E2E.M41AD, True)
  l01A4 = gv0D3E(l018C).M41AD
  l0192 = fn202F(pv018A) + &H1C
  l0196 = gv2E2E.M6720 \ 4
End If
  l01A6 = m0050
  gClassNamesPos(l01A4).M4B7C = m0050
  gClassNamesPos(l01A4).M4B85 = l0196
  gClassNamesPos(l01A4).M4B68 = gc0CF0
  sub121D l01A4, g_decl_objs(p0188).M41CD
  m0050 = m0050 + l0196
  ReDim Preserve m001E(m0050)
  Offset = fn1631(fn202F(pv018A), False)
  
  If gv161C(Offset).VOffset < l0192 Then
    l01A0 = gv161C(Offset).subRef_M3C6A
    Offset = fn164F(Offset)
  End If
  DbgMileStone "vec.sym->ExeSub"
  Do While Offset
    l018E = (gv161C(Offset).VOffset - l0192) \ 4
    If l018E >= l0196 Then Exit Do
    QuoteWithParag fn2023(gv161C(Offset).subRef_M3C6A), m0052, m001E(l01A6 + l018E).StringBuffOffs
    m001E(l01A6 + l018E).subRef_M3C6A = gv161C(Offset).subRef_M3C6A
    Offset = fn164F(Offset)
  Loop
If gIsVB_Ver Then
  Do While gv3450.M6729
    pv018A = fn14AE(pv018A, &H1C, 0, True)
    Assume pv018A, "sym3"
    Get g_hFileExe, fn1483(pv018A), gv3450
    gReloc_Arr(pv018A).M36B3_count = mc01A8
sub0C11 p0188, pv018A
    l01A2 = gv3450.M5A03
    sub0D9B pv018A, l019E & CStr(l01A2)
    If l01A2 = -1 Then
      If g_decl_objs(p0188).M41D7 Then
      Select Case gv3450.M6720
      Case &H70
        gv0D3E(l018C + l01A2).M41AD = fn11BA(l019C, "Form", gv3450.M41AD)
      Case &H58
        gv0D3E(l018C + l01A2).M41AD = fn11BA(l019C, "MDIForm", gv3450.M41AD)
      Case Else
        DbgMileStone "Not (MDI)Form"
      End Select
      Else
        DbgMileStone "handle class modules"
      End If
      gv0D3E(l018C + l01A2).M3F23 = l019C
    Else
      Select Case gv3450.M4B5E
      Case 3
        gv0D3E(l018C + l01A2).M41AD = fn124C(gv3450.M41AD, True)
      Case Else
        ErrBeep
      End Select
    End If
  Loop
Else
  Do While gv2E2E.M6729
    pv018A = fn14AE(pv018A, &H18, 0, True)
    Assume pv018A, "sym3"
    Get g_hFileExe, fn1483(pv018A), gv2E2E
    gReloc_Arr(pv018A).M36B3_count = mc01A8
sub0C11 p0188, pv018A
    l01A2 = gv2E2E.M5A03
    sub0D9B pv018A, l019E & CStr(l01A2)
    If l01A2 = -1 Then
      If g_decl_objs(p0188).M41D7 Then
      Select Case gv2E2E.M6720
      Case &H70
        gv0D3E(l018C + l01A2).M41AD = fn11BA(l019C, "Form", gv2E2E.M41AD)
      Case &H58
        gv0D3E(l018C + l01A2).M41AD = fn11BA(l019C, "MDIForm", gv2E2E.M41AD)
      Case Else
        DbgMileStone "Not (MDI)Form"
      End Select
      Else
        DbgMileStone "handle class modules"
      End If
      gv0D3E(l018C + l01A2).M3F23 = l019C
    Else
      Select Case gv2E2E.M4B5E
      Case 3
        gv0D3E(l018C + l01A2).M41AD = fn124C(gv2E2E.M41AD, True)
      Case Else
        ErrBeep
      End Select
    End If
  Loop
End If
End Sub

Sub sub11DD(p00F2 As Integer)
Dim l00F6 As Integer
  l00F6 = Asc(gClassNamesPos(p00F2).ResObjName.M3CE1)
  gClassNamesPos(p00F2).Count = m006C(l00F6)
  m006C(l00F6) = p00F2
End Sub

Private Sub load_400(ByVal p0142 As String)
Dim tmp16Bit%
Dim hFile As Integer
Dim l0148 As Integer
Dim l014A%
Dim l014C As Integer
Dim l014E As Integer
Dim l0150 As Integer
Dim l0152 As Integer
Dim l0154 As String
Dim l0158 As T4B33
Dim ext As String
' Const mc0164 = 16384 ' &H4000%
Dim l0166() As Integer
  gFileName_dat = App.Path & p0142
  If gIsVB5 Then
    ext = ".500"
    If Len(Dir$(gFileName_dat & ext)) = 0 Then
      FileCopy gFileName_dat & ".400", gFileName_dat & ".500"
    End If
  Else
    ext = ".400"
  End If
  
  hFile = FreeFile
  Open gFileName_dat & ext For Binary Access Read As hFile
  Get hFile, , tmp16Bit
  Get hFile, , l014A
  If tmp16Bit <> &H400 Then
    Close hFile
    Kill gFileName_dat
    ShowMsgBox mc0164
    Exit Sub
  End If
  gv0CE0 = gv0CE0 + 1
  If gv0CE0 > UBound(gv0CAE) Then ReDim gv0CAE(gv0CE0 + 10)
  gv0CAE(gv0CE0) = Mid$(p0142, 2)
  l014C = gClassNamesCount
  l014E = m0050
  l0150 = Len(m0052)
  ReDim Preserve gClassNamesPos(gClassNamesCount + l014A)
  ReDim l0166(l014A) As Integer
  For tmp16Bit = 1 To l014A
    Get hFile, , l0158
    Select Case l0158.M4B56
    Case gc0CEA, gc0CEC, gc0CE8
      gClassNamesCount = gClassNamesCount + 1
      l0166(tmp16Bit) = gClassNamesCount
      If l0158.M4B85 Then l0158.M4B7C = l0158.M4B7C + l014E
      sub11FF l0158, gClassNamesPos(gClassNamesCount)
      l0148 = Asc(l0158.ResObjName.M3CE1)
      sub11DD gClassNamesCount
      If l0158.M4B72 Then
        If l0166(l0158.M4B72) = 0 Then
          MsgBox l0158.M4B42, , "missing parent"
        Else
          gClassNamesPos(gClassNamesCount).M4B72 = l0166(l0158.M4B72)
        End If
      End If
      If l0158.M4B68 = gc0CF2 Then
        gClassNamesCount = gClassNamesCount + 1
        If gClassNamesCount > UBound(gClassNamesPos) Then ReDim Preserve gClassNamesPos(gClassNamesCount + 10)
        l0148 = l0148 + 1
        Assume l0148 <= mc006A, "hash class"
        Assume m006C(l0148) <> gClassNamesCount, "prev class"
        Mid$(l0158.ResObjName.M3CE1, 1, 1) = Chr$(l0148)
        sub11FF l0158, gClassNamesPos(gClassNamesCount)
        gClassNamesPos(gClassNamesCount).M4B72 = l0166(l0158.M4B72)
        sub11DD gClassNamesCount
      End If
    Case Else
    End Select
  Next
  Get hFile, , l014A
  ReDim Preserve m001E(m0050 + l014A)
  For tmp16Bit = m0050 + 1 To m0050 + l014A
    Get hFile, , l0152
    If l0152 Then m001E(tmp16Bit).StringBuffOffs = l0152 + l0150
  Next
  m0050 = m0050 + l014A
  Get hFile, , tmp16Bit
  l0154 = Space$(tmp16Bit)
  Get hFile, , l0154
  m0052 = m0052 & l0154
  Close hFile
End Sub

Private Sub sub11FF(p0176 As T4B33, p017A As T4BAA)
  QuoteWithParag p0176.M4B42, gClassNamesStr, p017A.StringBuffOffs
  p017A.M4B56 = p0176.M4B56
  p017A.ResObjName = p0176.ResObjName
  p017A.M4B5E = p0176.M4B5E
  p017A.M4B68 = p0176.M4B68
  p017A.M4B72 = p0176.M4B72
  p017A.M4B7C = p0176.M4B7C
  p017A.M4B85 = p0176.M4B85
  p017A.M3F0B = gv0CE0
End Sub

Sub sub121D(p0170 As Integer, p0172 As Integer)
  gClassNamesPos(p0172).M4B68 = gClassNamesPos(p0170).M4B68
  gClassNamesPos(p0172).M4B7C = gClassNamesPos(p0170).M4B7C
  gClassNamesPos(p0172).M4B85 = gClassNamesPos(p0170).M4B85
End Sub

Sub sub122F(p017E As Control, p0180 As Integer)
Dim l0182
Dim l0184 As Integer
Dim l0186 As Integer
  p017E.Clear
  p017E.ListIndex = -1
  p017E.AddItem "GUID: " & fn1285(gClassNamesPos(p0180).ResObjName)
  l0186 = gClassNamesPos(p0180).M4B85
  If l0186 Then
    l0184 = gClassNamesPos(p0180).M4B7C - 1
    For l0182 = 1 To l0186
      l0186 = m001E(l0184 + l0182).StringBuffOffs
      If l0186 Then
        p017E.AddItem ExtractKeyWord(m0052, l0186)
        p017E.ItemData(p017E.NewIndex) = l0184 + l0182
      End If
    Next
  End If
End Sub

Sub sub1260(p0116 As Integer)
Dim l0118
  MsgBox "repair AppDir"
  m006C(p0116) = 0
  For l0118 = 1 To gClassNamesCount
    If Asc(gClassNamesPos(l0118).ResObjName.M3CE1) = p0116 Then
      gClassNamesPos(l0118).Count = m006C(l0118)
      m006C(p0116) = l0118
    End If
  Next
End Sub

Sub load_400s()
  ReDim m006C(mc006A), gClassNamesPos(0)
  ReDim gv0CAE(10): gv0CE0 = 0
  ReDim gv0D3E(10): gv0D70 = 0
  ReDim m001E(0): m0052 = ""
  m0050 = 0
  gClassNamesCount = 0: gClassNamesStr = "§"
  load_400 "\STDOLE"
  load_400 "\VB"
  load_400 "\VEN2232"
  load_400 "\STDTYPE"
  load_400 "\VBEXT"
  load_400 "\DAO2516"
End Sub

Sub sub12B1(p00F8 As Integer, p00FA As Integer, p00FC As String, p00FE As Integer, p0100 As Long)
Dim l0102 As Integer
Dim l0104 As Integer
  l0102 = g_decl_objs(p00F8).M41E0
  DebugPrintVar l0102 > 0, "ctrl found"
  l0104 = l0102 + p00FA
  If l0104 > gv0D70 Then
    gv0D70 = l0104
    If gv0D70 > UBound(gv0D3E) Then ReDim Preserve gv0D3E(gv0D70 + 20)
  End If
  gv0D3E(l0104).ModuleIndex = p00F8
  gv0D3E(l0104).M4C81 = p00FA
  gv0D3E(l0104).M36B3_count = p00FE
  gv0D3E(l0104).M3F23 = p00FC
  gv0D3E(l0104).Flags = p0100
End Sub

