Attribute VB_Name = "MODULE35"
' Module35
Option Explicit
Const mc0016 = "\vb40016.400"
Dim m400_Arr2() As Integer
Dim m400_Size2 As Integer
Dim m005A As Integer
Dim m0068() As Integer
Const mc009A = 1 ' &H1%
Const mc009C = 2 ' &H2%
Const mc009E = 3 ' &H3%
Const mc00A0 = 4 ' &H4%
Const mc00A2 = 5 ' &H5%
Const mc00A4 = 6 ' &H6%
Const mc00A6 = 7 ' &H7%
Const mc00A8 = 8 ' &H8%
Const mc00AA = 9 ' &H9%
Const mc00AC = 10 ' &HA%
Const mc00AE = 11 ' &HB%
Const mc00B0 = 12 ' &HC%
Const mc00B2 = 13 ' &HD%
Const mc00B4 = 14 ' &HE%
Const mc00B6 = 15 ' &HF%
Const mc00B8 = 16 ' &H10%
Dim m00CC() As Integer
Const mc0224 = 4 ' &H4%

Function fn20BF(p0100 As Integer, p0102 As String) As Integer
Dim l0104
Dim l0106 As Integer
  l0106 = InStr(gFrmClass_Str, p0102 & "§")
  If p0100 = &HFF Then
    If l0106 Then
      For l0104 = gFrmProp_Size1 To 1 Step -1
        If gFrmClass_Arr(l0104).Name = l0106 Then Exit For
      Next
    End If
  Else
    If p0100 >= m400_Size2 Then
      m400_Size2 = p0100 + 1
      ReDim Preserve m400_Arr2(p0100)
    End If
    l0104 = m400_Arr2(p0100)
  End If
  If l0104 = 0 Then
    gFrmProp_Size1 = gFrmProp_Size1 + 1
    l0104 = gFrmProp_Size1
    ReDim Preserve gFrmClass_Arr(l0104)
  End If
  If l0106 Then
    gFrmClass_Arr(l0104).Name = l0106
  Else
    QuoteWithParag p0102, gFrmClass_Str, gFrmClass_Arr(l0104).Name
  End If
  gFrmClass_Arr(l0104).ResObjName = p0100
  If p0100 < m400_Size2 Then m400_Arr2(p0100) = l0104
  m005A = True
  fn20BF = l0104
End Function

Function fn20DA(p0112 As Integer, p0114 As Integer, p0116 As String, p0118 As Integer, p011A As Integer, p011C As Long) As Integer
Dim l011E
Dim l0120 As Integer
Dim l0122 As Integer
Dim l0124 As Integer
Dim l0126%
Dim l0128 As Integer
  If p0114 >= gFrmClass_Arr(p0112).M4B85 Then
    l0120 = gFrmClass_Arr(p0112).M4B7C
    l0122 = UBound(m0068) + 1
    ReDim Preserve m0068(UBound(m0068) + p0114 + 1)
    For l011E = 0 To gFrmClass_Arr(p0112).M4B85 - 1
      m0068(l0122 + l011E) = m0068(l0120 + l011E)
    Next
    gFrmClass_Arr(p0112).M4B7C = l0122
    gFrmClass_Arr(p0112).M4B85 = p0114 + 1
  Else
    l0124 = m0068(gFrmClass_Arr(p0112).M4B7C + p0114)
  End If
  QuoteWithParag p0116, gFrmProp_Str2, l0126
  If l0124 Then
    For l011E = 1 To gFrmProp_Size1
      l0120 = gFrmClass_Arr(l011E).M4B7C
      For l0122 = 0 To gFrmClass_Arr(l011E).M4B85 - 1
        If m0068(l0120 + l0122) = l0124 Then
          If l0128 Then l0124 = 0: GoTo L44640
          l0128 = 1
        End If
      Next
    Next
  Else
    For l0124 = gFrmProp_Size3 To 1 Step -1
      If gFrmProp_Arr3(l0124).Name = l0126 Then
        If gFrmProp_Arr3(l0124).Flags = p011A Then
          If gFrmProp_Arr3(l0124).M5FFD = p011C Then
            Exit For
          End If
        End If
      End If
    Next
  End If

L44640:
  If l0124 = 0 Then
    gFrmProp_Size3 = gFrmProp_Size3 + 1
    If gFrmProp_Size3 > UBound(gFrmProp_Arr3) Then ReDim Preserve gFrmProp_Arr3(gFrmProp_Size3 + 10)
    l0124 = gFrmProp_Size3
  End If
  m0068(gFrmClass_Arr(p0112).M4B7C + p0114) = l0124%
  gFrmProp_Arr3(l0124).Name = l0126
  gFrmProp_Arr3(l0124).M36B3_count = p0118
  gFrmProp_Arr3(l0124).Flags = p011A
  gFrmProp_Arr3(l0124).M5FFD = p011C
  m005A = True
  fn20DA = l0124
End Function

Function getFrmClass(Index As Integer) As String
  If gFrmClass_Arr(Index).Name Then
    getFrmClass = ExtractKeyWord(gFrmClass_Str, gFrmClass_Arr(Index).Name)
  Else
    getFrmClass = "ctrl" & Hex$(gFrmClass_Arr(Index).ResObjName)
  End If
End Function

Function fn20F2(p0132 As Integer, p0134 As Integer) As String
Dim l0136 As Integer
Dim l0138 As Integer
  l0136 = m400_Arr2(p0132)
  DebugPrintVar l0136, "ctrl index"
  l0138 = gFrmClass_Arr(l0136).M4B7C
  fn20F2 = GetFrmProp(l0138 + p0134, p0134)
End Function

Function fn2102(p013C As Integer) As String
Dim l013E As Integer
  If p013C <= UBound(m400_Arr2) Then
    l013E = m400_Arr2(p013C)
    fn2102 = "_D" & getFrmClass(l013E)
  Else
    MsgBox "Ctl" & fn085C(p013C), , "unknown ctltyp"
  End If
End Function

Function fn216B(p018E As Integer, p0190 As Integer) As String
  If gFrmProp_Arr4(p018E).Name Then
    fn216B = ExtractKeyWord(gv233E, gFrmProp_Arr4(p018E).Name)
  Else
    fn216B = "evnt" & Format$(p0190)
  End If
End Function

Function fn2178(p0194 As Integer) As String
  If gFrmProp_Arr4(p0194).M388F Then fn2178 = ExtractKeyWord(gv2342, gFrmProp_Arr4(p0194).M388F)
End Function

Function GetFrmProp(p0198 As Integer, p019A As Integer) As String
On Error Resume Next
  If gFrmProp_Arr3(p0198).Name = 0 Then
    GetFrmProp = "prop" & Format$(p019A)
  Else
    GetFrmProp = ExtractKeyWord(gFrmProp_Str2, gFrmProp_Arr3(p0198).Name)
  End If
End Function

Sub sub0CC1()
Dim tmp16Bit As Integer
Dim l022C As Integer
Dim hFile As Integer
Dim l0230 As Integer
Dim l0232 As Integer

If gIsVB_Ver Then
ElseIf g_IsNE Then
Else
  ShowMsgBox mc0224
  Exit Sub
End If
  m005A = False
  hFile = FreeFile
  gFileName_dat = App.Path & mc0016
  Open gFileName_dat For Binary As hFile
  Get hFile, , tmp16Bit
  Get hFile, , gFrmProp_Size1: ReDim gFrmClass_Arr(gFrmProp_Size1)
  Get hFile, , m400_Size2: ReDim m400_Arr2(m400_Size2)
  
  Get hFile, , gFrmProp_Size3: ReDim gFrmProp_Arr3(gFrmProp_Size3)
  Get hFile, , gFrmProp_Size4: ReDim gFrmProp_Arr4(gFrmProp_Size4)
  
  If tmp16Bit = 400 Then
  
    For tmp16Bit = 1 To gFrmProp_Size1
      Get hFile, , gFrmClass_Arr(tmp16Bit)
      
      gFrmClass_Arr(tmp16Bit).M4B7C = l0230 + 1
      gFrmClass_Arr(tmp16Bit).M4B8E = l0232 + 1
      
      l0230 = l0230 + gFrmClass_Arr(tmp16Bit).M4B85
      l0232 = l0232 + gFrmClass_Arr(tmp16Bit).M4B97
      
      l022C = gFrmClass_Arr(tmp16Bit).ResObjName
      
      If l022C <> &HFF Then
        If l022C > m400_Size2 Then
          m400_Size2 = l022C
          ReDim Preserve m400_Arr2(m400_Size2)
        End If
        m400_Arr2(l022C) = tmp16Bit
      End If
    Next
    
    'Read Str1
    Get hFile, , tmp16Bit: gFrmClass_Str = Space$(tmp16Bit): Get hFile, , gFrmClass_Str
    
    For tmp16Bit = 1 To gFrmProp_Size3: Get hFile, , gFrmProp_Arr3(tmp16Bit): Next
    
    'Read Str2
    Get hFile, , tmp16Bit: gFrmProp_Str2 = Space$(tmp16Bit): Get hFile, , gFrmProp_Str2
    
    ReDim m0068(l0230)
    For tmp16Bit = 1 To l0230: Get hFile, , m0068(tmp16Bit): Next
    For tmp16Bit = 1 To gFrmProp_Size4: Get hFile, , gFrmProp_Arr4(tmp16Bit): Next
    Get hFile, , tmp16Bit: gv233E = Space$(tmp16Bit): Get hFile, , gv233E
    Get hFile, , tmp16Bit: gv2342 = Space$(tmp16Bit): Get hFile, , gv2342
    ReDim m00CC(l0232)
    For tmp16Bit = 1 To l0232: Get hFile, , m00CC(tmp16Bit): Next
  Else
    ReDim m0068(0), m00CC(0)
    ShowMsgBox &H4000
  End If
  Close hFile
End Sub

Sub sub0D6B(p019C%, ByVal p019E As Long, ByVal p01A0 As Long)
Dim l01A4 As T60A3
Dim l01A6 As Long
Dim l01AA As Long
Dim l01AC As Long
Dim l01AE As String
Dim l01B0 As Integer
Dim l01B2 As Integer
Dim l01B4 As Integer
Dim l01B6 As Integer
  Get g_hFileExe, p019E + p01A0, l01A4
  l01AC = p019E + p01A0 + 5 + l01A4.M609B - 1
  GoSub L457E8
  Do While Seek(g_hFileExe) < l01AC
    Get g_hFileExe, , gSingleChar: l01B0 = Asc(gSingleChar)
    Select Case l01B0
    Case 1, 3
      GoSub L457E8
    Case 2
    Case 5
      Exit Do
    Case 4
      Exit Do
    Case Else
      ErrBeep
      Exit Do
    End Select
  Loop
  Do While Seek(g_hFileExe) < l01AC
    Select Case l01B0
    Case 2, 5
      GoSub L457E8
    Case 3
    Case Else
      Exit Do
    End Select
    Get g_hFileExe, , gSingleChar: l01B0 = Asc(gSingleChar)
  Loop
Exit Sub

L457E8:
  l01AA = Seek(g_hFileExe)
  Get g_hFileExe, l01AA, l01A6
  Get g_hFileExe, , gFormMainStruc
  l01B4 = Asc(gFormMainStruc.M3CCD)
  l01AE = Space$(gFormMainStruc.NameSize): Get g_hFileExe, , l01AE
  Get g_hFileExe, , gSingleChar
  If Asc(gSingleChar) Or l01AE < "A" Then
    Get g_hFileExe, l01AA + 7, l01B0
    l01AE = Space$(l01B0): Get g_hFileExe, , l01AE
    Get g_hFileExe, , gSingleChar: If Asc(gSingleChar) Then ErrBeep
  End If
  Get g_hFileExe, , gSingleChar: l01B6 = Asc(gSingleChar)
  sub12B1 p019C, l01B4, l01AE, l01B6, l01A6
  If l01B6 > m400_Size2 Then
    GoSub L459E0
    Get g_hFileExe, , l01B0
    Assume l01B0 < 30, "ctrl name"
    l01AE = Space$(l01B0): Get g_hFileExe, , l01AE
    ErrBeep2
  Else
    l01B2 = m400_Arr2(l01B6)
    If l01B2 = 0 Then GoSub L459E0
  End If
  If l01A6 < 0 Then
    l01A6 = l01A6 And &HFFFFFF
  End If
  Seek g_hFileExe, l01AA + l01A6
Return

L459E0:
  Assume False, "Unsupported Custom Control - Aborting"
  Return
End Sub

Sub sub20CE(p010A As Control, p010C As Integer, ByVal p010E As String)
  p010A.AddItem p010E
  p010A.ItemData(p010A.NewIndex) = p010C
End Sub

Sub sub2117(p0142 As Long, p0144 As Long)
Dim l0148 As T60A3
Dim l014C As T6063
Dim l014E As Long
Dim l0150 As Integer
Dim l0152 As Long
Dim l0154 As Long
Dim l0156 As String
If gIsVB5 Then
  Get g_hFileExe, p0142 + p0144, l014C
  l0154 = p0144 + Len(l014C) - 4 + l014C.M609B
  DisOutput Hex$(l014C.M36B3_count)
  l0150 = Asc(l014C.M606F): DisOutput Str$(l0150) & " Ctls ["
  DisOutput Hex$(l014C.M609B) & "] "
Else
  Get g_hFileExe, p0142 + p0144, l0148
  l0154 = p0144 + Len(l0148) - 4 + l0148.M609B
  DisOutput Hex$(l0148.M36B3_count)
  l0150 = Asc(l0148.M606F): DisOutput Str$(l0150) & " Ctls ["
  DisOutput Hex$(l0148.M609B) & "] "
End If
  WriteOutputTxt
  GoSub L44BD4
  Do While p0144 < l0154
    Get g_hFileExe, , gSingleChar: l0150 = Asc(gSingleChar)
    p0144 = p0144 + 1
    DisOutput fn0842(l0150): WriteOutputTxt
    Select Case l0150
    Case 1, 3
      GoSub L44BD4
    Case 2
    Case Else
      Exit Do
    End Select
  Loop
  Do While p0144 < l0154
    Select Case l0150
    Case 2, 5
      GoSub L44BD4
    Case 3
    Case Else
      Exit Do
    End Select
    Get g_hFileExe, , gSingleChar: l0150 = Asc(gSingleChar)
    p0144 = p0144 + 1
    DisOutput fn0842(l0150): WriteOutputTxt
  Loop
Exit Sub

L44BD4:
  p0144 = Seek(g_hFileExe) - p0142
  l0152 = p0144
  Get g_hFileExe, , l014E
  If l014E < 0 Then
    Get g_hFileExe, , gSingleChar
    Get g_hFileExe, , gSingleChar: l0150 = Asc(gSingleChar): p0144 = p0144 + 2
  End If
  Get g_hFileExe, , gFormMainStruc: DisOutput Str$(Asc(gFormMainStruc.M3CCD))
  l0156 = Space$(gFormMainStruc.NameSize): Get g_hFileExe, , l0156: DisOutput " " & l0156
  DisOutput " [" & Hex$(l014E) & "]"
  WriteOutputTxt
  p0144 = l0152
  sub03B9 g_hFileExe, p0142, p0144, p0144 + 8 + Len(l0156)
  Get g_hFileExe, p0142 + p0144, gFormMainStruc
  If Asc(gFormMainStruc.M3CCD) = &HFF Then
    l0156 = Space$(gFormMainStruc.NameSize)
    Get g_hFileExe, , l0156: DisOutput l0156: WriteOutputTxt
    p0144 = p0144 + 4 + gFormMainStruc.NameSize
  End If
  If l014E < 0 Then
    DisOutput "(" & Format$(l0150) & ")"
    l014E = l014E And &HFFFFFF
  End If
  WriteOutputTxt
  sub03B9 g_hFileExe, p0142, p0144, l0152 + l014E
Return
End Sub

Sub sub2123(p0168 As Control)
Dim l016A%
Dim l016C As Integer
  p0168.Clear
  p0168.ListIndex = -1
  For l016A = 0 To m400_Size2 - 1
    l016C = m400_Arr2(l016A)
    If l016C Then
      p0168.AddItem getFrmClass(l016C)
      p0168.ItemData(p0168.NewIndex) = l016C
    End If
  Next
  For l016A = 1 To gFrmProp_Size1
    l016C = gFrmClass_Arr(l016A).ResObjName
    If l016C = &HFF Then
      p0168.AddItem getFrmClass(l016A)
      p0168.ItemData(p0168.NewIndex) = l016A
    End If
  Next
End Sub

Sub sub2134(p0170 As Control)
  p0170.Clear
  sub20CE p0170, 0, "?"
  sub20CE p0170, mc009A, "byte"
  sub20CE p0170, mc009C, "int"
  sub20CE p0170, mc009E, "long"
  sub20CE p0170, mc00A0, "real"
  sub20CE p0170, mc00A2, "color"
  sub20CE p0170, mc00A4, "string"
  sub20CE p0170, mc00A6, "picture"
  sub20CE p0170, mc00A8, "font"
  sub20CE p0170, mc00AA, "xywh form"
  sub20CE p0170, mc00AC, "xywh ctrl"
  sub20CE p0170, mc00AE, "scalemode"
  sub20CE p0170, mc00B0, "object"
  sub20CE p0170, mc00B2, "date"
  sub20CE p0170, mc00B4, "style"
  sub20CE p0170, mc00B6, "t15?"
  sub20CE p0170, mc00B8, "collection"
End Sub

Sub sub2148(p0172 As Integer, p0176 As Control)
Dim l0178%
Dim l017A As Integer
Dim l017C As Integer
  p0176.Clear
  p0176.ListIndex = -1
  l017C = gFrmClass_Arr(p0172).M4B8E
  For l0178 = 0 To gFrmClass_Arr(p0172).M4B97 - 1
    l017A = m00CC(l017C + l0178)
    p0176.AddItem fn0842(l0178) & " " & fn216B(l017A, l0178) & "(" & fn2178(l017A) & ")"
    p0176.ItemData(p0176.NewIndex) = l017A
  Next
End Sub

Sub sub215A(p0180 As Integer, p0184 As Control)
Dim l0186%
Dim l0188 As Integer
Dim l018A As Integer
  p0184.Clear
  p0184.ListIndex = -1
  l018A = gFrmClass_Arr(p0180).M4B7C
  For l0186 = 0 To gFrmClass_Arr(p0180).M4B85 - 1
    l0188 = m0068(l018A + l0186)
    p0184.AddItem fn0842(l0186) & " " & GetFrmProp(l0188, l0186)
    p0184.ItemData(p0184.NewIndex) = l0188
  Next
End Sub

Sub sub2194(p01B8 As Integer)
Dim l01BA As Long
Dim l01BC As Integer
  l01BC = g_decl_objs(p01B8).M41D7
  If l01BC = 0 Then
    MsgBox "no sym for form"
    Exit Sub
  End If
  l01BA = fn202F(l01BC)
  sub0115 l01BA
  sub21A1 gVAToRaw2, fn202F(l01BC), g_decl_objs(p01B8).Name_M411A
End Sub

Sub sub21A1(ByVal p01CA As Long, ByVal p01CC As Long, ByVal p01CE As String)
Dim l01D2 As T60A3
Dim l01D4 As Long
Dim l01D8 As Long
Dim l01DA As Long
Dim l01DC As String
Dim l01DE As Long
Dim l01E0 As Integer
Dim l01E2%
Dim l01E4%
Dim l01E6 As Long
Dim l01E8 As Integer
Dim l01EA As Integer
Dim l01EC As Long
Dim l01F0 As Integer
Dim l01F2 As Integer
If gIsVB5 Then
  DecompFrmDesign p01CA, p01CC, p01CE
  Exit Sub
End If
  l01E0 = InStr(p01CE, ".")
  If l01E0 Then
    p01CE = Left$(p01CE, l01E0) & "frx"
  Else
    p01CE = p01CE & ".frx"
  End If
  DisOutput "VERSION 4.00": WriteOutputTxt
  Get g_hFileExe, p01CA + p01CC, l01D2
  l01DA = p01CA + p01CC + 5 + l01D2.M609B - 1
  GoSub L4604A
  Do While Seek(g_hFileExe) < l01DA
    Get g_hFileExe, , gSingleChar
    l01E0 = Asc(gSingleChar)
    Select Case l01E0
    Case 1
      l01EA = l01EA + 2
      GoSub L4604A
    Case 3
      DisOutput Space$(l01EA) & "End": WriteOutputTxt
      GoSub L4604A
    Case 2
      DisOutput Space$(l01EA) & "End": WriteOutputTxt
      l01EA = l01EA - 2
    Case 5
      Exit Do
    Case 4
      Exit Do
    Case Else
      ErrBeep
      Exit Do
    End Select
  Loop
  If l01E0 = 5 Then
    l01EA = l01EA + 2
    Do While Seek(g_hFileExe) < l01DA
      Select Case l01E0
      Case 2, 5
        GoSub L4604A
        If Left$(l01DC, 2) = "'T" Then
          l01EA = l01EA + 2
        Else
          DisOutput Space$(l01EA) & "End": WriteOutputTxt
        End If
      Case 3
        If l01EA > 0 Then
          l01EA = l01EA - 2
          DisOutput Space$(l01EA) & "End": WriteOutputTxt
        Else
          DbgMileStone "indent zu klein"
          DisOutput "End?"
          WriteOutputTxt
        End If
      Case 4
        DbgMileStone "unexpected menu code 4"
        Exit Do
      Case Else
        DbgMileStone "unexpected menu code " & l01E0
        DisOutput Space$(l01EA) & "End": WriteOutputTxt
        Exit Do
      End Select
      Get g_hFileExe, , gSingleChar: l01E0 = Asc(gSingleChar)
    Loop
    l01EA = l01EA - 2
  End If
  If Seek(g_hFileExe) < l01DA Then
    sub03B9 g_hFileExe, p01CA, p01CC - p01CA, l01DA - p01CA
  End If
  Get g_hFileExe, , gSingleChar: l01E0 = Asc(gSingleChar)
  DisOutput "End": WriteOutputTxt
Exit Sub

L4604A:
  l01D8 = Seek(g_hFileExe)
  Get g_hFileExe, , l01D4
  If l01D4 And &HFF000000 Then
    l01D4 = l01D4 And &HFFFFFF
  End If
  DisOutput Space$(l01EA) & "Begin "
  Get g_hFileExe, , gFormMainStruc
  l01DC = Space$(gFormMainStruc.NameSize): Get g_hFileExe, , l01DC
  Get g_hFileExe, , gSingleChar
  If Asc(gSingleChar) Or l01DC < "A" Then
    Get g_hFileExe, l01D8 + 5, l01E0
    Get g_hFileExe, , l01E4
    l01DC = Space$(l01E4): Get g_hFileExe, , l01DC
    Get g_hFileExe, , gSingleChar: If Asc(gSingleChar) Then ErrBeep
  End If
  Get g_hFileExe, , gSingleChar: l01E8 = Asc(gSingleChar)
  l01EA = l01EA + 2
  If l01E8 > m400_Size2 Then
    DisOutput l01DC
    l01E0 = InStr(gFrmClass_Str, l01DC & "§")
    If l01E0 Then
      For l01E4 = gFrmProp_Size1 To 1 Step -1
        If gFrmClass_Arr(l01E4).Name = l01E0 Then Exit For
      Next
    Else
      l01E4 = 0
    End If
    If l01E8 = &HFF Then
      Get g_hFileExe, , l01E2
      l01DC = Space$(l01E2): Get g_hFileExe, , l01DC
      DisOutput " " & l01DC
      Get g_hFileExe, , gSingleChar: If Asc(gSingleChar) Then ErrBeep
    End If
    WriteOutputTxt
    Get g_hFileExe, , gSingleChar
    If l01E4 Then
      GoSub L46488
    Else
    End If
  Else
    l01E4 = m400_Arr2(l01E8)
    If l01E4 Then
      DisOutput "VB." & getFrmClass(l01E4) & " " & l01DC: WriteOutputTxt
      GoSub L46488
    Else
      DisOutput "VB.ctrl" & toHex16(l01E8) & " " & l01DC: WriteOutputTxt
    End If
  End If
  p01CC = Seek(g_hFileExe)
  If p01CC < l01D8 + l01D4 Then
    Assume l01D8 + l01D4 <= l01DA, "ctl size"
    p01CC = p01CC - p01CA: l01D8 = l01D8 - p01CA
    sub03B9 g_hFileExe, p01CA, p01CC, l01D8 + l01D4
  End If
  l01EA = l01EA - 2
Return

L46488:
  l01F2 = gFrmClass_Arr(l01E4).M4B7C
  Do
    p01CC = Seek(g_hFileExe)
    Get g_hFileExe, , gSingleChar: l01E2 = Asc(gSingleChar)
    If l01E2 = &HFF Then
      If p01CC >= l01D8 + l01D4 - 1 Then
        gDisOutputTxt = "": Exit Do
      End If
    ElseIf l01E2 >= gFrmClass_Arr(l01E4).M4B85 Then
      gDisOutputTxt = ""
      Seek g_hFileExe, p01CC
      Exit Do
    Else
      l01F0 = m0068(l01F2 + l01E2)
      l01DC = GetFrmProp(l01F0, l01E2)
      If gFrmProp_Arr3(l01F0).Flags And gc22EC Then
        DisOutput Space$(l01EA) & l01DC: WriteOutputTxt
        ErrBeep
        Get g_hFileExe, , l01E4
        Get g_hFileExe, , l01E0
        For l01E4 = 1 To l01E4
          Get g_hFileExe, , l01E0
          If l01E0 Then
            If l01E0 < 64 Then
              l01DC = Space$(l01E0)
              Get g_hFileExe, , l01DC
            Else
              l01E6 = ToUnSigned(l01E0)
              GoSub L46D8E
            End If
          Else
            DisOutput " = """""
          End If
          WriteOutputTxt
        Next
      Else
        Select Case gFrmProp_Arr3(l01F0).M36B3_count
        Case mc009A: GoSub L46902
        Case mc009C: GoSub L4699A
        Case mc009E: GoSub L46B18
        Case mc00A2: GoSub L46B8A
        Case mc00A4: GoSub L46C0E
        Case mc00A6: GoSub L46D32
        Case mc00A8: GoSub L46E46
        Case mc00AA
          l01DC = "Left": l01EC = -60: GoSub L46A9E
          l01DC = "Top": l01EC = -345: GoSub L46A9E
          l01DC = "Width": l01EC = 120: GoSub L46A9E
          l01DC = "Height": l01EC = 405: GoSub L46A9E
        Case mc00AC
          l01DC = "Left": GoSub L4699A
          l01DC = "Top": GoSub L4699A
          l01DC = "Width": GoSub L4699A
          l01DC = "Height": GoSub L4699A
        Case mc00AE
          GoSub L4699A
          GoSub L46A0E
        Case mc00B4
          GoSub L46902
          GoSub L46A0E
        Case Else
          Seek g_hFileExe, p01CC
          Exit Do
        End Select
      End If
    End If
  Loop
Return

L46902:
  DisOutput Space$(l01EA) & l01DC
  Get g_hFileExe, , gSingleChar: l01E0 = Asc(gSingleChar)
  If l01E0 = &HFF Then l01E0 = -1
  DisOutput " =" & Str$(l01E0): WriteOutputTxt
Return

L4699A:
  DisOutput Space$(l01EA) & l01DC
  Get g_hFileExe, , l01E0
  DisOutput " =" & Str$(l01E0): WriteOutputTxt
Return

L46A0E:
  Get g_hFileExe, , l01E0
  Return

L46A28:
  DisOutput Space$(l01EA) & l01DC
  Get g_hFileExe, , l01E0
  DisOutput " = &H" & Hex$(l01E0): WriteOutputTxt
Return

L46A9E:
  DisOutput Space$(l01EA) & l01DC
  Get g_hFileExe, , l01E6
  DisOutput " =" & Str$(l01E6 + l01EC): WriteOutputTxt
Return

L46B18:
  DisOutput Space$(l01EA) & l01DC
  Get g_hFileExe, , l01E6
  DisOutput " =" & Str$(l01E6): WriteOutputTxt
Return
L46B8A:
  DisOutput Space$(l01EA) & l01DC
  Get g_hFileExe, , l01E6
  DisOutput " = &H" & Hex$(l01E6) & "&": WriteOutputTxt
Return

L46C0E:
  DisOutput Space$(l01EA) & l01DC
  Get g_hFileExe, , l01E0
  If l01E0 Then
    If l01E0 < 64 Then
      l01DC = Space$(l01E0)
      Get g_hFileExe, , l01DC
      DisOutput " = """ & l01DC & """"
    Else
      l01E6 = ToUnSigned(l01E0)
      GoSub L46D8E
    End If
    Get g_hFileExe, , gSingleChar
    If Asc(gSingleChar) Then ErrBeep
  Else
    DisOutput " = """""
  End If
  WriteOutputTxt
Return

L46D32:
  Get g_hFileExe, , l01E6
  If l01E6 <> -1 Then
    DisOutput Space$(l01EA) & l01DC
    GoSub L46D8E
  End If
Return

L46D8E:
  DebugPrintVar l01DE < &H10000, "frx<64K"
  DisOutput " = """ & p01CE & """:" & toHex16(l01DE): WriteOutputTxt
  Seek g_hFileExe, Seek(g_hFileExe) + l01E6
  l01DE = l01DE + l01E6 + 4
Return

L46E46:
  gDisOutputTxt = Space$(l01EA) & "BeginProperty Font": WriteOutputTxt
  l01EA = l01EA + 2
  GoSub L46A0E
  l01DC = "charset": GoSub L46902
  Get g_hFileExe, , gSingleChar: l01E0 = Asc(gSingleChar)
  DisOutput Space$(l01EA) & "underline =" & Str$((l01E0 And 1) <> 0): WriteOutputTxt
  DisOutput Space$(l01EA) & "italic =" & Str$((l01E0 And 2) <> 0): WriteOutputTxt
  DisOutput Space$(l01EA) & "strikethrough =" & Str$((l01E0 And 4) <> 0): WriteOutputTxt
  l01DC = "weight": GoSub L4699A
  Get g_hFileExe, , l01E6
  DisOutput Space$(l01EA) & "size =" & Str$(l01E6 / 10000)
  WriteOutputTxt
  Get g_hFileExe, , gSingleChar: l01E0 = Asc(gSingleChar)
  l01DC = Space$(l01E0): Get g_hFileExe, , l01DC
  DisOutput Space$(l01EA) & "name = """ & l01DC & """": WriteOutputTxt
  l01EA = l01EA - 2
  DisOutput Space$(l01EA) & "EndProperty": WriteOutputTxt
Return
End Sub

Sub DecompFrmDesign(p01F8 As Long, p01FA As Long, ByVal FormName As String)
Dim l0200 As T6063
Dim longFromFile As Long
Dim oldPos As Long
Dim l0208 As Long
Dim ControlName As String
Dim l020C As Long
Dim tmp As Integer
Dim ControlTypeInt2%
Dim l0212%
Dim l0214 As Long
Dim ControlTypeInt As Integer
Dim indent As Integer
Dim l021A As Long
Dim l021C As Integer
Dim l021E As Integer
' Const mc0224 = 4 ' &H4%
 
 'Gen frx Name
  tmp = InStr(FormName, ".")
  If tmp Then
    FormName = Left$(FormName, tmp) & "frx"
  Else
    FormName = FormName & ".frx"
  End If
  
  DisOutput "VERSION 5.00": WriteOutputTxt
  
  Get g_hFileExe, p01F8 + p01FA, l0200
  
  l0208 = p01F8 + p01FA + Len(l0200) - 4 + l0200.M609B - 1
  GoSub L47672
  Do While Seek(g_hFileExe) < l0208
    Get g_hFileExe, , gSingleChar
    tmp = Asc(gSingleChar)
    Select Case tmp
    Case 1
      indent = indent + 2
      GoSub L47672
    Case 3
      DisOutput Space$(indent) & "End": WriteOutputTxt
      GoSub L47672
    Case 2
      DisOutput Space$(indent) & "End": WriteOutputTxt
      indent = indent - 2
    Case 5
      Exit Do
    Case 4
      Exit Do
    Case Else
      ErrBeep
      Exit Do
    End Select
  Loop
  If tmp = 5 Then
    indent = indent + 2
    Do While Seek(g_hFileExe) < l0208
      Select Case tmp
      Case 2, 5
        GoSub L47672
        If Left$(ControlName, 2) = "'T" Then
          indent = indent + 2
        Else
          DisOutput Space$(indent) & "End": WriteOutputTxt
        End If
      Case 3
        If indent > 0 Then
          indent = indent - 2
          DisOutput Space$(indent) & "End": WriteOutputTxt
        Else
          DbgMileStone "indent zu klein"
          DisOutput "End?"
          WriteOutputTxt
        End If
      Case 4
        DbgMileStone "unexpected menu code 4"
        Exit Do
      Case Else
        DbgMileStone "unexpected menu code " & tmp
        DisOutput Space$(indent) & "End": WriteOutputTxt
        Exit Do
      End Select
      Get g_hFileExe, , gSingleChar: tmp = Asc(gSingleChar)
    Loop
    indent = indent - 2
  End If
  If Seek(g_hFileExe) < l0208 Then
    sub03B9 g_hFileExe, p01F8, p01FA - p01F8, l0208 - p01F8
  End If
  Get g_hFileExe, , gSingleChar: tmp = Asc(gSingleChar)
  DisOutput "End": WriteOutputTxt
Exit Sub

L47672:
  oldPos = Seek(g_hFileExe)
  Get g_hFileExe, , longFromFile
  If longFromFile And &HFF000000 Then
    longFromFile = longFromFile And &HFFFFFF
  End If
  
 'Get FormName
  DisOutput Space$(indent) & "Begin "
  Get g_hFileExe, , gFormMainStruc
  ControlName = Space$(gFormMainStruc.NameSize): Get g_hFileExe, , ControlName
  
  Get g_hFileExe, , gSingleChar
  
  If Asc(gSingleChar) Or ControlName < "A" Then
    Get g_hFileExe, oldPos + 5, tmp
    Get g_hFileExe, , l0212
    ControlName = Space$(l0212): Get g_hFileExe, , ControlName
    Get g_hFileExe, , gSingleChar: If Asc(gSingleChar) Then ErrBeep
  End If
  
  Get g_hFileExe, , gSingleChar: ControlTypeInt = Asc(gSingleChar)
  indent = indent + 2
  
  If ControlTypeInt > m400_Size2 Then
    DisOutput ControlName
    tmp = InStr(gFrmClass_Str, ControlName & "§")
    If tmp Then
      For l0212 = gFrmProp_Size1 To 1 Step -1
        If gFrmClass_Arr(l0212).Name = tmp Then Exit For
      Next
    Else
      l0212 = 0
    End If
    If ControlTypeInt = &HFF Then
      Get g_hFileExe, , ControlTypeInt2
      ControlName = Space$(ControlTypeInt2): Get g_hFileExe, , ControlName
      DisOutput " " & ControlName
      Get g_hFileExe, , gSingleChar: If Asc(gSingleChar) Then ErrBeep
    End If
    WriteOutputTxt
    Get g_hFileExe, , gSingleChar
    If l0212 Then
      GoSub L47AB0
    Else
    End If
  Else
    l0212 = m400_Arr2(ControlTypeInt)
    If l0212 Then
      DisOutput "VB." & getFrmClass(l0212) & " " & ControlName: WriteOutputTxt
      GoSub L47AB0
    Else
      DisOutput "VB.ctrl" & toHex16(ControlTypeInt) & " " & ControlName: WriteOutputTxt
    End If
  End If
  
  p01FA = Seek(g_hFileExe)
  If p01FA < oldPos + longFromFile Then
    Assume oldPos + longFromFile <= l0208, "ctl size"
    p01FA = p01FA - p01F8: oldPos = oldPos - p01F8
    sub03B9 g_hFileExe, p01F8, p01FA, oldPos + longFromFile
  End If
  indent = indent - 2
Return

L47AB0:
  l021E = gFrmClass_Arr(l0212).M4B7C
  Do
    p01FA = Seek(g_hFileExe)
    Get g_hFileExe, , gSingleChar: ControlTypeInt2 = Asc(gSingleChar)
    If ControlTypeInt2 = &HFF Then
      If p01FA >= oldPos + longFromFile - 1 Then
        gDisOutputTxt = "": Exit Do
      End If
    ElseIf ControlTypeInt2 >= gFrmClass_Arr(l0212).M4B85 Then
      gDisOutputTxt = ""
      Seek g_hFileExe, p01FA
      Exit Do
    Else
      If ControlTypeInt2 = 1 Then
        ControlTypeInt2 = 0
      End If
      l021C = m0068(l021E + ControlTypeInt2)
      ControlName = GetFrmProp(l021C, ControlTypeInt2)
      If gFrmProp_Arr3(l021C).Flags And gc22EC Then
        DisOutput Space$(indent) & ControlName: WriteOutputTxt
        ErrBeep
        Get g_hFileExe, , l0212
        Get g_hFileExe, , tmp
        For l0212 = 1 To l0212
          Get g_hFileExe, , tmp
          If tmp Then
            If tmp < 64 Then
              ControlName = Input$(tmp, g_hFileExe)
            Else
              l0214 = ToUnSigned(tmp)
              GoSub L48520
            End If
          Else
            DisOutput " = """""
          End If
          WriteOutputTxt
        Next
      Else
        Select Case gFrmProp_Arr3(l021C).M36B3_count
        Case mc009A: GoSub L47FF8
        Case mc009C: GoSub L48090
        Case mc009E
          If Left$(ControlName, 6) = "Client" Then
            l021A = 1: GoSub L48194
            ControlName = "ClientTop": l021A = 6: GoSub L48194
            ControlName = "ClientWidth": l021A = -2: GoSub L48194
            ControlName = "ClientHeight": l021A = -6: GoSub L48194
          Else
            GoSub L4820E
          End If
        Case mc00A2: GoSub L48280
        Case mc00A4: GoSub L48304
        Case mc00A6: GoSub L484C4
        Case mc00A8: GoSub L485D8
        Case mc00AA
          ControlName = "Left": l021A = -60: GoSub L48194
          ControlName = "Top": l021A = -345: GoSub L48194
          ControlName = "Width": l021A = 120: GoSub L48194
          ControlName = "Height": l021A = 405: GoSub L48194
        Case mc00AC
          ControlName = "Left": GoSub L48090
          ControlName = "Top": GoSub L48090
          ControlName = "Width": GoSub L48090
          ControlName = "Height": GoSub L48090
        Case mc00AE
          GoSub L48090
          GoSub L48104
        Case mc00B4
          GoSub L47FF8
          GoSub L48104
        Case Else
          Seek g_hFileExe, p01FA
          Exit Do
        End Select
      End If
    End If
  Loop
Return

L47FF8:
  DisOutput Space$(indent) & ControlName
  Get g_hFileExe, , gSingleChar: tmp = Asc(gSingleChar)
  If tmp = &HFF Then tmp = -1
  DisOutput " =" & Str$(tmp): WriteOutputTxt
Return

L48090:
  DisOutput Space$(indent) & ControlName
  Get g_hFileExe, , tmp
  DisOutput " =" & Str$(tmp): WriteOutputTxt
Return

L48104:
  Get g_hFileExe, , tmp
  Return

L4811E:
  DisOutput Space$(indent) & ControlName
  Get g_hFileExe, , tmp
  DisOutput " = &H" & Hex$(tmp): WriteOutputTxt
Return

L48194:
  DisOutput Space$(indent) & ControlName
  Get g_hFileExe, , l0214
  DisOutput " =" & Str$(l0214 + l021A): WriteOutputTxt
Return

L4820E:
  DisOutput Space$(indent) & ControlName
  Get g_hFileExe, , l0214
  DisOutput " =" & Str$(l0214): WriteOutputTxt
Return
L48280:
  DisOutput Space$(indent) & ControlName
  Get g_hFileExe, , l0214
  DisOutput " = &H" & Hex$(l0214) & "&": WriteOutputTxt
Return

L48304:
  DisOutput Space$(indent) & ControlName
  Get g_hFileExe, , tmp
  If tmp Then
    If tmp < 64 Then
      ControlName = Input$(tmp, g_hFileExe)
      DisOutput " = """ & ControlName & """"
    Else
      l0214 = ToUnSigned(tmp)
      DebugPrintVar l020C < &H10000, "frx<64K"
      DisOutput " = $""" & FormName & """:" & toHex16(l020C): WriteOutputTxt
      Seek g_hFileExe, Seek(g_hFileExe) + l0214
      l020C = l020C + l0214 + 4
    End If
    Get g_hFileExe, , gSingleChar
    If Asc(gSingleChar) Then ErrBeep
  Else
    DisOutput " = """""
  End If
  WriteOutputTxt
Return

L484C4:
  Get g_hFileExe, , l0214
  If l0214 <> -1 Then
    DisOutput Space$(indent) & ControlName
    GoSub L48520
  End If
Return

L48520:
  DebugPrintVar l020C < &H10000, "frx<64K"
  DisOutput " = """ & FormName & """:" & toHex16(l020C): WriteOutputTxt
  Seek g_hFileExe, Seek(g_hFileExe) + l0214
  l020C = l020C + l0214 + 4
Return

L485D8:
  gDisOutputTxt = Space$(indent) & "BeginProperty Font": WriteOutputTxt
  indent = indent + 2
  GoSub L48104
  ControlName = "charset": GoSub L47FF8
  Get g_hFileExe, , gSingleChar: tmp = Asc(gSingleChar)
  DisOutput Space$(indent) & "underline =" & Str$((tmp And 1) <> 0): WriteOutputTxt
  DisOutput Space$(indent) & "italic =" & Str$((tmp And 2) <> 0): WriteOutputTxt
  DisOutput Space$(indent) & "strikethrough =" & Str$((tmp And 4) <> 0): WriteOutputTxt
  ControlName = "weight": GoSub L48090
  Get g_hFileExe, , l0214
  DisOutput Space$(indent) & "size =" & Str$(l0214 / 10000)
  WriteOutputTxt
  Get g_hFileExe, , gSingleChar: tmp = Asc(gSingleChar)
  ControlName = Input$(tmp, g_hFileExe)
  DisOutput Space$(indent) & "name = """ & ControlName & """": WriteOutputTxt
  indent = indent - 2
  DisOutput Space$(indent) & "EndProperty": WriteOutputTxt
Return
End Sub

Sub sub21BE()
' Const mc023A = 16384 ' &H4000%
End Sub

