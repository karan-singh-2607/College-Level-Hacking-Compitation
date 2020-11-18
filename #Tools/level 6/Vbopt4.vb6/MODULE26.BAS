Attribute VB_Name = "MODULE26"
' Module26
Option Explicit
Dim m0022 As Long
Dim m0026 As Long
Dim m002A As Long
Dim m002E As Integer
Dim m0030 As Long
Dim m0034 As Long
Dim m003A As T66F3
Dim m0042 As Integer
Dim m0044 As Integer
Dim m0048 As T5EAD
Dim m0054 As Integer
Dim m0056 As Integer
Dim arrayTyp As Integer
Dim m005A%
Dim m005E As T42C7
Dim m0064 As Integer
Dim m0068 As T55A6
Dim m007A As T696F
Dim m008A As Integer
Dim m008C As Integer
Dim m008E As Integer
Dim m0090 As Integer
Dim m0092 As Integer
Dim m0118() As T6A40
Dim m014A As Integer
Dim m014C As Integer
Dim m014E As Integer
Dim m0150 As Integer
Dim m015C As Integer
Dim m015E As Integer
Dim m0160%
Dim m0162 As Integer
Const mc016A = -1 ' &HFFFF%
Const mc016C = -2 ' &HFFFE%
Const mc016E = -4 ' &HFFFC%
Const mc0170 = -8 ' &HFFF8%
Const mc0172 = -3 ' &HFFFD%
Const mc0174 = -5 ' &HFFFB%
Const mc0176 = -6 ' &HFFFA%
Const mc0178 = -7 ' &HFFF9%
Const mc017A = -9 ' &HFFF7%
Const mc017C = -10 ' &HFFF6%
Const mc017E = -11 ' &HFFF5%
Const mc01AA = 4096 ' &H1000%
Const mc01AC = 20480 ' &H5000%
Const mc01AE = 28672 ' &H7000%
Const mc01B8 = 12288 ' &H3000%
Const mc01C0 = 64 ' &H40%
Const mc01C2 = 128 ' &H80%
Const mc01C4 = 192 ' &HC0%
Const mc01C6 = -20480 ' &HB000%
Const mc01EA = -8192 ' &HE000%
Const mc01EC = 24576 ' &H6000%
Const mc01F2 = 16384 ' &H4000%
Const mc01F6 = 8192 ' &H2000%
Const mc01FA = -32768 ' &H8000%
Const mc0230 = 32 ' &H20%
Const mc0248 = 11 ' &HB%
Const mc024A = 1 ' &H1%
Const mc024C = 28 ' &H1C%
Const mc024E = 2 ' &H2%
Const mc0250 = 10 ' &HA%
Const mc0252 = 3 ' &H3%
Const mc0254 = 29 ' &H1D%
Const mc0256 = 4 ' &H4%
Const mc0258 = 18 ' &H12%
Const mc025A = 10 ' &HA%
Const mc025C = 2 ' &H2%
Const mc0260 = 5 ' &H5%
Const mc0264 = 32 ' &H20%
Const mc0266 = 16 ' &H10%
Const mc0268 = 9 ' &H9%
Const mc026A = 11 ' &HB%
Const mc0284 = 4 ' &H4%
Const mc02B2 = 2 ' &H2%
Const mc02B4 = 4 ' &H4%
Const mc02B6 = 8 ' &H8%
Const mc02B8 = 11 ' &HB%
Const mc02BA = 9 ' &H9%
Const mc02BC = 50 ' &H32%
Const mc030C = 1 ' &H1%
Const mc030E = 3 ' &H3%
Const mc0310 = 4 ' &H4%
Const mc0312 = 5 ' &H5%
Const mc0314 = 6 ' &H6%
Const mc031E = 7 ' &H7%
Const mc0320 = 8 ' &H8%
Const mc0322 = 9 ' &H9%
Const mc0372 = 32 ' &H20%
Const mc0374 = 33 ' &H21%
Const mc037E = 35 ' &H23%
Const mc0380 = 38 ' &H26%
Const mc0382 = 34 ' &H22%
Const mc03A0 = 2 ' &H2%
Const mc03A2 = 3 ' &H3%
Const mc03A4 = 16 ' &H10%
Const mc03A6 = 1 ' &H1%
Const mc03A8 = 8 ' &H8%
Const mc03AC = 41 ' &H29%
Const mc03B4 = -8192 ' &HE000%
Const mc03B6 = -32768 ' &H8000%
Const mc03BE = -16384 ' &HC000%
Const mc0414 = 4096 ' &H1000%
Const mc0418 = 4096 ' &H1000%
Const mc041A = 64 ' &H40%
Const mc0426 = 16 ' &H10%
Const mc0428 = 32 ' &H20%
Const mc042C = 15 ' &HF%
Const mc047C = 5 ' &H5%
Const mc047E = 3 ' &H3%
Const mc0494 = -832 ' &HFCC0%

Function fn177E(p0498 As Long) As Integer
Dim l049A As Integer
Dim l049C As Integer
Dim l049E As Long
Dim l04A0 As Integer
  l049E = p0498
  l049A = gv30F0(l049E)
  If l049A > 0 Then
    l049C = gEXEdefEntries(l049A).owningScope - gLWhatEver
  Else
    If l049A = 0 Then
        For l049C = g_decl_num To 1 Step -1
          If l049E >= g_decl_objs(l049C).M4150 Then
            If l049E - g_decl_objs(l049C).M4150 < g_decl_objs(l049C).Size Then Exit For
          End If
        Next
        Assume l049C, "gbl scope"
      l04A0 = l049C + gLWhatEver
    End If
    If l049A < 0 Then
      fn177E = l049A
      Exit Function
    End If
    Assume l049C > 0, "mod known"
    l049A = fn18EA(0, mc01AC, l049E - g_decl_objs(l049C).M4126, l04A0, l049C, g_gbl(l04A0).M4950)
  End If
  If l049C <> gv16BA Then
    gEXEdefEntries(l049A).M6877 = 0
  End If
  fn177E = l049A
End Function

Private Function fn17D8(p04C6 As Integer, p04C8 As Integer) As Integer
Dim l04CA As Long
  If p04C6 = 0 Then Exit Function
  m0034 = fn202F(p04C6) + p04C8
  l04CA = VAtoRaw(m0034)
  m002E = p04C6
  m0030 = ToUnSigned(p04C8)
  gv014C = fn1624(m0034)
  If gv014C Then
    gVAModuleEnd = gv161C(gv014C).VOffset
  Else
    gVAModuleEnd = &H7FFFFFFF
  End If
  m0034 = m0034 - l04CA
If gIsVB5 Then
  Get g_hFileExe, l04CA, m0048
  m0044 = m0048.Entries And &HFFF
  m0026 = m0048.M5EBD + l04CA
Else
  Get g_hFileExe, l04CA, m003A
  If gIsVB_Ver Then Get g_hFileExe, , gv0120
  m0044 = m003A.Entries And &HFFF
  m0026 = m003A.M5EBD + l04CA
End If
  m0022 = Seek(g_hFileExe)
  m002A = m0022
  m0042 = 0
  m0056 = -1
  If m0044 > 0 Then fn17D8 = fn1804()
End Function

Function fn1804() As Integer
Dim i%
Dim l02F4 As Integer
Dim l02FE%
  If m0056 = 0 Then Exit Function
  If m0042 >= m0044 Then Exit Function
  If m0022 >= m0026 Then
    MsgBox "Tbl error"
    GoTo L2CC74
  End If
  m0042 = m0042 + 1
  Get g_hFileExe, m0022, m0054
  Get g_hFileExe, , m0056
  fn1804 = m0056 And &H191F
  arrayTyp = 0: m005A = 0: m0092 = 0: m008A = 0
  Select Case m0056 And &H1F
  Case mc024A, mc024E, mc0252, 6
    If m0056 And &H2000 Then Get g_hFileExe, , gv0120
    If m0056 And &H800 Then
      DebugPrintVar m0054 > 0, "para"
    End If
  Case mc0256, mc025A
If gIsVB_Ver Then
    Get g_hFileExe, , gv0120
Else
    Get g_hFileExe, , i
End If
    Get g_hFileExe, , m005E.M42D1
  Case mc0260
If gIsVB_Ver Then
    Get g_hFileExe, , gv0120
Else
    Get g_hFileExe, , i
End If
    Get g_hFileExe, , arrayTyp
    gEXEdef_Arr1Count = gEXEdef_Arr1Count + 1
    If gEXEdef_Arr1Count > UBound(gEXEdef_Arr1) Then ReDim Preserve gEXEdef_Arr1(gEXEdef_Arr1Count + 10)
    m005E.M42D1 = gEXEdef_Arr1Count
    If m0056 And &H100 Then
      Select Case arrayTyp And &H1F
      Case 0 To 3
      Case 9
        If gIsVB_Ver Then Get g_hFileExe, , i
        sub17E7 arrayTyp
      Case Else
        DbgMileStone "unknown array typ " & Hex$(arrayTyp)
        GoTo L2CC74
      End Select
      sub1812 arrayTyp, gEXEdef_Arr1(gEXEdef_Arr1Count).M3503, gEXEdef_Arr1(gEXEdef_Arr1Count).M6894
      m005A = 4
    Else
      Select Case arrayTyp And &H1F
      Case 0 To 3, 11
        Get g_hFileExe, , i
        If gIsVB_Ver Then Get g_hFileExe, , gv0120
      Case 9
        If gIsVB_Ver Then Get g_hFileExe, , i
        sub17E7 arrayTyp
      Case mc025A
        Get g_hFileExe, , i
        Get g_hFileExe, , m0092
        Get g_hFileExe, , i
      Case Else
        MsgBox "unknown fixed array"
        GoTo L2CC74
      End Select
      Get g_hFileExe, , m0068
      m0068.M55C2 = gEXEdef_Arr2Count
      m0068.M6894.M42D1 = m0092
      gEXEdef_Arr1(gEXEdef_Arr1Count) = m0068
      l02F4 = m0068.M34F5
      ReDim Preserve gEXEdef_Arr2(gEXEdef_Arr2Count + l02F4)
        If m0064 = mc01C6 Then
        m005A = m0068.M3503
        For i = 1 To l02F4
          gEXEdef_Arr2Count = gEXEdef_Arr2Count + 1
          Get g_hFileExe, , gEXEdef_Arr2(gEXEdef_Arr2Count)
          m005A = m005A * gEXEdef_Arr2(gEXEdef_Arr2Count).M54E9
        Next
      Else
        m005A = &H10 + l02F4 * 8
        For i = 1 To l02F4
          gEXEdef_Arr2Count = gEXEdef_Arr2Count + 1
          Get g_hFileExe, , gEXEdef_Arr2(gEXEdef_Arr2Count)
        Next
      End If
      sub1812 arrayTyp, l02FE, gEXEdef_Arr1(gEXEdef_Arr1Count).M6894
    End If
  Case mc0268
    If m0056 And &H2000 Then
      Get g_hFileExe, , gv0120
    Else
      Get g_hFileExe, , i
    End If
    sub17E7 m0056
  Case Else
    GoTo L2CC74
  End Select

L2CC5A:
  m0022 = Seek(g_hFileExe)
Exit Function

L2CC74:
  fn1804 = 0
Exit Function
End Function

Function fn1820(ByVal pv0302 As Integer) As String
  
  If pv0302 >= m014A Then
    fn1820 = "<empty>"
    Exit Function
  End If
  
  If pv0302 < 0 Then
    pv0302 = m014C - pv0302
  Else
    pv0302 = m014A - pv0302
  End If
  fn1820 = CStr(m0118(pv0302).Name)
End Function

Function fn1828(ByVal pv0306 As Integer, p0308 As Integer) As String
Dim l030A As Integer
' Const mc030C = 1 ' &H1%
' Const mc030E = 3 ' &H3%
' Const mc0310 = 4 ' &H4%
' Const mc0312 = 5 ' &H5%
' Const mc0314 = 6 ' &H6%
' Const mc031E = 7 ' &H7%
' Const mc0320 = 8 ' &H8%
' Const mc0322 = 9 ' &H9%
  If pv0306 >= m014A Then
    fn1828 = "?"
    Exit Function
  End If
  pv0306 = m014A - pv0306
  If p0308 = 0 Then
    fn1828 = CStr(m0118(pv0306).Name)
    Exit Function
  End If
  Select Case m0118(pv0306).M6A4D
  Case mc016C
    l030A = m0118(pv0306).Name
    Select Case p0308
    Case mc030C
      If l030A And &HFF00 Then
        fn1828 = "?°" & l030A
      Else
        fn1828 = CStr(l030A)
      End If
    Case mc030E
      If l030A = False Then
        fn1828 = "False"
      ElseIf l030A = True Then
        fn1828 = "True"
      Else
        fn1828 = "?²" & l030A
      End If
    Case mc0310
      fn1828 = CStr(l030A)
    Case Else
      fn1828 = "&H" & Hex$(l030A)
    End Select
  Case mc016E
    Select Case p0308
    Case mc0312
      fn1828 = m0118(pv0306).Name
    Case mc0314
      gv078C.M3D0D = m0118(pv0306).Name
      LSet gv0798 = gv078C
      fn1828 = CStr(gv0798.M3D1B)
    Case mc031E, mc0320, mc0322
      gv078C.M3D0D = m0118(pv0306).Name
      LSet gv0798 = gv078C
      gv07E4.M3D62 = gv0798.M3D1B
      GoTo L2D04A
    Case Else
      fn1828 = "&H" & m0118(pv0306).Name
    End Select
  Case mc0170
    gv07DA.M3D1B = m0118(pv0306).Name
L2D04A:
    Select Case p0308
    Case mc0320
      LSet gv07E4 = gv07DA
      fn1828 = CStr(gv07E4.M3D62)
    Case mc0322
      fn1828 = CVDate(gv07DA.M3D1B)
    Case Else
      fn1828 = CStr(gv07DA.M3D1B)
    End Select
  Case mc0172
    fn1828 = m0118(pv0306).Name
  Case Else
    fn1828 = m0118(pv0306).Name
  End Select
End Function

Function fn1831() As String
  If m014A <= 0 Then
    DbgMileStone "too many drops"
    m014A = 0
    m0150 = 0
    fn1831 = "?"
  Else
    While m0118(m014A).M6A4D = mc016A
      m014A = m014A - 1
      m0150 = m0150 - 1
    Wend
    If m0118(m014A).TypeCode_M5F96 Then
      DbgVar m0118(m014A).M6A4D <> mc017C, "no ptr"
      sub1838
    End If
    fn1831 = CStr(m0118(m014A).Name)
    m014A = m014A - 1
  End If
End Function

Function fn1844(p0330 As Integer) As String
  If m014A <= 0 Then
    fn1844 = "?"
    gv34F4 = gc3500
  ElseIf p0330 = 0 Then
    fn1844 = fn1831()
  Else
    fn1844 = fn1828(0, p0330)
    m014A = m014A - 1
  End If
End Function

Function fn187B() As String
Dim l03CA
Dim l03CC As Integer
Dim l03CE As String
  
  For l03CA = 1 To m014A
    l03CC = m0118(l03CA).M6A4D
    Select Case l03CC
    Case mc016A: l03CE = l03CE & "-"
    Case mc0170: l03CE = l03CE & "8"
    Case mc016E: l03CE = l03CE & "4"
    Case mc016C: l03CE = l03CE & "2"
    Case 0: l03CE = l03CE & "?"
    Case mc0172: l03CE = l03CE & "v"
    Case mc017C: l03CE = l03CE & "p"
    Case Is < 0
      DbgMileStone "unknown type stacked"
      l03CE = l03CE & "?"
    Case Else
      l03CE = l03CE & Mid$(gc2A0A, l03CC And &HF, 1)
    End Select
    
  Next
  
  fn187B = l03CE
End Function

Function fn18AB(p03EA&) As String
Dim l03EC As Integer
Dim l03EE%
Dim l03F0 As String
Dim l03F2 As Integer
Dim l03F4 As Integer
Dim scope As Integer
Dim arg_exists As String
Dim l03FA As Integer
Dim l03FC As Integer
Dim l03FE As Integer
Dim l0402 As T6A40
Dim l0404 As String
' Const mc0414 = 4096 ' &H1000%
' Const mc0418 = 4096 ' &H1000%
' Const mc041A = 64 ' &H40%
' Const mc0426 = 16 ' &H10%
' Const mc0428 = 32 ' &H20%
' Const mc042C = 15 ' &HF%
  gDisOutputTxt = ""
  l03EC = fn18CE(p03EA)
  l03F0 = gDatArr2(l03EC)
  If Len(l03F0) = 0 Then
    fn18AB = "<undefined>"
    ErrBeep2
    gv34F4 = gc3504
    Exit Function
  End If

L2EF62:
  For l03EE = 1 To Len(l03F0)
    If l03EE > Len(l03F0) Then Exit For
    Select Case Mid$(l03F0, l03EE, 1)
    Case "/"
    Case " "
      sub07B7 ""
    Case "'"
      l03EE = l03EE + 1
      l03F4 = InStr(l03EE + 1, l03F0, "'")
      If l03F4 = 0 Then l03F4 = Len(l03F0) + 1
      sub07B7 Mid$(l03F0, l03EE, l03F4 - l03EE)
      l03EE = l03F4
    Case ">"
      sub1854 m014C + 1, 0
      m014C = m014C + 1
      If fn18DE(Mid$(l03F0, l03EE + 1, 1)) Then l03EE = l03EE + 1
    Case "r"
      l03FE = True
    Case "C"
      l03EE = l03EE + 1
      scope = InStr(gc2A0A, Mid$(l03F0, l03EE, 1))
      l03EE = l03EE + 1
      l03F2 = Val(Mid$(l03F0, l03EE))
      sub184C l03F2, scope
      If l03F2 = 10 Then l03EE = l03EE + 1
    Case "c"
      l03EE = l03EE + 1
      l03F2 = Val(Mid$(l03F0, l03EE, 1))
      DebugPrintVar l03F2, "const<>0"
      scope = m0118(m014C + l03F2).M6A4D
      sub1854 m014C + l03F2, scope
    Case "v"
      l03EE = l03EE + 1
      l03F2 = Val(Mid$(l03F0, l03EE, 1))
      sub1854 m014C + l03F2, 0
      sub1838
      If fn18DE(Mid$(l03F0, l03EE + 1, 1)) Then l03EE = l03EE + 1
    Case "0" To "9"
      l03F2 = Val(Mid$(l03F0, l03EE, 1))
      scope = InStr(gc2A0A, Mid$(l03F0, l03EE + 1, 1))
      If scope Then
        l03EE = l03EE + 1
        DisOutput fn1828(l03F2, scope)
      Else
        DisOutput fn1820(l03F2)
      End If
    Case "<"
      If fn18DE(Mid$(l03F0, l03EE + 1, 1)) Then l03EE = l03EE + 1
      DisOutput fn1831()
    Case "-"
      l03EE = l03EE + 1
      l03F2 = Val(Mid$(l03F0, l03EE, 1))
      sub1886 l03F2
    Case "\"
      l03EE = l03EE + 1
      scope = fn18DE(Mid$(l03F0, l03EE, 1))
      l03EE = l03EE + 1
      gSingleChar = Mid$(l03F0, l03EE, 1)
      scope = InStr(gc2A0A, gSingleChar)
      If scope = 0 Then
        If gSingleChar = "8" Then scope = mc031E
      End If
      If scope Then
        l0404 = fn1831()
        arg_exists = Trim$(Mid$(gc2A1A, scope * 4, 4))
        If Len(arg_exists) Then
          l0404 = "C" & arg_exists & "(" & l0404 & ")"
        Else
          DbgVar Len(arg_exists), "arg exists"
        End If
        sub184C l0404, scope
      Else
        Select Case gSingleChar
        Case "8"
          Select Case m0118(m014A).M6A4D
          Case mc031E, mc0320, mc0322, mc0170
          Case Else
            scope = fn18DE("#")
            m0118(m014A).M6A4D = mc0170
          End Select
        End Select
      End If
    Case "="
L2F6FE:
      sub1854 m014C + 1, 0
      GoTo L2F73C
    Case "~"
L2F73C:
      l03FC = fn18DE(Mid$(l03F0, l03EE + 1, 1))
      If l03FC Then l03EE = l03EE + 1
      arg_exists = fn1844(l03FC)
      If m0118(m014A + 1).M6A4D = mc0172 Then
        scope = m0118(m014A + 1).TypeCode_M5F96
      Else
        scope = 0
      End If
      l0404 = fn1844(l03FC)
      If m0118(m014A + 1).M6A4D = mc0172 Then
        If m0118(m014A + 1).TypeCode_M5F96 = m0162 Then
          m0162 = 0
          m015E = True
        End If
      End If
      If scope > 0 Then
        If gEXEdefEntries(scope).M689A < 0 Then
          If (gEXEdefEntries(scope).M6894.M36B3_count And mc0414) = 0 Then
            If Left$(arg_exists, 1) <> "t" Then
              DbgMileStone fn1B19(scope) & " " & arg_exists & " temp?"
            End If
            m015E = False
          End If
        End If
      End If
      If scope > 0 And (m0162 Or m014A - m0150 > 0 Or Not m015E) Then
        DbgVar scope, "expect scope"
        sub18B7 mc0418 Or mc041A, l0404, scope
        gDisOutputTxt = ""
        sub184C l0404, mc024C
        m0118(m014A).TypeCode_M5F96 = scope
        m015E = False
      Else
        If Left$(l0404, 1) = "'" Then
          ErrBeep
          arg_exists = "'" & arg_exists & "=" & l0404
        Else
          arg_exists = arg_exists & " = " & l0404
        End If
        DisOutput arg_exists
        If scope > 0 Then
          If gEXEdefEntries(scope).M6894.M36B3_count And mc0414 Then
            sub18B7 mc041A, l0404, scope
          Else
            ErrBeep
            sub18B7 mc0418 Or mc041A, l0404, scope
          End If
        Else
          sub18B7 0, arg_exists, 0
        End If
        m015E = True
        GoTo L31B38
      End If
    Case "_"
      arg_exists = " "
      m015E = True
      GoTo L31B6A
    Case "}"
      If p03EA >= gModuleVAEnd Then
        DisOutput "End " & Mid$(l03F0, l03EE + 1)
        gv34F4 = gc34FE
      Else
        DisOutput "Exit " & Mid$(l03F0, l03EE + 1)
      End If
      sub18B7 0, gDisOutputTxt, 0
      m015E = True
      l03EE = Len(l03F0)
    Case "F", "S"
      scope = m0118(m014C + 1).M6A4D
      l03F2 = m0118(m014C + 1).TypeCode_M5F96
      DisOutput fn1820(-1)
      If l03F2 Then
        Select Case scope
        Case mc017C
          If l03F2 > 0 Then
            l03F2 = gReloc_Arr(l03F2).subRef_M3C6A
            While l03F2 < 0
              l03F2 = gReloc_Arr(-l03F2).subRef_M3C6A
            Wend
            Assume l03F2 > 0, "sym ref"
            l03F2 = -gEXEdef_Arr4_imports(l03F2).TypeCode_M5F96
          Else
            DbgMileStone "unnamed symbol?"
          End If
          If l03F2 > 0 Then
            Assume l03F2 <= gEXEdef_Arr5Count, "sub nr"
            If gEXEdef_Subs(l03F2).ModuleIndex Then
              scope = gL2Max + l03F2
              scope = g_gbl(scope).M4950
              If gEXEdef_Subs(l03F2).M4277 And (mc0426 Or mc0428) Then
                l03F0 = "F=": l03EE = 1
              End If
              l03F2 = 0
              While scope
                l03F2 = l03F2 + 1
                scope = gEXEdefEntries(scope).Count
              Wend
            Else
              DbgMileStone "Declare?"
              l03F2 = m014A - m0150
            End If
          ElseIf l03F2 Then
            If Left$(gDisOutputTxt, 3) = "rtc" Or Left$(gDisOutputTxt, 4) = "_RTC" Then
              sub18C3 l03F0, l03EE
              l03F2 = InStr("-0123456789", Right$(l03F0, 1))
              If l03F2 Then
                l03F2 = l03F2 - 2
              ElseIf Left$(l03F0, 1) = "S" Then
                l03F2 = -1
              Else
                l03F2 = Len(l03F0) - 2
              End If
              If l03F2 < 0 Then
                l03F2 = m014A - m0150
              End If
            Else
              l03F2 = m014A - m0150
            End If
          End If
        Case mc0178
            l03F2 = Len(l03F0) - l03EE - 1
          If l03F2 >= 0 Then
            If l03F2 > m014A - m0150 - 1 Then
              l03F2 = m014A - m0150 - 1
            End If
          Else
            DbgMileStone "argument count of prop/meth?"
            l03F2 = m014A - m0150
          End If
        Case Else
          l03F2 = m014A - m0150
        End Select
      Else
        l03F2 = m014A - m0150
      End If
      If Mid$(l03F0, l03EE, 1) = "F" Then
        m015E = False
        gSingleChar = Mid$(l03F0, l03EE + 1, 1)
        l03FC = InStr(gc2A0A, gSingleChar)
If gIsVB_Ver Then
        l03F4 = InStr("=~", gSingleChar)
Else
        l03F4 = InStr("=~#", gSingleChar)
End If
        If l03F4 Then
          l03EE = l03EE + 1
          l0402 = m0118(m014A)
          sub1886 1
          If l03FC = 0 Then l03FC = mc0250
        ElseIf l03FC Then
          l03EE = l03EE + 1
          If m0118(m014C + 1).M6A4D <> mc017C Then
            DbgMileStone "stktyp: " & Hex$(m0118(m014C + 1).M6A4D)
            m0118(m014C + 1).TypeCode_M5F96 = 0
          End If
        End If
      ElseIf Mid$(l03F0, l03EE, 1) <> "S" Then
        GoTo L2EF62
      End If
      If m015E Then
        sub07B7 ""
      Else
        If l03F2 >= 0 Then
          DisOutput "("
        End If
      End If
      arg_exists = ""
      If l03FE Then
        While l03F2 > 0
          arg_exists = ", " & fn1831() & arg_exists
          l03F2 = l03F2 - 1
        Wend
      Else
        While l03F2 > 0
          arg_exists = arg_exists & ", " & fn1831()
          l03F2 = l03F2 - 1
        Wend
      End If
      l03EE = Len(l03F0)
      DisOutput Mid$(arg_exists, 3)
      If m015E Then
        sub18B7 0, gDisOutputTxt, 0
      Else
        If l03F2 = 0 Then DisOutput ")"
        sub184C gDisOutputTxt, l03FC
        gDisOutputTxt = ""
        If l03FC Then
          If l03FC = mc0250 Then
            m014A = m014A + 1
            DbgVar l0402.M6A4D = mc0172, "scope ref"
            m0118(m014A) = l0402
            m015E = False
            GoTo L2F73C
          ElseIf l03FC = mc031E Then
If gIsVB_Ver Then
Else
            ErrBeep
            m014A = m014A + 1
            DbgVar l0402.M6A4D = mc0172, "scope ref"
            m0118(m014A) = l0402
            m015E = False
            GoTo L2F73C
End If
          End If
        End If
      End If
    Case "("
      l03EE = l03EE + 1
      l03FC = InStr(gc2A0A, Mid$(l03F0, l03EE, 1))
      If l03FC Then l03EE = l03EE + 1
      l03F4 = FindLastOccurenceOf(l03F0, "(")
      DisOutput Mid$(l03F0, l03EE, l03F4 - l03EE + 1)
      l03EE = l03F4 + 1
      arg_exists = ""
      l03F2 = InStr("0123456789", Mid$(l03F0, l03EE, 1))
      If l03F2 Then
        If l03FE Then
          While l03F2 > 1
            arg_exists = ", " & fn1831() & arg_exists
            l03F2 = l03F2 - 1
          Wend
        Else
          While l03F2 > 1
            arg_exists = arg_exists & ", " & fn1831()
            l03F2 = l03F2 - 1
          Wend
        End If
      Else
        If l03FE Then
          For l03F2 = l03EE To Len(l03F0)
            scope = InStr(gc2A0A, Mid$(l03F0, l03F2, 1))
            arg_exists = ", " & fn1844(scope) & arg_exists
          Next
        Else
          For l03F2 = l03EE To Len(l03F0)
            scope = InStr(gc2A0A, Mid$(l03F0, l03F2, 1))
            arg_exists = arg_exists & ", " & fn1844(scope)
          Next
        End If
        l03EE = l03F2
      End If
      DisOutput Mid$(arg_exists, 3) & ")"
    Case "a"
      sub184C fn1831() & "()", mc024C
    Case "{"
      l03EE = l03EE + 1
      l03F2 = 1
      If l03F2 < m014A - m0150 Then
        sub1886 l03F2
        l03EE = Len(l03F0)
        m015E = False
      Else
        m015E = True
        DisOutput Mid$(l03F0, l03EE + 1)
        arg_exists = " "
        GoTo L31B6A
      End If
    Case "["
      If l03EE = Len(l03F0) Then
        If l03FE Then
          DisOutput CStr(m0118(1).Name)
          m0118(1).M6A4D = mc016A
        Else
          DisOutput fn1831()
        End If
      Else
        DisOutput Mid$(l03F0, l03EE + 1)
      End If
      arg_exists = " "
      GoTo L31B6A
    Case "i"
      l03EE = l03EE + 1
      scope = fn18DE(Mid$(l03F0, l03EE, 1))
      l03EE = l03EE + 1
      arg_exists = fn1831()
      If scope Then
        If m0118(m014A).M6A4D = mc0172 Then
          l03FA = m0118(m014A).TypeCode_M5F96
          sub10BA scope, gEXEdefEntries(l03FA).M6894.M36B3_count, mc042C
        End If
      End If
      DisOutput fn1831() & " " & Mid$(l03F0, l03EE) & " " & arg_exists
      Exit For
    Case "p"
      l03EE = l03EE + 1
      If fn18DE(Mid$(l03F0, l03EE, 1)) Then l03EE = l03EE + 1
      l03F2 = InStr(l03EE + 1, l03F0, "'")
      If l03F2 = 0 Then l03F2 = Len(l03F0) + 1
      sub184C Mid$(l03F0, l03EE, l03F2 - l03EE) & fn1831(), mc024C
      l03EE = l03F2
    Case "+"
      sub184C fn1844(mc0254) & Mid$(l03F0, l03EE + 1), mc024C
      l03EE = Len(l03F0)
    Case "A"
      DisOutput fn1831()
      l03F2 = m0118(m014C + 1).Name
      arg_exists = "("
      While l03F2 > 0
        DisOutput arg_exists & fn1831()
        l03F2 = l03F2 - 1
        arg_exists = ", "
      Wend
      DisOutput ")"
    Case "J"
      arg_exists = " " & fn1820(-1)
      l03EE = l03EE + 1
      Select Case Mid$(l03F0, l03EE, 1)
      Case "t"
        l0404 = fn1831()
        sub18B7 mc03A0, l0404 & arg_exists, m0160
        DisOutput "if " & l0404 & " goto"
      Case "f"
        l0404 = fn1831()
        sub18B7 mc03A2, l0404 & arg_exists, m0160
        DisOutput "if not " & l0404 & " goto"
      Case "j"
        sub18B7 mc03A6, "GoTo" & arg_exists, m0160
        DisOutput "GoTo"
      Case Else
        sub07B7 Mid$(l03F0, l03EE)
      End Select
      DisOutput arg_exists
      GoTo L31B38
    Case "V"
      l03EE = l03EE + 1
      scope = InStr(gc2A0A, Mid$(l03F0, l03EE, 1))
      DbgMileStone "type para byval"
    Case "."
      l03EE = l03EE + 1
      l03F2 = Val(Mid$(l03F0, l03EE, 1))
      sub184C fn1831() & "." & fn1820(-l03F2), mc024C
      GoTo L2F6FE
    Case "X"
      l03EE = l03EE + 1
      Select Case Mid$(l03F0, l03EE, 1)
      Case "B"
        arg_exists = fn1831()
        l03EE = l03EE + 1
        DisOutput Mid$(l03F0, l03EE, 1) & "Bound(" & fn1831()
        If Val(arg_exists) <> 1 Then
          DisOutput ", " & arg_exists
        End If
        DisOutput ")"
      Case "L"
        l03F2 = m0118(m014E).Name
        If l03F2 And 1 Then DisOutput "Un"
        DisOutput "Lock"
        If l03F2 And 4 Then
          arg_exists = fn1831()
        Else
          arg_exists = fn1831()
          arg_exists = fn1831() & " To " & arg_exists
          sub184C arg_exists, mc024C
        End If
        l03FE = True
        GoTo L31B6A
      Case "N"
        arg_exists = fn1831()
        arg_exists = "Name " & fn1831() & " As " & arg_exists
        DisOutput arg_exists
        GoTo L31B6A
      Case "O"
        l03F2 = m0118(m014E).Name
        arg_exists = fn1831()
        If arg_exists <> "-1" Then
          arg_exists = " Len = " & arg_exists
        Else
          arg_exists = ""
        End If
        arg_exists = " As #" & fn1831() & arg_exists
        Select Case l03F2 And &H3000
        Case &H1000: arg_exists = " Lock Read Write" & arg_exists
        Case &H2000: arg_exists = " Lock Write" & arg_exists
        Case &H3000: arg_exists = " Lock Read" & arg_exists
        End Select
        If l03F2 And &H200 Then arg_exists = "Write " & arg_exists
        If l03F2 And &H100 Then arg_exists = "Read " & arg_exists
        If l03F2 And &H300 Then arg_exists = " Access " & arg_exists
        If l03F2 And &H4000 Then arg_exists = " Shared" & arg_exists
        If l03F2 And &H8CC0 Then
          arg_exists = "? " & Hex$(l03F2 And &HFCC0) & arg_exists
        End If
        Select Case l03F2 And &H3F
        Case 1: arg_exists = "Input" & arg_exists
        Case 2: arg_exists = "Output" & arg_exists
        Case 4: arg_exists = "Random" & arg_exists
        Case 8: arg_exists = "Append" & arg_exists
        Case &H20: arg_exists = "Binary" & arg_exists
        Case Else
          ErrBeep
          arg_exists = Hex$(l03F2) & arg_exists
        End Select
        arg_exists = "Open " & fn1831() & " For " & arg_exists
        DisOutput arg_exists
        GoTo L31B38
      Case "P"
        If IsNumeric(m0118(m014C + 3).Name) Then
          l03F2 = m0118(m014C + 3).Name
        Else
          l03F2 = 0
        End If
        l03F0 = "(~" & fn1831() & "." & m0118(m014C + 2).Name & "(" & Format(l03F2)
        GoTo L2EF62
      Case "p"
        If IsNumeric(m0118(m014C + 1).Name) Then
          l03F2 = (m0118(m014C + 1).Name - 4 + 2) \ 4
        Else
          l03F2 = 0
        End If
        l03F0 = "(~" & m0118(m014C + 2).Name & "(" & Format(l03F2)
        GoTo L2EF62
      Case "c"
        arg_exists = fn1831() & "." & fn1820(-1)
        sub184C arg_exists, mc0174
      Case "R"
        l03F2 = m014A
        DisOutput Mid$(l03F0, l03EE) & " " & fn1831()
        arg_exists = ")"
        While l03F2 > 1
          l03F2 = l03F2 - 2
          arg_exists = " To " & fn1831() & arg_exists
          arg_exists = fn1831() & arg_exists
          If l03F2 > 1 Then arg_exists = ", " & arg_exists
        Wend
        DisOutput "(" & arg_exists
        GoTo L31B38
      Case Else
        fn18AB = " ?"
        gv34F4 = gc3504
        ErrBeep
        Exit Function
      End Select
    Case Else
        fn18AB = " ?"
      gv34F4 = gc3504
      ErrBeep
      Exit Function
    End Select
  Next
  If m015E Then GoTo L31B38
  If Len(gDisOutputTxt) Then
    If l03FC = 0 Then l03FC = mc024C
    sub184C gDisOutputTxt, l03FC
    gDisOutputTxt = ""
  End If
Exit Function

L31B38:
  If m014A > m0150 Then DisOutput "???"
L31B6A:
  arg_exists = " "
  While m014A > m0150
    If True Then
      If l03FE Then
        l03F0 = fn1831()
        If InStr(",;", Right$(l03F0, 1)) Then
          arg_exists = ", " & l03F0 & Mid$(arg_exists, 2)
        ElseIf Left$(l03F0, 1) = ")" Then
          arg_exists = l03F0 & arg_exists
        Else
          arg_exists = ", " & l03F0 & arg_exists
        End If
      Else
        arg_exists = arg_exists & ", " & fn1831()
      End If
    Else
      sub1886 1
      m014A = m014A - 1
    End If
  Wend
  If Len(arg_exists) Then sub07B7 Mid$(arg_exists, 3)
  If Len(gDisOutputTxt) Then
    sub18B7 0, gDisOutputTxt, 0
    fn18AB = gDisOutputTxt
    gDisOutputTxt = ""
  Else
  End If
End Function

Function fn18CE(p0436 As Long) As Integer
Dim l0438 As Integer
Dim l043A As Integer
Dim str_len%
Dim l043E%
Dim l0442 As Integer
Dim l0444%
Dim l0446 As Long
Dim l0448 As String
Dim l044A As Long
Dim ref As Integer
Dim l044E As Integer
Dim l0450 As Integer
Dim l0452 As Integer
Dim signed%
Dim l0456 As Long
Dim l045C As Integer
Dim l046A
Dim l0474 As Long
' Const mc047C = 5 ' &H5%
' Const mc047E = 3 ' &H3%
  DebugPrintVar gVAModuleEnd > p0436, "pat missed"
  m015E = False
  gVACmdStart = p0436
  Get g_hFileExe, gVAToRaw2 + p0436, gv34F6: p0436 = p0436 + 2
  l043A = gDatArrLinks(gv34F6)
  fn18CE = l043A
  l0438 = gDatArr1(l043A).M3818
  m014C = m014A
  If l0438 Then
    Do
      Select Case Mid$(gDatStrings1, l0438, 1)
      Case "§"
        Exit Do
      Case ":"
        GoSub L33A82
        Exit Do
      Case "_", "="
        m015E = True
        Exit Do
      Case "1"
        GoSub L33A82
        m015E = True
        Exit Do
      Case "-"
        Exit Do
      Case "="
        Exit Do
      Case "ß"
      Case "~"
        Get g_hFileExe, , str_len: p0436 = p0436 + 2
        l044A = p0436 + str_len
      Case "l"
        GoSub L3314E
        sub184C l0448, mc0172
        m0118(m014A).TypeCode_M5F96 = ref
      Case "i"
        GoSub L3314E
        sub107E ref
        sub184C l0448, mc0172
        m0118(m014A).TypeCode_M5F96 = ref
      Case "L"
        Assume l044A > p0436, "next token known"
        While p0436 < l044A
          GoSub L3314E
          sub178C ref
        Wend
        Exit Do
      Case "M"
        GoSub L3327E
        sub184C l0448, mc0172
        m0118(m014A).TypeCode_M5F96 = ref
      Case "G"
        GoSub L331F2
        sub184C l0448, mc0172
        m0118(m014A).TypeCode_M5F96 = ref
      Case "F"
        DbgMileStone "assume unused!"
        GoSub L331F2
        If ref >= 0 Then
          DbgMileStone "no Form accessed!"
        End If
        sub184C l0448, mc0172
        m0118(m014A).TypeCode_M5F96 = ref
      Case "s"
        l0442 = mc024C
        GoSub L3364E
        If ref Then l0442 = mc0172
        sub184C l0448, l0442
        m0118(m014A).TypeCode_M5F96 = ref
      Case "c"
        Get g_hFileExe, , l043E: p0436 = p0436 + 2
        sub184C l043E, mc0174
      Case "m"
        Get g_hFileExe, , l043E: p0436 = p0436 + 2
        sub184C l043E, mc0176
      Case "j", "t", "f", "N"
        Get g_hFileExe, , l0444: p0436 = p0436 + 2
        l044A = p0436
        GoSub L33838
        sub184C l0448, mc017A
      Case "e"
        Get g_hFileExe, , l0444: p0436 = p0436 + 2
        l044A = p0436
        If l0444 = -1 Then
          l0448 = "Resume Next"
        ElseIf l0444 = -2 Then
          l0448 = "GoTo 0"
        Else
          l0448 = "GoTo "
          GoSub L33838
        End If
        sub184C l0448, mc017A
      Case "D"
        Assume l044A > p0436, "next token known"
        l0448 = ""
        While p0436 < l044A
          l0448 = l0448 & ", "
          Get g_hFileExe, , l0444: p0436 = p0436 + 2
          GoSub L33838
        Wend
        sub184C Mid$(l0448, 3), mc024C
      Case "P"
        GoSub L338D2
        sub184C l0448, l0442
        m0118(m014A).TypeCode_M5F96 = l044E
      Case "p"
        GoSub L338D2
        l0438 = l0438 + 1
        If l044E Then
          Select Case Mid$(gDatStrings1, l0438, 1)
          Case "s"
          Case "L"
          Case "z"
If gIsVB_Ver Then
            l0448 = fn1561(fn202F(l044E))
Else
            l0448 = fn0218(fn202F(l044E))
End If
          Case "x", "p"
          Case "d"
          Case "c"
          Case Else
            gv34F4 = gc3500
          End Select
        End If
        sub184C l0448, l0442
        m0118(m014A).TypeCode_M5F96 = l044E
      Case "."
        Get g_hFileExe, , l043E: p0436 = p0436 + 2
        If Mid$(gDatStrings1, l0438, 3) = ".pc" Then
          GoSub L338D2
          If l044E = 0 Then
            sub184C l0448 & "." & toHex16(l043E), mc024C
          Else
            l0450 = fn2047(l044E)
            If l0450 Then
              l045C = gEXEdef_Arr4_imports(l0450).TypeCode_M5F96
            End If
            l0448 = "." & fn12C6(l045C, l043E)
            sub184C l0448, mc0178
            m0118(m014A).TypeCode_M5F96 = gClassNamesPos(l045C).M4B7C + (l043E \ 4)
            Seek g_hFileExe, p0436 + gVAToRaw2
          End If
          l0438 = l0438 + 2
        Else
          sub184C l043E, mc0178
        End If
      Case "%"
        Get g_hFileExe, , l043E: p0436 = p0436 + 2
        sub184C l043E, mc0310
      Case "&"
        Get g_hFileExe, , gv078C.M3D0D: p0436 = p0436 + 4
        sub184C gv078C.M3D0D, mc0312
      Case "!"
        Get g_hFileExe, , gv0798.M3D1B: p0436 = p0436 + 4
        sub184C gv0798.M3D1B, mc031E
      Case "#"
        Get g_hFileExe, , gv07DA.M3D1B: p0436 = p0436 + 8
        sub184C gv07DA.M3D1B, mc031E
      Case "@"
        Get g_hFileExe, , gv07E4.M3D62: p0436 = p0436 + 8
        sub184C gv07E4.M3D62, mc0320
      Case "u"
        Get g_hFileExe, , gv0120
        If gv0120 Then
          gv0120 = gv0120 \ 2
          gv0138 = Space$(gv0120)
          For l046A = 1 To gv0120
            Get g_hFileExe, , str_len
            DbgVar str_len, "str len"
            Mid$(gv0138, l046A, 1) = Chr$(str_len)
          Next
        Else
          gv0138 = ""
        End If
        GoTo L32E24
      Case "$"
        Get g_hFileExe, , gv0120
        If gv0120 Then
If gIsVB_Ver Then
          gv0138 = Space$(gv0120)
          For l046A = 1 To gv0120
            Get g_hFileExe, , str_len
            DbgVar str_len, "str len"
            If str_len = 0 Then
              gv0138 = Left$(gv0138, l046A - 1)
              Exit For
            Else
              Mid$(gv0138, l046A, 1) = Chr$(str_len)
            End If
          Next
Else
          gv0138 = Space$(gv0120)
          Get g_hFileExe, , gv0138
End If
        Else
          gv0138 = ""
        End If
L32E24:
        If InStr(gv0138, """") Then
          For str_len = Len(gv0138) To 1 Step -1
            If Mid$(gv0138, str_len, 1) = """" Then
              gv0138 = Left$(gv0138, str_len) & Mid$(gv0138, str_len)
            End If
          Next
        End If
        sub184C """" & gv0138 & """", mc024C
        Exit Do
      Case "n", "+", "w", "2"
        Get g_hFileExe, , l043E: p0436 = p0436 + 2
        sub184C l043E, mc016C
      Case "4"
        Get g_hFileExe, , gv078C.M3D0D: p0436 = p0436 + 4
        sub184C gv078C.M3D0D, mc016E
      Case "8"
        Get g_hFileExe, , gv07DA.M3D1B: p0436 = p0436 + 8
        sub184C gv07DA.M3D1B, mc0170
      Case "g"
      Case "a"
        Get g_hFileExe, , l043E: p0436 = p0436 + 2
        l0438 = l0438 + 1
        sub184C l043E, mc017E
      Case Else
        ShowMsgBox mc0230
      End Select
      l0438 = l0438 + 1
    Loop Until gv34F4
  End If
  If l044A Then
    p0436 = l044A
  End If
  m014E = m014A
  m014A = m014C
Exit Function

L3314E:
  Get g_hFileExe, gVAToRaw2 + p0436, signed: p0436 = p0436 + 2
  l0456 = signed
  If signed < 0 Or signed > 8 Then
    ref = fn0F99(gL3Min, l0456)
    GoTo L3332E
  Else
    ref = -(gv16BA + gLWhatEver)
  End If
  l0448 = fn1B19(ref)
Return

L331F2:
  Get g_hFileExe, p0436 + gVAToRaw2, l0456: p0436 = p0436 + 4
  ref = gv30F0(l0456)
  If ref >= 0 Then
    l0452 = gEXEdefEntries(ref).owningScope
    l0456 = l0456 - g_gbl(l0452).M49E2
  End If
  GoTo L3332E
Return

L3327E:
  Get g_hFileExe, p0436 + gVAToRaw2, l0452
  Get g_hFileExe, , signed: l0456 = ToUnSigned(signed)
  DbgVar signed >= 0, "unsigned=signed"
  p0436 = p0436 + 4
  If l0452 = gv2FDA Then
    ref = fn0F99(gLWhatEver + gv16BA, l0456)
L3332E:
    If ref < 0 Then
      l0456 = 0
      l0448 = fn104E(-ref)
    ElseIf ref > 0 Then
      l0456 = l0456 - gEXEdefEntries(ref).M689A
      l0448 = fn1B19(ref)
      Do While l0456 > 0
        sub184C l0448, mc0172
        ref = fn0F99(-gEXEdefEntries(ref).M6894.M42D1, l0456)
        If ref = 0 Then
          DbgMileStone "no base?"
          Exit Do
        ElseIf ref > gEXEdef6_Num_scopes Then
          DbgMileStone "ill. base"
          Exit Do
        End If
        l0448 = fn1831() & fn1B19(ref)
        l0456 = l0456 - gEXEdefEntries(ref).M689A
      Loop
    End If
    If l0456 Then
      l0448 = l0448 & "." & toHex16(l0456)
      DbgVar ref <= 0, "ref"
    End If
    Return
  ElseIf l0452 <= 8 And l0452 > 0 Then
    DbgMileStone "wrong para base"
  Else
    ref = fn0F99(gL3Min, CLng(l0452))
    If ref = 0 Then
      DbgMileStone "no base?"
    ElseIf ref > gEXEdef6_Num_scopes Then
      DbgMileStone "ill. base"
    Else
      l0448 = fn1B19(ref)
      ref = fn0F99(ref, l0456)
    End If
    GoTo L3332E
  End If
  l0448 = fn1B19(ref)
Return

L3364E:
  Get g_hFileExe, gVAToRaw2 + p0436, signed: p0436 = p0436 + 2
  l0456 = ToUnSigned(signed)
  l043E = m0118(m014C).M6A4D
  If l043E = mc0172 Then
    ref = m0118(m014C).TypeCode_M5F96
    If ref = 0 Then
      l0448 = fn1831() & "." & toHex16(l0456)
      Return
    ElseIf ref < 0 Then
      l0448 = fn1831()
      l0448 = fn111E(-ref, l0456)
      ref = fn0F99(-ref, l0456)
      GoTo L3332E
    Else
      l0452 = fn1175(ref)
      ErrBeep
      ref = fn0F99(ref, l0456)
      l0448 = fn1831() & fn1B19(ref)
    End If
  Else
    l0448 = fn1831()
    If l0456 Then l0448 = l0448 & "." & toHex16(l0456)
  End If
Return

L33814:
  Get g_hFileExe, , l0444: p0436 = p0436 + 2
L33838:
If gIsVB_Ver Then
  l0446 = ToUnSigned(l0444)
  l0474 = l0446 + gModuleVAStart
Else
  l0474 = l044A + ToUnSigned(l0444)
  l0446 = l0474 - gModuleVAStart
End If
  m0160 = fn1D12(l0474)
  l0448 = l0448 & fn1D20(m0160)
Return

L338D2:
  If p0436 = gVAModuleEnd Then
    l044E = gv0150
    l0442 = mc017C
    Select Case gv014E
    Case mc047C, mc047E
      l0448 = fn2023(gv0150)
    Case Else
      MsgBox "unexpected fixup"
      Return
    End Select
    p0436 = p0436 + 4
    Seek g_hFileExe, gVAToRaw2 + p0436
    sub019B
  Else
    l044E = 0
    l0442 = mc016E
    Get g_hFileExe, , gv0120: p0436 = p0436 + 4
    If gv0120 = 0 Then
      l0448 = "NULL"
    ElseIf fn087D(gv0120) = 0 Then
      l0448 = "0:" & toHex16(fn088F(gv0120))
    Else
      ErrBeep
      l0448 = "?" & Hex$(gv0120) & "?"
    End If
  End If
Return

L33A82:
  m015C = True
  If m014C Then
    DisOutput "'<unused args>": WriteOutputTxt
    While m014A > 0
      DisOutput " " & fn1831(): WriteOutputTxt
    Wend
    DisOutput "'<end unused>": WriteOutputTxt
  End If
Return
End Function

Function fn18DE(ByVal p048C As String) As Integer
Dim l048E As Integer
Dim l0490 As Integer
Dim l0492 As Integer
' Const mc0494 =-832 ' &HFCC0%
  If Len(p048C) = 0 Then Exit Function
  l048E = InStr(gc2A0A, p048C)
  If l048E = 0 Then Exit Function
  fn18DE = l048E
  Select Case m0118(m014A).M6A4D
  Case l048E
    Exit Function
  Case mc024C
    m0118(m014A).M6A4D = l048E
  Case mc0172
    l0492 = m0118(m014A).TypeCode_M5F96
    If l0492 < 0 Then
      l0492 = -l0492
    End If
    l0490 = gEXEdefEntries(l0492).M6894.M36B3_count And mc042C
    If l0490 <> l048E Then
      If l0490 = 0 Then
        gEXEdefEntries(l0492).M6894.M36B3_count = gEXEdefEntries(l0492).M6894.M36B3_count Or l048E
        Select Case gEXEdefEntries(l0492).M6894.M36B3_count And mc0494
        Case mc01B8: gIs3158 = True
        Case Else: gIs3124 = True
        End Select
      Else
      End If
    End If
  Case mc016C
    DbgVar VarType(m0118(m014A).Name) = 2, "int data"
    Select Case l048E
    Case mc0310, mc030C
    Case mc030E
    Case Else
      DbgVar False, "compatible types"
      l048E = mc016C
    End Select
    m0118(m014A).M6A4D = l048E
  Case mc016E
    DbgVar VarType(m0118(m014A).Name) = 3, "long"
    Select Case l048E
    Case mc0312
    Case mc0314
      gv078C.M3D0D = m0118(m014A).Name
      LSet gv0798 = gv078C
      m0118(m014A).Name = gv0798.M3D1B
    Case Else
      DbgVar False, "compatible types"
      l048E = mc016E
    End Select
    m0118(m014A).M6A4D = l048E
  Case mc0170
    gv07DA.M3D1B = m0118(m014A).Name
    Select Case l048E
    Case mc0320
      LSet gv07E4 = gv07DA
      m0118(m014A).Name = gv07E4.M3D62
    Case mc0322
      m0118(m014A).Name = CVDate(gv07DA.M3D1B)
    Case mc031E
    Case Else
      DbgVar False, "compatible types"
      l048E = mc0170
    End Select
    m0118(m014A).M6A4D = l048E
  End Select
End Function

Function fn18EA(p0182 As Integer, p0184 As Integer, ByVal pv0186%, p0188 As Integer, p018A As Integer, p018C As Integer) As Integer
Dim l018E As Integer
Dim l0190 As Integer
Dim l0194 As Long
Dim l019C As Integer
Dim l019E As Integer
' Const mc01AA = 4096 ' &H1000%
' Const mc01AC = 20480 ' &H5000%
' Const mc01AE = 28672 ' &H7000%
' Const mc01B8 = 12288 ' &H3000%
' Const mc01C0 = 64 ' &H40%
' Const mc01C2 = 128 ' &H80%
' Const mc01C4 = 192 ' &HC0%
' Const mc01C6 =-20480 ' &HB000%
  l0194 = ToUnSigned(pv0186)
  l019C = (g_gbl(p0188).M36B3_count And mc01AA) > 0
  Select Case p0184
  Case mc01AC, mc01AE
    l0194 = l0194 + g_gbl(p0188).M49E2
    If l0194 > gblsCountVars Then
      DbgMileStone "missing global space"
      ReDim Preserve gv30F0(l0194)
    End If
    l0190 = gv30F0(l0194)
  Case mc01B8
    l0194 = pv0186
    If l0194 > UBound(gv3126) Or l0194 < LBound(gv3126) Then
      DbgMileStone "wrong scancode"
      Exit Function
    End If
    l0190 = gv3126(l0194)
  Case mc01C0, mc01C2, mc01C4
    If pv0186 <= 8 Then
      DbgMileStone "skip me-ref"
      Exit Function
    End If
    l0190 = gv3126(l0194)
  Case mc01C6
  Case Else
    DbgMileStone "lookup scope " & CStr(p0184)
    Error gc34EA
    l0190 = 0
  End Select
  If l0190 > 0 Then
    If gEXEdefEntries(l0190).M68A4 <> l0194 Then

L2A976:
      l018E = fn1175(l0190)
      If l018E > 0 Then
        fn18EA = fn18EA(0, mc01C6, pv0186 - gEXEdefEntries(l0190).M689A, l018E, 0, g_gbl(l018E).M4950)
      End If
      Exit Function
    End If
    fn18EA = l0190
    If l0190 > 0 Then
      If p018A Then gEXEdefEntries(l0190).M6877 = p018A
    End If
    Exit Function
  ElseIf l0190 < 0 Then
    fn18EA = l0190
    Exit Function
  End If
  l019E = pv0186 Xor l019C
  If p018C Then
    If p0182 Then
      If (gEXEdefEntries(p0182).M689A Xor l019C) > l019E Then p0182 = 0
    End If
    If p0182 Then
      l018E = gEXEdefEntries(p0182).Count
    Else
      l018E = p018C
    End If
    Do While l018E
      If (gEXEdefEntries(l018E).M689A Xor l019C) >= l019E Then
        If gEXEdefEntries(l018E).M689A = pv0186 Then
          l0190 = l018E
        End If
        Exit Do
      End If
      p0182 = l018E
      l018E = gEXEdefEntries(p0182).Count
    Loop
  Else
    p0182 = 0
    l018E = 0
  End If
  If p0182 Then
    If gEXEdefEntries(p0182).M689A < pv0186 Then
      If gEXEdefEntries(p0182).M689A + ToUnSigned(gEXEdefEntries(p0182).Size) > pv0186 Then
        l0190 = p0182
        GoTo L2A976
      End If
    End If
  End If
  If l018E Then
    If gEXEdefEntries(l018E).M689A < pv0186 Then
      If gEXEdefEntries(l018E).M689A + ToUnSigned(gEXEdefEntries(l018E).Size) > pv0186 Then
        l0190 = l018E
        GoTo L2A976
      End If
    End If
  End If
  If l0190 = 0 Then
    If p0182 Then
      l0190 = fn1103(gEXEdefEntries(p0182).Count)
    Else
      l0190 = fn1103(p018C)
    End If
    gEXEdefEntries(l0190).M689A = pv0186
    gEXEdefEntries(l0190).M68A4 = l0194
    gEXEdefEntries(l0190).M6894.M36B3_count = p0184
    gEXEdefEntries(l0190).owningScope = p0188
    gEXEdefEntries(l0190).M6877 = p018A
    Select Case p0184
    Case mc01AC, mc01AE
      gv30F0(l0194) = l0190
      gIs3124 = True
    Case mc01B8, mc01C0, mc01C2, mc01C4
      If l0194 < LBound(gv3126) Then
        DbgMileStone "wrong scancode"
        Exit Function
      End If
      gv3126(l0194) = l0190
      gIs3158 = True
    Case mc01C0, mc01C2, mc01C4
      gv3126(l0194) = l0190
      gIs3124 = True
    Case mc01C6
      gIs3124 = True
    Case Else
      DbgMileStone "lookup scope " & toHex16(p0184)
      Error gc34EA
    End Select
  End If
  fn18EA = l0190
End Function

Function fn18F8(p01DC As Integer, p01DE As Long) As Integer
Dim l01E0 As Integer
Dim l01E2 As Integer
Dim l01E6 As Integer
Dim l01E8 As Long
' Const mc01EA =-8192 ' &HE000%
' Const mc01EC = 24576 ' &H6000%
' Const mc01F2 = 16384 ' &H4000%
' Const mc01F6 = 8192 ' &H2000%
Static m01F8 As Integer
' Const mc01FA =-32768 ' &H8000%
  DbgVar p01DC > 0, "scope ref"
  l01E8 = g_gbl(p01DC).M49E2 + p01DE
  l01E2 = g_gbl(p01DC).M36B3_count
  l01E0 = l01E2 And mc01EA
  Select Case l01E0
  Case mc01EC
    l01E6 = gv30F0(l01E8)
    If l01E6 = 0 Then
      l01E6 = fn18EA(0, mc01AC, p01DE, p01DC, gv16BA, g_gbl(p01DC).M4950)
    End If
    If l01E6 > 0 Then
      If gv16BA + gLWhatEver <> p01DC Then
        gEXEdefEntries(l01E6).M6877 = 0
      End If
    End If
  Case mc01F2
    If l01E6 = 0 Then
      l01E6 = fn18EA(0, mc01AE, p01DE, p01DC, gv16BC, g_gbl(p01DC).M4950)
    Else
      gEXEdefEntries(l01E6).M6877 = gv16BC
    End If
  Case mc01F6, m01F8
    l01E6 = gv3126(p01DE)
    If l01E6 = 0 Then
      l01E6 = fn18EA(0, mc01B8, p01DE, p01DC, gv16BC, g_gbl(p01DC).M4950)
    End If
  Case mc01FA
    l01E6 = fn18EA(0, mc01C6, p01DE, p01DC, 0, g_gbl(p01DC).M4950)
  End Select
  fn18F8 = l01E6
End Function

Function fn1908(p01FE As Integer) As Integer
Dim l0202 As Integer
Dim l0208 As Long
Dim l020A As Long
Dim l020C As Integer
  l0202 = gEXEdefEntries2(p01FE).M56B4
  If l0202 Then fn1908 = l0202: Exit Function
  l020A = Seek(g_hFileExe)
  l020C = gEXEdefEntries2(p01FE).subRef_M3C6A
  l0202 = fn1C9B(p01FE, True)
  l0208 = fn1483(l020C)
  Get g_hFileExe, l0208 + 8, gEXEdefEntries2(p01FE).M56BF
  g_gbl(l0202).Size = gEXEdefEntries2(p01FE).M56BF
  Seek g_hFileExe, l020A
  fn1908 = l0202
End Function

Sub sub00E9()
  m014A = 0: m0150 = 0
  ReDim m0118(100), gv356E(gv315A To gv315C)
  
  gv014C = fn1624(gVACmdEnd)
  
  sub00AC
  
  gv34F4 = 0
  gv34EC = gVACmdEnd
End Sub

Sub sub113F()
Dim l04B0%
Dim Udt As Integer
Dim l04B4 As Integer
  For l04B0 = 1 To g_userTypesCount
    Udt = gEXEdefEntries2(l04B0).subRef_M3C6A
    Assume Udt, "udt defined"
    If Udt Then
      l04B4 = fn1C9B(l04B0, False)
      sub17AA l04B4, 0, mc01C6
      sub17B8 l04B0
    End If
  Next
End Sub

Sub sub178C(p04A2 As Integer)
Dim l04A4 As Integer
Dim l04A6 As Integer
  Assume p04A2 > 0, "var ref"
  l04A4 = gEXEdefEntries(p04A2).M689A
  Assume l04A4 < 0, "lcl var"
  sub107E p04A2
  l04A6 = gv356E(l04A4)
  If l04A6 Then
    gMetaCodeLines(l04A6).M36B3_count = gMetaCodeLines(l04A6).M36B3_count Or mc0418
    gv356E(l04A4) = 0
  End If
End Sub

Sub sub1799(p04A8 As Integer, p04AA As Integer)
  Assume p04AA = mc01AE, "statics"
  If gv16BA = p04A8 Then Exit Sub
  gL1Min = gL1Max + p04A8
  sub17AA gL1Min, 0, mc01AE
  gv16BA = p04A8
End Sub

Sub sub17AA(ByVal pv0236 As Integer, p0238%, p023A%)
  If g_gbl(pv0236).M4950 = -1 Then
    g_gbl(pv0236).M4950 = 0
  End If
  m0090 = g_userTypesCount
  sub1916 g_gbl(pv0236).M49D8, p0238, pv0236, p023A, g_gbl(pv0236).M4950
  While m0090 < g_userTypesCount
    m0090 = m0090 + 1
    m008A = gEXEdefEntries2(m0090).subRef_M3C6A
    pv0236 = fn1C9B(m0090, False)
    sub1916 m008A, 0, pv0236, mc01C6, g_gbl(pv0236).M4950
    sub17B8 m0090
  Wend
End Sub

Sub sub17B8(p02C0 As Integer)
Dim l02C2 As Integer
Dim l02C4 As Integer
Dim l02C6
Dim l02C8 As Integer
Dim l02CA As Integer
Dim l02CC
Dim l02CE As Integer
  Assume p02C0, "type nr"
  If gv097C = 0 Then Exit Sub
  l02C8 = gEXEdefEntries2(p02C0).M56AD
  If l02C8 Then Exit Sub
  l02C4 = fn1C9B(p02C0, False)
  l02C2 = g_gbl(l02C4).Size
  DebugPrintVar gEXEdefEntries2(p02C0).Size = l02C2, "size mismatch"
  l02C8 = 0
  For l02C6 = 1 To gv38EC
    If gv38BA(l02C6).Size = l02C2 Then
      l02CA = gv38BA(l02C6).M4950
      l02CE = g_gbl(l02C4).M4950
      For l02CC = l02CA + 1 To l02CA + gv38BA(l02C6).M4959
        If gv38F0(l02CC).M6894.M42D9 Then
          If l02CE = 0 Then
            Select Case gv38F0(l02CC).M6894.M42D9
            Case mc0260, mc025A
            Case Else
              l02CE = -1
              Exit For
            End Select
          Else
            If gEXEdefEntries(l02CE).M689A <> gv38F0(l02CC).hFile Then
              If gEXEdefEntries(l02CE).M689A < gv38F0(l02CC).hFile Then
                Exit For
              End If
            ElseIf gEXEdefEntries(l02CE).M6894.M42D9 <> gv38F0(l02CC).M6894.M42D9 Then
              Exit For
            Else
              l02CE = gEXEdefEntries(l02CE).Count
            End If
          End If
        End If
      Next
      If l02CE = 0 Then
        If l02C8 Then
          MsgBox "ambiguous Type"
          l02C8 = -1
          Exit For
        End If
        l02C8 = l02C6
      End If
    End If
  Next
  If l02C8 = 0 Then
    DbgMileStone "no possible type found!?"
  ElseIf l02C8 > 0 Then
    gEXEdefEntries2(p02C0).M56AD = l02C8
    Assume gEXEdefEntries2(p02C0).subRef_M3C6A, "type sym"
    sub0D9B gEXEdefEntries2(p02C0).subRef_M3C6A, gv38BA(l02C8).Name
  Else
    DbgMileStone "ambiguous type"
  End If
End Sub

Sub sub17C9(p04B6 As Integer, p04B8 As Integer)
Dim l04BC As Integer
  Assume p04B8 = mc0284, "scan lcls"
  gv16BC = p04B6
  If gIsVB_Ver = 4 Then
    l04BC = &H10
  ElseIf gIsVB5 Then
    If gIsVB5 < 0 Then
      DbgMileStone "scan native sub"
      Exit Sub
    End If
    l04BC = &H18
  Else
    l04BC = &HE
  End If
  sub0EB3 p04B6
  gL3Min = gL3Max + p04B6
  gL2Min = gL2Max + p04B6
  If gL3Min > gEXEdef6_Num_scopes Then
    DebugPrintVar gL3Min <= gEXEdef6_Num_scopes, "alloc scope"
    ReDim Preserve g_gbl(gL3Min)
  End If
  gv315A = -g_gbl(gL3Min).Size
  gv315C = g_gbl(gL2Min).Size
  ReDim gv3126(gv315A To gv315C)
  sub17AA gL3Min, l04BC, mc01B8
  l04BC = gv2F50.M6832
  If l04BC Then
    sub17AA gL3Min, l04BC, mc01B8
  End If
  gIs3158 = True
End Sub

Sub sub17E7(p04CE As Integer)
Dim l04D0 As Long
Dim l04D2 As Integer
' Const mc04D4 = 39 ' &H27%
  Assume m008A = 0, "single udt"
  m008E = 0
  If p04CE And &H20 Then
    l04D0 = m0034 + Seek(g_hFileExe)
    gv014C = fn1624(l04D0)
    Assume gv014C, "patched"
    Assume gv161C(gv014C).VOffset = l04D0, "pat here"
    m008A = gv161C(gv014C).subRef_M3C6A
    While gReloc_Arr(m008A).subRef_M3C6A < 0
      m008A = -gReloc_Arr(m008A).subRef_M3C6A
    Wend
  End If
  Get g_hFileExe, , m007A
  If p04CE And &H40 Then
If gIsVB5 Then
Else
    DbgVar gIsVB_Ver, "32 bit"
    Get g_hFileExe, , gv0120
    l04D0 = m0034 + Seek(g_hFileExe) - 4
    gv014C = fn1624(l04D0)
    If gv161C(gv014C).VOffset <> l04D0 Then
      ShowMsgBox mc0230
    End If
    l04D2 = gv161C(gv014C).subRef_M3C6A
    DebugPrintVar fn202F(l04D2) = gv0120, "loc off"
    gv014C = fn1624(gv0120)
    Assume gv014C, "patches"
    m008A = gv161C(gv014C).subRef_M3C6A
End If
  ElseIf gIsVB_Ver Then
    Get g_hFileExe, , gv0120
  End If
  If m008A Then
    If gReloc_Arr(m008A).M36B3_count <> &H27 Then
      sub0C60 m008A
      If l04D2 Then
        If gReloc_Arr(l04D2).subRef_M3C6A = 0 Then gReloc_Arr(l04D2).subRef_M3C6A = -m008A
      End If
    End If
    m008C = gReloc_Arr(m008A).subRef_M3C6A
    m008E = -gEXEdef_Arr4_imports(m008C).TypeCode_M5F96
  End If
End Sub

Sub sub17F6(p028A As Integer)
Dim l028C As Long
Dim l028E As Integer
Dim l0290
Dim l0292 As Integer
Dim l0294 As Integer
Dim l0296 As Integer
' Const mc02B2 = 2 ' &H2%
' Const mc02B4 = 4 ' &H4%
' Const mc02B6 = 8 ' &H8%
' Const mc02B8 = 11 ' &HB%
' Const mc02BA = 9 ' &H9%
' Const mc02BC = 50 ' &H32%
  l0292 = g_decl_objs(p028A).M41F0
  If l0292 Then
    l028C = fn202F(l0292)
    gv356C = fn203B(l0292) \ 4
  Else
    gv356C = 0
  End If
  ReDim gv353A(gv356C)
  If gv356C = 0 Then Exit Sub
  l028E = fn1624(l028C)
  For l0290 = 0 To gv356C - 1
    l0292 = gv161C(l028E).subRef_M3C6A
    gv353A(l0290).Name = fn2023(l0292)
    l0296 = fn2047(l0292)
    l0294 = gReloc_Arr(l0292).M36B3_count
    Select Case l0294
    Case 0, 1
      l0294 = mc024C
    Case mc02B2 To mc02B4
      l0294 = mc024C
    Case mc02B6
      DbgMileStone "handle sub? ref"
    Case mc02B8
      DbgMileStone "handle sub ref"
    Case mc02BA
      DbgMileStone "handle dll ref"
    Case mc02BC To mc02BC + 16
      DbgMileStone "handle tables"
    Case Else
      l0294 = mc024C
    End Select
    gv353A(l0290).M6A4D = l0294
    gv353A(l0290).TypeCode_M5F96 = l0292
    l028E = fn164F(l028E)
  Next
End Sub

Private Sub sub1812(p023E As Integer, p0240 As Integer, p0244 As T42C7)
Static m0246 As Integer
' Const mc0248 = 11 ' &HB%
' Const mc024A = 1 ' &H1%
' Const mc024C = 28 ' &H1C%
' Const mc024E = 2 ' &H2%
' Const mc0250 = 10 ' &HA%
' Const mc0252 = 3 ' &H3%
' Const mc0254 = 29 ' &H1D%
' Const mc0256 = 4 ' &H4%
' Const mc0258 = 18 ' &H12%
' Const mc025A = 10 ' &HA%
' Const mc025C = 2 ' &H2%
' Const mc0260 = 5 ' &H5%
' Const mc0264 = 32 ' &H20%
' Const mc0266 = 16 ' &H10%
' Const mc0268 = 9 ' &H9%
' Const mc026A = 11 ' &HB%
  p0244.M42D9 = p023E And &H111F
  Select Case p023E And &H1F
  Case m0246, mc0248
    p0244.M42D9 = 0
    p0244.M42D1 = 0: p0244.M36B3_count = 0: p0240 = 0
  Case mc024A: p0244.M42D1 = 0: p0244.M36B3_count = mc024C: p0240 = 4
  Case mc024E: p0244.M42D1 = 0: p0244.M36B3_count = mc0250: p0240 = &H10
  Case mc0252: p0244.M42D1 = 0: p0244.M36B3_count = mc0254: p0240 = 4
  Case mc0256
    p0244.M36B3_count = mc0258
    p0240 = 4
  Case mc025A
    p0244.M36B3_count = mc025C
    p0240 = p0244.M42D1
If gIsVB_Ver Then
    p0240 = p0240 * 2
End If
  Case mc0260
    p0244.M42D1 = gEXEdef_Arr1Count
    If p023E And &H100 Then
      p0240 = 4
      p0244.M36B3_count = mc0264 Or mc0266
    Else
      p0244.M36B3_count = mc0264
    End If
  Case 6: p0244.M42D1 = 0: p0244.M36B3_count = 0
  Case mc0268
    p0244.M36B3_count = mc026A
    If p023E And &H1000 Then
      DebugPrintVar m008E, "unaligned udt"
      p0240 = gEXEdefEntries2(m008E).M56B4
      If p0240 = 0 Then
        p0240 = fn1908(m008E)
      End If
      p0244.M42D1 = -p0240
      p0240 = gEXEdefEntries2(m008E).M56BF
    Else
      If m008E Then
        p0244.M42D1 = -fn1C9B(m008E, False)
        p0240 = gEXEdefEntries2(m008E).Size
      End If
    End If
  Case Else
    ShowMsgBox mc0230
    p0244.M42D1 = 0: p0244.M36B3_count = 0
  End Select
  If p023E And &H800 Then
    DbgVar m0054 > 0, "para off"
    DbgVar m0064 = mc01B8, "lcl scope"
    p0244.M36B3_count = p0244.M36B3_count Or mc01C2
  Else
    p0244.M36B3_count = p0244.M36B3_count Or m0064
  End If
End Sub

Sub sub1838()
Dim l0430
Dim l0432 As Integer
  l0432 = m0118(m014A).TypeCode_M5F96
  If l0432 = m0162 Then m0162 = 0
  For l0430 = m014A - 1 To 1 Step -1
    If m0118(l0430).TypeCode_M5F96 = l0432 Then
      If m0118(l0430).M6A4D <> mc0172 Then
        m0118(m014A).Name = m0118(l0430).Name
        m0118(l0430).M6A4D = mc016A
        m0118(l0430).TypeCode_M5F96 = 0
        m0150 = m0150 + 1
      End If
      Exit For
    End If
  Next
End Sub

Sub sub184C(p0332 As Variant, p0336 As Integer)
Dim l0338 As Integer
  m014A = m014A + 1
  If m014A > UBound(m0118) Then ReDim Preserve m0118(m014A + 20)
  Select Case p0336
  Case mc0172
    If Left$(p0332, 1) = "^" Then p0332 = Mid$(p0332, 2)
    l0338 = InStr(p0332, "(")
    If l0338 = 0 Then l0338 = InStr(p0332, " ")
    If l0338 Then p0332 = Left$(p0332, l0338 - 1)
  Case mc017E
    m0118(m014A) = gv353A(p0332)
    m0118(m014A).VOffset = gVACmdStart
    Exit Sub
  End Select
  m0118(m014A).VOffset = gVACmdStart
  m0118(m014A).Name = p0332
  m0118(m014A).M6A4D = p0336
  m0118(m014A).TypeCode_M5F96 = 0
End Sub

Sub sub1854(p033A As Integer, p033C As Integer)
Dim l0340 As T6A40
  l0340 = m0118(p033A)
  If p033C Then
    If l0340.TypeCode_M5F96 Then
      Select Case l0340.M6A4D
      Case mc017C
        l0340.TypeCode_M5F96 = 0
      End Select
    End If
    l0340.M6A4D = p033C
  End If
  m014A = m014A + 1
  m0118(m014A) = l0340
End Sub

Sub sub1860(p0342 As Integer, p2%)
' called (p1 As type, p2%)
Dim l0346 As Integer
Dim l0348 As Integer
Dim l034A
Dim l034C As Integer
Dim l034E As Integer
Dim l0350 As Long
Dim l0352 As Integer
Dim l0354 As Integer
Dim l0356 As Integer
Dim l035A As Integer
Dim l0360 As Integer
Dim l0364%
Dim l0366 As Long
' Const mc0372 = 32 ' &H20%
' Const mc0374 = 33 ' &H21%
' Const mc037E = 35 ' &H23%
' Const mc0380 = 38 ' &H26%
' Const mc0382 = 34 ' &H22%
' Const mc03A0 = 2 ' &H2%
' Const mc03A2 = 3 ' &H3%
' Const mc03A4 = 16 ' &H10%
' Const mc03A6 = 1 ' &H1%
' Const mc03A8 = 8 ' &H8%
' Const mc03AC = 41 ' &H29%
' Const mc03B4 =-8192 ' &HE000%
' Const mc03B6 =-32768 ' &H8000%
' Const mc03BE =-16384 ' &HC000%
Dim l03C2 As Long
Dim l03C4%
On Error GoTo L2E324
  gv34F4 = gv34FA
If gIsVB5 Then
  DbgMileStone "scan_sub VB5 native"
  Exit Sub
End If
  gv0468 = p0342
  If p0342 > gEXEdef_Arr5Count Then
    gv34F4 = gc34FE
    Exit Sub
  ElseIf p0342 < 0 Then
    l0346 = -p0342
  ElseIf gEXEdef_Subs(p0342).ModuleIndex Then
    l0346 = gEXEdef_Subs(p0342).lOffsetSym_M3C6A
    Assume gReloc_Arr(l0346).M36B3_count = mc0372, "hdr"
    l0346 = gReloc_Arr(l0346).M54BB
    Assume gReloc_Arr(l0346).M36B3_count = mc0374, "tokens"
    gv0466 = gEXEdef_Subs(p0342).ModuleIndex
    If gv0466 = 0 Then
      DbgMileStone "Scan Sub ohne Module?"
      Exit Sub
    End If
  Else
    Exit Sub
  End If
  gv0300 = fn202F(l0346)
  gModuleVACur = fn202F(gReloc_Arr(l0346).Count)
  gv0462 = gv0300
  Select Case gReloc_Arr(l0346).M36B3_count
  Case mc037E, mc0380
    DbgMileStone "scan form"
  Case mc0382
    DbgMileStone "scan decls"
  Case mc0374
    sub0EB3 p0342
    sub0487 p0342, True
    If g_gbl(gL3Min).M4950 < 0 Then
      DbgMileStone "scopes VB5"
      Exit Sub
    End If
    gv0144 = gModuleVACur
    gv014C = fn1624(gv0300)
    sub00AC
    gVACmdEnd = gv0300
    Do While gVACmdEnd < gModuleVACur
      gVACmdStart = gVACmdEnd
If gIsVB5 Then
      l0352 = fn1EA9(gVAToRaw2 + gVACmdEnd)
      gVACmdEnd = gVACmdEnd + l0352
Else
      Get g_hFileExe, gVAToRaw2 + gVACmdEnd, gv34F6: gVACmdEnd = gVACmdEnd + 2
End If
      l034C = gDatArrLinks(gv34F6)
      Assume l034C, "arg scanned"
      l0350 = 0
      l034E = gDatArr1(l034C).M3818
      Do
        Select Case Mid$(gDatStrings1, l034E, 1)
        Case "§"
          Exit Do
        Case "}"
          Exit Sub
        Case ":"
          sub186C gVACmdStart, 0
          Exit Do
        Case "_", "="
          sub186C gVACmdEnd, 0
          Exit Do
        Case "1"
          sub186C gVACmdStart, 0
          sub186C gVACmdEnd, 0
          Exit Do
        Case "-"
          GoTo L2E236
        Case "ß"
        Case "~"
          Get g_hFileExe, , l0352: gVACmdEnd = gVACmdEnd + 2
          l0350 = gVACmdEnd + l0352
        Case "i", "l"
          GoSub L2E426
        Case "L"
          Assume l0350 > gVACmdEnd, "next token pos"
          While gVACmdEnd < l0350
            GoSub L2E426
            l035A = 0
          Wend
          Exit Do
        Case "M"
          GoSub L2E55C
        Case "G"
          GoSub L2E516
        Case "s"
          GoSub L2E346
        Case "c"
          gVACmdEnd = gVACmdEnd + 2
        Case "m"
          gVACmdEnd = gVACmdEnd + 2
        Case "t"
          l0360 = mc03A0
          GoSub L2E7C8
        Case "f"
          l0360 = mc03A2
          GoSub L2E7C8
        Case "N"
          l0360 = mc03A4
          GoSub L2E7C8
        Case "j", "d"
          l0360 = mc03A6
          GoSub L2E7C8
        Case "e"
          l0360 = mc03A8
          GoSub L2E7C8
        Case "D"
          l0360 = mc03A8
          Assume l0350 > gVACmdEnd, "next token pos"
          While gVACmdEnd < l0350
            GoSub L2E87E
          Wend
        Case "P"
          GoSub L2E6B0
        Case "a"
          Get g_hFileExe, , l034A: gVACmdEnd = gVACmdEnd + 2
          If l034A >= 0 And l034A < gv356C Then
            l0348 = gv353A(l034A).TypeCode_M5F96
            GoTo L2DDD6
          Else
            MsgBox "token " & Hex$(gv34F6), , "NOT arg"
          End If
        Case "p"
          GoSub L2E6B0
L2DDD6:
          l034E = l034E + 1
          If l0348 Then
            Select Case Mid$(gDatStrings1, l034E, 1)
            Case "s"
            Case "L"
              l0346 = fn206B(l0348, mc0382)
            Case "z", "x", "p"
              l0346 = fn206B(l0348, mc02B2)
            Case "d"
              l0346 = fn206B(l0348, mc03AC)
            Case "c"
              If gReloc_Arr(l0348).M36B3_count <> mc0380 Then
                l0346 = fn206B(l0348, mc0380)
              End If
            Case "$"
              l0346 = fn206B(l0348, mc02B2)
            Case "?"
            Case Else
              If InStr("%&!#@$?", Mid$(gDatStrings1, l034E, 1)) Then
              Else
              gv34F4 = gc3500
              End If
            End Select
          End If
        Case "."
          gVACmdEnd = gVACmdEnd + 2
        Case "°"
          gVACmdEnd = gVACmdEnd + 1
        Case "%"
          gVACmdEnd = gVACmdEnd + 2
        Case "&", "!", "4"
          gVACmdEnd = gVACmdEnd + 4
        Case "#", "@", "8"
          gVACmdEnd = gVACmdEnd + 8
        Case "u"
          If l0350 Then Exit Do
          If g_IsNE Then ShowMsgBox mc0230
        Case "$"
          If l0350 Then Exit Do
        Case "n": gVACmdEnd = gVACmdEnd + 2
        Case "+": gVACmdEnd = gVACmdEnd + 2
        Case "w": gVACmdEnd = gVACmdEnd + 2
        Case "g"
          gVACmdEnd = l0350
        Case Else
          ShowMsgBox mc0230
        End Select
        l034E = l034E + 1
      Loop Until gv34F4
      gv2FDE = l035A
      l035A = 0
L2E236:
      If gv34F4 Then Exit Do
      If l0350 Then
        gVACmdEnd = l0350
      End If
      If gVACmdEnd > gVAModuleEnd Then
        ShowMsgBox mc0230
      End If
      l0354 = l034C
    Loop
    gv0462 = gVACmdStart
    If gv34F4 = gv34FA Then
      If gVACmdEnd = gModuleVACur Then
        Exit Sub
      End If
      gv34F4 = gc3500
    End If
    Exit Sub
  Case Else
    ShowMsgBox mc0230
  End Select
  gv34F4 = gc34FC
Exit Sub

L2E324:
  gv34F4 = gc3500
  gv0462 = gVACmdStart
Exit Sub

L2E346:
  Get g_hFileExe, gVAToRaw2 + gVACmdEnd, l0364: gVACmdEnd = gVACmdEnd + 2
  l0366 = ToUnSigned(l0364)
  If gv2FDE = 0 Then Return
  If gv2FDE < 0 Then
    l0356 = -gv2FDE
    l035A = fn18F8(l0356, l0366)
  Else
    l0356 = gEXEdefEntries(gv2FDE).M6894.M36B3_count And mc03B4
    l035A = fn1BD4(gv2FDE, gEXEdefEntries(gv2FDE), l0356 Or mc03B6, gv16BC, l0366)
  End If
Return

L2E426:
  Get g_hFileExe, gVAToRaw2 + gVACmdEnd, l0364: gVACmdEnd = gVACmdEnd + 2
  l0366 = l0364
    l0356 = mc01B8
    If l0364 < 0 Then
      l035A = fn18EA(0, l0356, l0366, gL3Min, gv16BC, g_gbl(gL3Min).M4950)
    ElseIf l0364 > 8 Then
      l035A = fn18EA(0, l0356, l0366, gL2Min, gv16BC, g_gbl(gL2Min).M4950)
    Else
      l035A = -(gv16BA + gLWhatEver)
    End If
Return

L2E516:
  Get g_hFileExe, gVACmdEnd + gVAToRaw2, l0366: gVACmdEnd = gVACmdEnd + 4
  l035A = fn177E(l0366)
Return

L2E55C:
  Get g_hFileExe, gVACmdEnd + gVAToRaw2, l0356
  Get g_hFileExe, , l0364: l0366 = ToUnSigned(l0364)
  gVACmdEnd = gVACmdEnd + 4
  If l0356 = gv2FDA Then
      l035A = fn18EA(0, mc01AC, l0366, gLWhatEver + gv16BA, gv16BA, g_gbl(gLWhatEver + gv16BA).M4950)
  ElseIf l0356 <= 8 Then
    DbgMileStone "wrong para base"
  Else
    l035A = fn18EA(0, mc01B8, l0356, gL2Min, gv16BC, g_gbl(gL2Min).M4950)
    l035A = fn1BD4(l035A, gEXEdefEntries(l035A), mc03BE, gv16BC, l0366)
  End If
Return

L2E6B0:
  If gVAModuleEnd = gVACmdEnd Then
    l0348 = gv0150
    gVACmdEnd = gVACmdEnd + 4
    sub019B
  Else
    l0348 = 0
    Get g_hFileExe, gVAToRaw2 + gVACmdEnd, gv0120: gVACmdEnd = gVACmdEnd + 4
    If gv0120 Then
If gIsVB_Ver = 4 Then
      Select Case gv34F6
      Case &H5CC, &H5CE
        Return
      End Select
ElseIf g_IsNE Then
      Select Case gv34F6
      Case &H3AE5, &H3AF5
        Return
      End Select
End If
      gv34F4 = gc3500
    End If
  End If
Return

L2E7C8:
  Get g_hFileExe, , l03C4: gVACmdEnd = gVACmdEnd + 2
  Select Case l03C4
  Case -2, -1
  Case Else
If gIsVB_Ver Then
    l03C2 = ToUnSigned(l03C4) + gModuleVAStart
Else
    l03C2 = gVACmdEnd + ToUnSigned(l03C4)
End If
    sub186C l03C2, l0360
  End Select
Return

L2E87E:
  Get g_hFileExe, , l03C4: gVACmdEnd = gVACmdEnd + 2
  Select Case l03C4
  Case -2, -1
  Case Else
If gIsVB_Ver Then
    l03C2 = ToUnSigned(l03C4) + gModuleVAStart
Else
    l03C2 = l0350 + ToUnSigned(l03C4)
End If
    sub186C l03C2, l0360
  End Select
Return
End Sub

Sub sub1886(p03D2 As Integer)
  If p03D2 > m014A - m0150 Then
    DbgMileStone "drop>stack"
    gv34F4 = gc3500
    Exit Sub
  End If
  While p03D2 > 0
    While m0118(m014A).M6A4D = mc016A
      m014A = m014A - 1
      m0150 = m0150 - 1
    Wend
    m014A = m014A - 1
    p03D2 = p03D2 - 1
  Wend
End Sub

Sub sub1892(p03D6 As Control)
Dim l03D8
Dim l03DA As Integer
  p03D6.AddItem "Stack:"
  For l03D8 = 1 To m014A
    If m0118(l03D8).M6A4D <> mc016A Then
      p03D6.AddItem fn086D(m0118(l03D8).VOffset) & " " & fn1820(m014A - l03D8)
    End If
  Next
  p03D6.AddItem "Parameter:"
  For l03D8 = l03D8 To m014E
    p03D6.AddItem m0118(l03D8).Name
  Next
  p03D6.AddItem "Temps:"
  For l03D8 = gv315A To 0
    If gv356E(l03D8) Then
      p03D6.AddItem "l" & toHex16(l03D8) & " = " & gMetaCodeLines(l03DA).M574B
    End If
  Next
End Sub

Sub sub189E()
  While m0118(m014A).M6A4D = mc016A
    m014A = m014A - 1
    m0150 = m0150 - 1
  Wend
  Assume m0118(m014A).M6A4D = mc0172, "ref tmpvar"
  sub178C m0118(m014A).TypeCode_M5F96
  m014A = m014A - 1
End Sub

Sub sub18C3(p0484 As String, p0486 As Integer)
Dim l0488 As String
  If Left$(gDisOutputTxt, 1) = "_" Then
    l0488 = Mid$(gDisOutputTxt, 5)
  Else
    l0488 = Mid$(gDisOutputTxt, 4)
  End If
  Select Case UCase$(l0488)
  Case "ANSIVALUEBSTR"
    p0484 = "F%$": gDisOutputTxt = "Asc"
  Case "APPACTIVATE"
    p0484 = "S~~": gDisOutputTxt = l0488
  Case "ATN", "COS", "EXP", "LOG", "SIN", "SQR", "TAN"
    p0484 = "F#~": gDisOutputTxt = l0488
  Case "ErrBeep2"
    p0484 = "S": gDisOutputTxt = l0488
  Case "BSTRFROMANSI"
    p0484 = "F$&": gDisOutputTxt = "Chr$"
  Case "BSTRFROMBYTE"
    p0484 = "F$&": gDisOutputTxt = "ChrB$"
  Case "BSTRFROMERROR"
    p0484 = "F$~": gDisOutputTxt = "Error$"
  Case "BSTRFROMFORMATVAR"
    p0484 = "F$~~~~": gDisOutputTxt = "Format$"
  Case "BYTEVALUEBSTR"
    p0484 = "F°$": gDisOutputTxt = "AscB"
  Case "CREATEOBJECT"
    p0484 = "F.$": gDisOutputTxt = l0488
  Case "CHANGEDIR", "CHANGEDRIVE"
    p0484 = "S$": gDisOutputTxt = "Ch" & Mid$(l0488, 7)
  Case "CHOOSE"
    p0484 = "F=~~": gDisOutputTxt = l0488
  Case "COMMANDBSTR"
    p0484 = "F$": gDisOutputTxt = "Command$"
  Case "COMMANDVAR"
    p0484 = "F=0": gDisOutputTxt = "Command"
  Case "CURRENTDIR"
    p0484 = "F=~": gDisOutputTxt = "CurDir"
  Case "CURRENTDIRBSTR"
    p0484 = "F$~": gDisOutputTxt = "CurDir$"
  Case "DATEADD"
    p0484 = "F=$~~": gDisOutputTxt = l0488
  Case "DATEDIFF"
    p0484 = "F=$~~~~": gDisOutputTxt = l0488
  Case "DATEPART"
    p0484 = "F=$~~~": gDisOutputTxt = l0488
  Case "DDB"
    p0484 = "F#~~~~~": gDisOutputTxt = l0488
  Case "DELETESTRING"
    p0484 = "S": gDisOutputTxt = l0488
  Case "DIR"
    p0484 = "F$~~": gDisOutputTxt = l0488
  Case "DOEVENTS"
    If Mid$(p0484, 2) = "%" Then
      p0484 = "F%": gDisOutputTxt = l0488
    Else
      p0484 = "S": gDisOutputTxt = l0488
    End If
  Case "ENDOFFILE"
    p0484 = "F%%": gDisOutputTxt = "EOF"
  Case "ENVIRONBSTR"
    p0484 = "F$~": gDisOutputTxt = "Environ$"
  Case "ENVIRONVAR"
    p0484 = "F=~": gDisOutputTxt = "Environ"
  Case "ERROBJ"
    p0484 = "'Err": gDisOutputTxt = ""
  Case "FILEATTRIBUTES"
    p0484 = "F&%%": gDisOutputTxt = Left$(l0488, 8)
  Case "FILECOPY"
    p0484 = "S": gDisOutputTxt = l0488
  Case "FILEDATETIME"
    p0484 = "F=$": gDisOutputTxt = l0488
  Case "FILELEN"
    p0484 = "F&$": gDisOutputTxt = l0488
  Case "FILELENGTH"
    p0484 = "F&%": gDisOutputTxt = "LOF"
  Case "FILELOCATION"
    p0484 = "F%%": gDisOutputTxt = "Loc"
  Case "FILERESET"
    p0484 = "S": gDisOutputTxt = "Reset"
  Case "FILESEEK"
    p0484 = "F&%": gDisOutputTxt = "Seek"
  Case "FILEWIDTH"
    p0484 = "S1": gDisOutputTxt = "Width"
  Case "FREEFILE"
    p0484 = "F%~": gDisOutputTxt = l0488
  Case "FV"
    p0484 = "F#~~~~~": gDisOutputTxt = l0488
  Case "GETALLSETTINGS"
    p0484 = "F=$$": gDisOutputTxt = l0488
  Case "GETDATEBSTR", "GETTIMEBSTR"
    p0484 = "F$0": gDisOutputTxt = Mid$(l0488, 4, 4)
  Case "GETDATEVAR", "GETTIMEVAR"
    If gIsVB_Ver Then
      p0484 = "F~0"
    Else
      p0484 = "F="
    End If
    gDisOutputTxt = Mid$(l0488, 4, 4)
  Case "GETDATEVALUE", "GETTIMEVALUE"
    p0484 = "F=$": gDisOutputTxt = Mid$(l0488, 4)
  Case "GETDAYOFWEEK"
    p0484 = "F=~~": gDisOutputTxt = "WeekDay"
  Case "GETERL"
    p0484 = "'Erl": gDisOutputTxt = ""
  Case "GETFILEATTR"
    p0484 = "F%$": gDisOutputTxt = "GetAttr"
  Case "GETDAYOFMONTH", "GETHOUROFDAY", "GETMINUTEOFHOUR", "GETMONTHOFYEAR", "GETSECONDOFMINUTE"
    p0486 = InStr(UCase(l0488), "OF")
    p0484 = "F=~": gDisOutputTxt = Mid$(l0488, 4, p0486 - 4)
  Case "GETOBJECT"
    p0484 = "F.~~": gDisOutputTxt = l0488
  Case "GETPRESENTDATE"
    p0484 = "F=0": gDisOutputTxt = "Now"
  Case "GETSETTING"
    p0484 = "F$$$$~": gDisOutputTxt = l0488
  Case "GETTIMER"
    p0484 = "F#0"
    gDisOutputTxt = "Timer"
  Case "GETYEAR"
    p0484 = "F=~": gDisOutputTxt = "Year"
  Case "HEXBSTRFROMVAR", "OCTBSTRFROMVAR"
    p0484 = "F$~": gDisOutputTxt = Left$(l0488, 3) & "$"
  Case "HEXVARFROMVAR", "OCTVARFROMVAR"
    p0484 = "F=~": gDisOutputTxt = Left$(l0488, 3)
  Case "IMESTATUS"
    p0484 = "'" & l0488: gDisOutputTxt = ""
  Case "IMMEDIATEIF"
    p0484 = "F=~~~": gDisOutputTxt = "IIf"
  Case "INPUTBOX"
    p0484 = "F$~~~~~~~": gDisOutputTxt = l0488
  Case "INPUTCHARCOUNT"
    p0484 = "F$&%": gDisOutputTxt = "Input$"
  Case "INPUTCHARCOUNTVAR"
    p0484 = "F=&%": gDisOutputTxt = "Input"
  Case "INPUTCOUNT"
    p0484 = "F$&%": gDisOutputTxt = "InputB$"
  Case "INPUTCOUNTVAR"
    p0484 = "F=$&": gDisOutputTxt = "InputB"
  Case "IPMT"
    p0484 = "F#~~~~~~": gDisOutputTxt = "IPmt"
  Case "ISARRAY", "ISDATE", "ISEMPTY", "ISERROR", "ISMISSING", "ISNULL", "ISNUMERIC", "ISOBJECT"
    p0484 = "F²~": gDisOutputTxt = l0488
  Case "KILLFILES"
    p0484 = "S1": gDisOutputTxt = "Kill"
  Case "LEFTCHARBSTR"
    p0484 = "F$$&": gDisOutputTxt = "Left$"
  Case "LEFTCHARVAR"
    p0484 = "F=$&": gDisOutputTxt = "Left"
  Case "LEFTBSTR"
    p0484 = "F$$&": gDisOutputTxt = "LeftB$"
  Case "LEFTVAR"
    p0484 = "F=$&": gDisOutputTxt = "LeftB"
  Case "LOWERCASEBSTR", "UPPERCASEBSTR"
    p0484 = "F$$": gDisOutputTxt = Left$(l0488, 1) & "Case$"
  Case "LOWERCASEVAR", "UPPERCASEVAR"
    p0484 = "F=~": gDisOutputTxt = Left$(l0488, 1) & "Case"
  Case "MAKEDIR"
    p0484 = "S": gDisOutputTxt = "MkDir"
  Case "MIDCHARBSTR"
    p0484 = "F$$&~": gDisOutputTxt = "Mid$"
  Case "MIDCHARVAR"
    p0484 = "F=~&~": gDisOutputTxt = "Mid"
  Case "MIDVAR"
    p0484 = "F=~&~": gDisOutputTxt = "MidB"
  Case "MSGBOX"
    If Mid$(p0484, 2, 1) = "%" Then
      p0484 = "F%~~~~~"
    Else
      p0484 = "S"
    End If
    gDisOutputTxt = l0488
  Case "NPER"
    p0484 = "F#~~~~~": gDisOutputTxt = l0488
  Case "PACKDATE", "PACKTIME"
    p0484 = "F=%%%": gDisOutputTxt = Right$(l0488, 4) & "Serial"
  Case "PARTITION"
    p0484 = "F=~~~~": gDisOutputTxt = l0488
  Case "PMT"
    p0484 = "F#~~~~~": gDisOutputTxt = "Pmt"
  Case "PPMT"
    p0484 = "F#~~~~~~": gDisOutputTxt = "PPmt"
  Case "PV"
    p0484 = "F#~~~~~": gDisOutputTxt = l0488
  Case "QBCOLOR"
    p0484 = "F&%": gDisOutputTxt = l0488
  Case "R8VALFROMBSTR"
    If gIsVB_Ver Then
      p0484 = "F#$": gDisOutputTxt = "Val"
    Else
      p0484 = "F=$": gDisOutputTxt = "Val"
    End If
  Case "RANDOMIZE"
    p0484 = "S": gDisOutputTxt = l0488
  Case "RANDOMNEXT"
    p0484 = "F#~": gDisOutputTxt = "Rnd"
  Case "RATE"
    p0484 = "F#~~~~~~": gDisOutputTxt = l0488
  Case "REMOVEDIR"
    p0484 = "S": gDisOutputTxt = "RmDir"
  Case "RGB"
    p0484 = "F&%%%": gDisOutputTxt = l0488
  Case "RIGHTBSTR"
    p0484 = "F$$&": gDisOutputTxt = "RightB$"
  Case "RIGHTCHARBSTR"
    p0484 = "F$$&": gDisOutputTxt = "Right$"
  Case "RIGHTCHARVAR"
    p0484 = "F=$&": gDisOutputTxt = "Right"
  Case "RIGHTVAR"
    p0484 = "F=~&": gDisOutputTxt = "RightB"
  Case "SENDKEYS"
    p0484 = "S": gDisOutputTxt = l0488
  Case "SETDATEBSTR", "SETTIMEBSTR"
    p0484 = "S1": gDisOutputTxt = Mid$(l0488, 4, 4) & "$ ="
  Case "SETDATEVAR", "SETTIMEVAR"
    p0484 = "S1": gDisOutputTxt = Mid$(l0488, 4, 4) & " ="
  Case "SETFILEATTR"
    p0484 = "S1": gDisOutputTxt = "SetAttr"
  Case "SHELL"
    p0484 = "F#~~": gDisOutputTxt = l0488
  Case "SLN"
    p0484 = "F#~~~": gDisOutputTxt = l0488
  Case "SPACEBSTR"
    p0484 = "F$&": gDisOutputTxt = "Space$"
  Case "SPACEVAR"
    p0484 = "F=&": gDisOutputTxt = "Space"
  Case "STRCONVVAR"
    p0484 = "F=~%": gDisOutputTxt = "StrConv"
  Case "STRFROMVAR"
    p0484 = "F$~": gDisOutputTxt = "Str$"
  Case "STRINGBSTR"
    p0484 = "F$&$": gDisOutputTxt = "String$"
  Case "STRINGVAR"
    p0484 = "F=&~": gDisOutputTxt = "String"
  Case "SWITCH"
    p0484 = "F=~": gDisOutputTxt = l0488
  Case "SYD"
    p0484 = "F#~~~~": gDisOutputTxt = l0488
  Case "TRIMBSTR"
    p0484 = "F$$": gDisOutputTxt = "Trim$"
  Case "TRIMVAR"
    p0484 = "F=~": gDisOutputTxt = "Trim"
  Case "LEFTTRIMBSTR", "RIGHTTRIMBSTR"
    p0484 = "F$$": gDisOutputTxt = Left$(l0488, 1) & "Trim$"
  Case "LEFTTRIMVAR", "RIGHTTRIMVAR"
    p0484 = "F=~": gDisOutputTxt = Left$(l0488, 1) & "Trim"
  Case "TYPENAME"
    p0484 = "F$~": gDisOutputTxt = l0488
  Case "VARBSTRFROMANSI"
    p0484 = "F=~": gDisOutputTxt = "Chr"
  Case "VARBSTRFROMBYTE"
    p0484 = "F=~": gDisOutputTxt = "ChrB"
  Case "VARDATEFROMVAR"
    p0484 = "F=~": gDisOutputTxt = "CVDate"
  Case "VARFROMERROR"
    p0484 = "F=~": gDisOutputTxt = "Error"
  Case "VARFROMFORMATVAR"
    p0484 = "F=~~~~": gDisOutputTxt = "Format"
  Case "VARSTRFROMVAR"
    p0484 = "F=~": gDisOutputTxt = "Str"
  Case "VARTYPE"
    p0484 = "F%~": gDisOutputTxt = l0488
  Case Else
    gv34F4 = gc3500
    Exit Sub
  End Select
  p0486 = 1
End Sub

Private Sub sub1916(p021E As Integer, p0220%, p0222 As Integer, p0224 As Integer, p0226%)
Dim l0228 As Integer
Dim l022A%
Dim l022C As Integer
Dim l022E As Integer
' Const mc0230 = 32 ' &H20%
: l022E = -1
  If p021E = 0 Then Exit Sub
  m0064 = p0224
  Select Case p0224
  Case mc01AC:  l022C = p0222
  Case mc01AE: l022C = 0
  Case mc01B8:  l022C = gv16BC
  Case mc01C6
  Case Else
    ShowMsgBox mc0230
    Exit Sub
  End Select
  l0228 = fn17D8(p021E, p0220)
  If l0228 Then
    Do
      sub1812 l0228, m005A, m005E
      Select Case l0228
      Case 0 To 6, 9, 10, 11, &H105
      Case &H803
      Case &H1009
      Case Else
        MsgBox "VBTyp " & Hex$(l0228)
        m005E.M36B3_count = 0
      End Select
      If l0228 And &H800 Then
        DebugPrintVar m0056 = &H803, "obj byval"
        Assume l022C, "owner"
        DbgVar p0224 = mc01B8, "lcl scope"
        l022A = fn18EA(l022A, mc01C2, m0054, gL2Max + l022C, l022C, g_gbl(gL2Max + l022C).M4950)
        If l022A Then gEXEdefEntries(l022A).owningScope = gL2Max + l022C
      Else
        l022A = fn18EA(l022A, p0224, m0054, p0222, l022C, p0226)
        If l022A Then gEXEdefEntries(l022A).owningScope = p0222
      End If
      If l022A Then
        gEXEdefEntries(l022A).M6894 = m005E
        gEXEdefEntries(l022A).Size = m005A
        gEXEdefEntries(l022A).M6877 = l022C
      Else
        DbgMileStone "ref not created"
        Exit Sub
      End If
      l0228 = fn1804()
    Loop While l0228
  Else
  End If
If gIsVB5 Then
  If m0048.M5EC8 Then
    g_gbl(p0222).Size = m0048.M5EC8
  End If
Else
  If m003A.M5EC8 Then
    g_gbl(p0222).Size = m003A.M5EC8
  End If
End If
If gIsVB5 Then
  DbgMileStone "handle offset tables"
Else
  Select Case p0224
  Case mc01AC
    sub0FFC p0222, gv30F0()
  Case mc01AE
    sub0FFC p0222, gv30F0()
  Case mc01B8
    sub0FFC p0222, gv3126()
  End Select
End If
End Sub

Sub sub1925()
Dim i_decl As Integer
Dim l026E As Integer
Dim l0270 As Integer
Dim l0272 As Integer
' Const mc0284 = 4 ' &H4%
Dim l0286 As Integer
  
  frmScan.control5 = "Scanning"
  frmScan.control3 = "Module"
  frmScan.Show
  DoEvents
  ScanModule
  
  frmScan.control4 = "declarations"
  For i_decl = 1 To g_decl_num
    frmScan.control2 = g_decl_objs(i_decl).Name_M411A
    DoEvents
    Assume g_decl_objs(i_decl).Size = g_gbl(i_decl).Size, "modvar size"
    sub17AA gLWhatEver + i_decl, 0, mc01AC
  Next
  
On Error GoTo Err_Subroutine
  frmScan.control5 = "Subroutine"
  For i_decl = 1 To g_decl_num
    frmScan.control2 = g_decl_objs(i_decl).Name_M411A
    frmScan.control4 = "statics"
    DoEvents
    sub1799 i_decl, mc01AE
If gIsVB5 >= 0 Then
    l026E = g_decl_objs(i_decl).M3F1B
    For l0270 = l026E + 1 To l026E + g_decl_objs(i_decl).M3F13
      frmScan.control4 = fn0C98(l0270)
      DoEvents
      sub17C9 l0270, mc0284
      SaveSub True
    Next
Else
    DbgMileStone "scan native locals"
End If
    sub1943 False
  Next
  frmScan.control4 = "User Types"
  For l0272 = 1 To g_userTypesCount
    l026E = fn1C9B(l0272, False)
    sub17AA l026E, 0, mc01C6
    l026E = gEXEdefEntries2(l0272).M56B4
    If l026E Then
      sub17AA l026E, 0, mc01C6
    End If
  Next
  If Not l0286 Then
    frmScan.control2 = "all"
    For l0270 = 1 To gEXEdef_Arr5Count
      frmScan.control4 = fn0C98(l0270)
      DoEvents
      sub1860 l0270, False
    Next
  End If
  gIs3124 = True
  SaveAll False

L2BF92:
  Unload frmScan
Exit Sub

Err_Subroutine:
  ShowMsgBox mc0230
  Resume L2BF92
End Sub

