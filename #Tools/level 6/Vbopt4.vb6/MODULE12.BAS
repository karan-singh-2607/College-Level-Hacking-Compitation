Attribute VB_Name = "MODULE12"
' Module12
Option Explicit
Dim m0018 As String
Dim m001C As String
Dim m0020 As String
Dim m0024 As String
Dim m002A() As T40FB
Dim m005C As Integer
Dim m005E As Integer
Dim m0060 As Integer
Dim m0062 As Integer
Dim m0064 As Integer
Const mc0066 = 1 ' &H1%
Const mc0068 = 2 ' &H2%
Const mc006A = 3 ' &H3%
Const mc006C = 4 ' &H4%
Dim m0070() As T40FB
Dim m00A4() As T40FB
Const mc00DC = 2 ' &H2%
Const mc00E8 = 16 ' &H10%
Const mc00EC = 3 ' &H3%
Const mc00F4 = 1 ' &H1%
Const mc0142 = 2 ' &H2%
Const mc0150 = -20480 ' &HB000%
Const mc0152 = 10 ' &HA%
Const mc0154 = 32 ' &H20%
Const m015A = 0 ' &H0%
Const mc015C = 9 ' &H9%
Const mc018E = 6 ' &H6%
Const mc0196 = 63 ' &H3F%
Const VariantType = 10 ' &HA%
Const mc019A = 1 ' &H1%
Const mc019C = 5 ' &H5%
Const StringType = 28 ' &H1C%
Const mc01A0 = 29 ' &H1D%
Const mc01A2 = 9 ' &H9%
Const mc01A4 = 11 ' &HB%
Const mc01A6 = 64 ' &H40%
Const mc01A8 = 192 ' &HC0%
Const mc01AA = 12288 ' &H3000%
Const mc01AC = 112 ' &H70%
Const mc01B2 = 5 ' &H5%
Const mc01B4 = 128 ' &H80%
Const mc01B6 = 256 ' &H100%
Const mc01B8 = 4 ' &H4%
Const mc01BA = 7 ' &H7%
Const mc01BC = 8192 ' &H2000%
Const mc01EA = 522 ' &H20A%
Const mc01EC = 554 ' &H22A%
Const mc01F2 = 12 ' &HC%
Const mc01F4 = 16 ' &H10%
Const mc01F6 = 2 ' &H2%
Const mc01F8 = 64 ' &H40%
Const mc01FA = 15 ' &HF%
Const mc0206 = 5 ' &H5%
Const mc022E = 3 ' &H3%
Const mc026C = 32 ' &H20%
Const mc026E = 4 ' &H4%
Const mc0270 = 6 ' &H6%
Const mc0272 = 4 ' &H4%
Const mc0284 = 1 ' &H1%
Const mc0286 = 2 ' &H2%
Const mc02A0 = -32766 ' &H8002%
Const mc02A2 = -32767 ' &H8001%
Const mc02A4 = 4096 ' &H1000%
Const mc02A6 = 8192 ' &H2000%
Const mc02A8 = 16384 ' &H4000%
Const mc02B4 = 1 ' &H1%
Const mc02B6 = 2 ' &H2%
Const mc02B8 = 3 ' &H3%
Const mc02BA = 4 ' &H4%
Const mc02BC = 18 ' &H12%
Const mc02D2 = 4 ' &H4%
Const mc02D4 = 3 ' &H3%
Const mc02DE = 6 ' &H6%
Const mc02E0 = 7 ' &H7%
Const mc02E2 = 8 ' &H8%
Const mc0302 = 3 ' &H3%

Function fn0A27(p0164 As String) As Integer
Dim l0166
  For l0166 = gv38EC To 1 Step -1
    If gv38BA(l0166).Name = p0164 Then fn0A27 = l0166: Exit Function
  Next
End Function

Private Function fn0A97(ByVal p020C As String) As Long
Dim l020E As Long
Dim l0210 As Integer
Dim l0212 As String
Dim l0214 As Integer
Dim l0216 As Long
  l0210 = InStr(p020C, " Then")
  If l0210 Then
    p020C = Trim$(Left$(p020C, l0210))
  Else
    p020C = Trim$(p020C)
  End If
  While Len(p020C)
    l0210 = InStr(p020C, " ")
    If l0210 Then
      l0212 = Left$(p020C, l0210 - 1): p020C = Trim$(Mid$(p020C, l0210 + 1))
    Else
      l0212 = p020C: p020C = ""
    End If
    For l0210 = m005C To 1 Step -1
      If l0212 = m002A(l0210).Name Then
        l0216 = m002A(l0210).M4106
        Exit For
      End If
    Next
    If l0210 = 0 Then
      l0216 = Val(l0212)
    End If
    Select Case l0214
    Case 0: l020E = l0216
    Case 1: l020E = l020E And l0216
    Case 2: l020E = l020E Or l0216
    Case 3: l020E = l020E + l0216
    Case Else
      ErrBeep
    End Select
    l0210 = InStr(p020C, " ")
    If l0210 Then
      Select Case Left$(p020C, l0210 - 1)
      Case "And": l0214 = 1
      Case "Or": l0214 = 2
      Case "+": l0214 = 3
      Case Else
        MsgBox p020C, , "pre-exec how?"
      End Select
      p020C = Trim$(Mid$(p020C, l0210 + 1))
    End If
  Wend
  fn0A97 = l020E
End Function

Function fn0B6A(p02B0 As T42C7) As Integer
Dim l02B2 As Integer
' Const mc02B4 = 1 ' &H1%
' Const mc02B6 = 2 ' &H2%
' Const mc02B8 = 3 ' &H3%
' Const mc02BA = 4 ' &H4%
' Const mc02BC = 18 ' &H12%
  If p02B0.M36B3_count And mc0154 Then
    If p02B0.M36B3_count And mc01F4 Then
      fn0B6A = mc0206
      Exit Function
    End If
  End If
  Select Case p02B0.M36B3_count And mc0196
  Case StringType: l02B2 = mc02B4
  Case VariantType: l02B2 = mc02B6
  Case mc01A0: l02B2 = mc02B8
  Case mc01F6
If gIsVB_Ver Then
    If p02B0.M36B3_count And mc0150 Then
      l02B2 = mc0152
    ElseIf p02B0.M42D1 <= mc01F8 Then
      l02B2 = 0
    Else
      DbgMileStone "string too big -> ptr"
      l02B2 = mc02BA
    End If
Else
    If p02B0.M36B3_count And mc0150 Then
      l02B2 = mc0152
    End If
End If
  Case mc02BC
    l02B2 = mc02BA
  Case mc01A4
    l02B2 = gv38BA(p02B0.M42D1).M42D9
  End Select
  fn0B6A = l02B2
End Function

Function fn0B77(p02C0 As String, p02C2 As String) As Integer
  If Len(p02C2) = 0 Then Exit Function
  If InStr(p02C0, ":") = 0 Then
    p02C0 = g_ExeFilePath & p02C0
    If Right$(p02C0, 1) <> "\" Then p02C0 = p02C0 & "\"
  End If
  fn0B77 = Len(Dir$(p02C0 & p02C2))
End Function

Function fn0B84(p02C8 As Integer, ByVal pv02CA As Integer, p02CE As T42C7) As Long
Dim l02D0 As Integer
' Const mc02D2 = 4 ' &H4%
' Const mc02D4 = 3 ' &H3%
If gIsVB_Ver Then
  If p02CE.M36B3_count And mc0150 Then
    Select Case p02CE.M36B3_count And mc01FA
    Case mc019A
      l02D0 = 1
    Case mc01F6
      l02D0 = 2
      If pv02CA And 1 Then DbgMileStone "expected even $*"
    Case mc02D2, mc02D4
      l02D0 = 2
    Case mc01A4
      l02D0 = gv38BA(p02CE.M42D1).Size
    Case Else
      l02D0 = 4
    End Select
  ElseIf p02CE.M36B3_count And mc01A8 Then
    l02D0 = 4
  ElseIf (p02CE.M36B3_count And mc0196) = mc01F6 Then
    l02D0 = 2
  Else
    l02D0 = fn0B8F(p02CE)
  End If
  If pv02CA < 0 Then
    p02C8 = p02C8 + pv02CA: pv02CA = 0
    If l02D0 >= 4 Then
      If p02C8 And 3 Then
        p02C8 = p02C8 And -4
      End If
    ElseIf l02D0 >= 2 Then
      If p02C8 And 1 Then
        p02C8 = p02C8 And -2
      End If
    End If
  Else
    If l02D0 >= 4 Then
      If p02C8 And 3 Then
        p02C8 = (p02C8 + 3) And -4
      End If
    ElseIf l02D0 >= 2 Then
      If p02C8 And 1 Then
        p02C8 = (p02C8 + 1) And -2
      End If
    End If
  End If
Else
  If pv02CA < 0 Then
    p02C8 = p02C8 + pv02CA: pv02CA = 0
  ElseIf p02CE.M36B3_count And mc01A6 Then
    pv02CA = 4
  End If
End If
  fn0B84 = p02C8
  p02C8 = p02C8 + pv02CA
End Function

Function fn0B8F(p02DA As T42C7) As Integer
Dim l02DC As Integer
' Const mc02DE = 6 ' &H6%
' Const mc02E0 = 7 ' &H7%
' Const mc02E2 = 8 ' &H8%
  If p02DA.M36B3_count And mc01AC Then
    l02DC = 4
  Else
    Select Case p02DA.M36B3_count And mc01FA
    Case mc019A: l02DC = 1
    Case mc02D2, mc02D4: l02DC = 2
    Case mc019C, mc02DE: l02DC = 4
    Case mc02E0, mc02E2, mc01A2: l02DC = 8
    Case VariantType: l02DC = 16
    Case mc01A4: l02DC = gv38BA(p02DA.M42D1).Size
    Case mc01F6
If gIsVB_Ver Then
      l02DC = p02DA.M42D1 * 2
Else
      l02DC = p02DA.M42D1
End If
    Case Else
      ErrBeep
    End Select
  End If
  fn0B8F = l02DC
End Function

Function fn0B9D(p02E6 As Integer) As String
Dim l02E8 As Integer
Dim l02EA As String
  If gv097C Then
    l02EA = gv094A(p02E6).M3F0B
    l02E8 = FindLastOccurenceOf(l02EA, "\")
    If l02E8 Then
      fn0B9D = Mid$(l02EA, l02E8 + 1)
    Else
      fn0B9D = l02EA
    End If
  ElseIf p02E6 > gEXEdef1 Then
    fn0B9D = "form" & p02E6 - gEXEdef1
  Else
    fn0B9D = "module" & p02E6
  End If
End Function

Private Sub sub09AB(p00D6 As String)
Dim l00D8 As Integer
' Const mc00DC = 2 ' &H2%
  l00D8 = InStr(p00D6, ";")
  If l00D8 Then
    p00D6 = Trim$(Mid$(p00D6, l00D8 + 1))
  End If
  gv0980 = gv0980 + 1
If gv0B1C Then
  sub09B8 p00D6, mc00DC
Else
  If gv0980 > UBound(m0070) Then ReDim Preserve m0070(gv0980 + 20)
  m0070(gv0980).Name = p00D6
  m0070(gv0980).M4106 = mc00DC
End If
End Sub

Sub sub09B8(ByVal p00FE As String, p0100 As Integer)
  gv097C = gv097C + 1
  If gv097C > UBound(gv094A) Then ReDim Preserve gv094A(gv097E + 20)
  gv094A(gv097C).M3F0B = p00FE
  gv094A(gv097C).M36B3_count = p0100
End Sub

Private Sub sub09C4(p00DE As String)
Dim l00E0 As Integer
Dim l00E2 As String
Dim l00E4 As Integer
' Const mc00E8 = 16 ' &H10%
' Const mc00EC = 3 ' &H3%
  l00E4 = InStr(p00DE, ";")
  If l00E4 Then p00DE = Trim$(Mid$(p00DE, l00E4 + 1))
  l00E0 = FreeFile
  If InStr(p00DE, "\") Then
    Open p00DE For Input As l00E0
  Else
    Open g_ExeFilePath & p00DE For Input As l00E0
  End If
  Line Input #l00E0, l00E2
  Close l00E0
  If Left$(l00E2, 8) <> "VERSION " Then
    MsgBox p00DE & Chr$(13) & l00E2, mc00E8, "not a form"
    sub09D0 p00DE
    Exit Sub
  End If
  gv0982 = gv0982 + 1
If gv0B1C Then
  sub09B8 p00DE, mc00EC
Else
  If gv0982 > UBound(m00A4) Then ReDim Preserve m00A4(gv0982 + 20)
  m00A4(gv0982).Name = p00DE
  m00A4(gv0982).M4106 = mc00EC
End If
End Sub

Private Sub sub09D0(p00EE As String)
Dim l00F0 As Integer
' Const mc00F4 = 1 ' &H1%
  l00F0 = InStr(p00EE, ";")
  If l00F0 Then p00EE = Trim$(Mid$(p00EE, l00F0 + 1))
  gv097E = gv097E + 1
If gv0B1C Then
  sub09B8 p00EE, mc00F4
Else
  If gv097E > UBound(gv094A) Then ReDim Preserve gv094A(gv097E + 20)
  gv094A(gv097E).M3F0B = p00EE
  gv094A(gv097E).M36B3_count = mc00F4
End If
End Sub

Private Sub sub09DB(p0104 As Integer, p0106 As String, p0108 As Integer)
Dim l0114 As Integer
  gv0B08 = gv0B08 + 1
  If gv0B08 > UBound(gv0AD6) Then ReDim Preserve gv0AD6(gv0B08 + 20)
  gv0AD6(gv0B08).M403B = p0104
  If Left$(p0106, 1) = "@" Then
    l0114 = InStr(p0106, " ")
    gv0AD6(gv0B08).Name = Mid$(p0106, l0114 + 1)
  Else
    gv0AD6(gv0B08).Name = p0106
  End If
  gv0AD6(gv0B08).Flags = p0108
End Sub

Private Sub sub09E8(p0116 As String)
Dim l0118 As Integer
Dim l011A As String
  l0118 = InStr(p0116, "'")
  If l0118 Then p0116 = Trim$(Left$(p0116, l0118 - 1))
  m005C = m005C + 1
  If m005C > UBound(m002A) Then ReDim Preserve m002A(m005C + 10)
  l0118 = InStr(p0116, "=")
  m002A(m005C).Name = Trim$(Left$(p0116, l0118 - 1))
  l011A = Trim$(Mid$(p0116, l0118 + 1))
  If IsNumeric(l011A) Then
    m002A(m005C).M4106 = Val(l011A)
  Else
    For l0118 = m005C To 1 Step -1
      If l011A = m002A(l0118).Name Then
        m002A(m005C).M4106 = m002A(l0118).M4106
        Exit For
      End If
    Next
    If l0118 = 0 Then MsgBox l011A, , "eval?"
  End If
End Sub

Sub sub09F4()
  frmScan.control4 = "User Defined"
  frmScan.control3 = "Type"
  sub0A01
End Sub

Sub sub0A01()
Dim l0120
Dim l0122 As Integer
Dim l0124 As Integer
Dim l0126 As String
Dim l0128 As String
Dim l012A As Integer
Dim l012C As Integer
Dim l012E%
Dim l0130 As Integer
Dim l0134 As T42C7
Dim l0136 As Integer
' Const mc0142 = 2 ' &H2%
' Const mc0150 =-20480 ' &HB000%
' Const mc0152 = 10 ' &HA%
' Const mc0154 = 32 ' &H20%
' Const m015A = 0 ' &H0%
' Const mc015C = 9 ' &H9%
  l0124 = FreeFile
  Do
    l012A = False
    For l0120 = 1 To gv38EC
      l0136 = 0
      If gv38BA(l0120).Size = 0 Then
        frmScan.control2 = gv38BA(l0120).Name
        DoEvents
        l0124 = fn171A(gv38BA(l0120).M403B, 0, mc0142)
        Seek l0124, gv38BA(l0120).hFile
        gv38BA(l0120).M4950 = gv3922
        l012C = 0: l0130 = 0
        Do Until EOF(l0124)
          Line Input #l0124, l0126
          l0126 = Trim$(l0126)
          l0122 = InStr(l0126, "'")
          If l0122 Then l0126 = Trim$(Left$(l0126, l0122 - 1))
          If Len(l0126) Then
            If l0126 = "End Type" Then Exit Do
            gv3922 = gv3922 + 1
            If gv3922 > UBound(gv38F0) Then ReDim Preserve gv38F0(gv3922 + 10)
            l0134.M36B3_count = mc0150
            sub0A10 l0126, l0128, l0134, l012E
            l0136 = l0136 Or l0134.M42D9
            If l0134.M42D9 = mc0152 Then
              l0136 = l0136 Or &H8000
            ElseIf l0136 < 0 Then
              If l0134.M36B3_count And mc0154 Then
                l0134.M42D9 = 5
              End If
            End If
            gv38F0(gv3922).M6894 = l0134
            If l012E = 0 Then
              l012A = True
              gv3922 = gv38BA(l0120).M4950
              Exit Do
            End If
If gIsVB_Ver Then
            gv38F0(gv3922).hFile = fn0B84(l012C, l012E, l0134)
Else
            gv38F0(gv3922).hFile = l012C
            l012C = l012C + l012E
End If
            QuoteWithParag l0128, gv3606, gv38F0(gv3922).StringBuffOffs
            gv38F0(gv3922).Size = l012E
            l0130 = l0130 + 1
          End If
        Loop
        l0124 = fn171A(-1, 0, m015A)
        If l012E Then
          gv38BA(l0120).M4959 = l0130
          gv38BA(l0120).Size = l012C
          If l0136 Then gv38BA(l0120).M42D9 = mc015C
        End If
      End If
    Next
  Loop While l012A
End Sub

Sub sub0A10(p01D6 As String, p01D8 As String, p01DC As T42C7, p01DE As Integer)
Dim l01E0 As Integer
Dim l01E2 As Integer
Dim l01E4 As String
Dim l01E6 As String
Dim l01E8 As Integer
' Const mc01EA = 522 ' &H20A%
' Const mc01EC = 554 ' &H22A%
' Const mc01F2 = 12 ' &HC%
' Const mc01F4 = 16 ' &H10%
' Const mc01F6 = 2 ' &H2%
' Const mc01F8 = 64 ' &H40%
' Const mc01FA = 15 ' &HF%
Dim l0200 As Integer
Dim l0202 As Integer
Dim l0204 As String
' Const mc0206 = 5 ' &H5%
  l01E0 = InStr(p01D6, ",")
  If l01E0 Then
    l01E2 = InStr(Left$(p01D6, l01E0), "(")
    If l01E2 Then
      l01E2 = InStr(p01D6, ")")
      l01E0 = InStr(l01E2, p01D6, ",")
    End If
  End If
  If l01E0 Then
    p01D8 = RTrim$(Left$(p01D6, l01E0 - 1))
    p01D6 = LTrim$(Mid$(p01D6, l01E0 + 1))
  Else
    p01D8 = Trim$(p01D6)
    p01D6 = ""
  End If
  If Left$(p01D8, 9) = "Optional " Then
    p01D8 = Mid$(p01D8, 10)
    p01DC.M36B3_count = mc01EA
  Else
  End If
  If Left$(p01D8, 5) = "ByVal" Then
    p01DC.M36B3_count = p01DC.M36B3_count Or mc01B4
    p01D8 = Mid$(p01D8, 7)
  ElseIf Left$(p01D8, 5) = "ByRef" Then
    p01DC.M36B3_count = p01DC.M36B3_count Or mc01A6
    p01D8 = Mid$(p01D8, 7)
  ElseIf p01DC.M36B3_count And mc01EA Then
    p01DC.M36B3_count = p01DC.M36B3_count Or mc01A6
  End If
  If Left$(p01D8, 11) = "ParamArray" Then
    p01DC.M36B3_count = mc01EC
    p01D8 = Mid$(p01D8, 12)
  End If
  If Left$(p01D8, 1) = "@" Then
    DbgMileStone p01D8
    l01E0 = InStr(p01D8, " ")
    p01D8 = Mid$(p01D8, l01E0 + 1)
  End If
  l01E2 = InStr(p01D8, " As ")
  If l01E2 Then
    l01E4 = Trim$(Mid$(p01D8, l01E2 + 4))
    l01E8 = InStr(gc2A0E, l01E4 & "§")
    If l01E8 Then
      l01E8 = InStr(gc2A0A, Mid$(gc2A0E, l01E8 - 1, 1))
    End If
    If l01E8 Then
      If l01E8 >= mc01F2 Then l01E8 = l01E8 Or mc01F4
    Else
      If Left$(l01E4, 8) = "String *" Then
        l01E4 = Mid$(l01E4, 9)
        If IsNumeric(l01E4) Then
          p01DC.M42D1 = Val(l01E4)
        Else
          For l01E0 = m005C To 1 Step -1
            If l01E4 = m002A(l01E0).Name Then
              Exit For
            End If
          Next
          If l01E0 Then
            p01DC.M42D1 = m002A(l01E0).M4106
          Else
            l01E4 = InputBox$("Enter the Value of this Expression", l01E4)
            p01DC.M42D1 = Val(l01E4)
          End If
        End If
If gIsVB_Ver Then
        If p01DC.M36B3_count And mc0150 Then
          l01E8 = mc01F6
        ElseIf p01DC.M42D1 <= mc01F8 Then
          l01E8 = mc01F6
        Else
          l01E8 = mc01F6 Or mc01F4
        End If
Else
        l01E8 = mc01F6
End If
      Else
        l01E0 = fn0A27(l01E4)
        If l01E0 Then
          p01DC.M42D1 = l01E0
          l01E8 = mc01A4
        Else
          l01E8 = mc01A0
        End If
      End If
    End If
    p01DC.M36B3_count = p01DC.M36B3_count Or l01E8
    p01D8 = Trim$(Left$(p01D8, l01E2 - 1))
    GoSub L15BB0
    p01D8 = p01D8 & Mid$(gc2A0A, p01DC.M36B3_count And mc01FA, 1)
  Else
    GoSub L15BB0
    l01E8 = InStr(gc2A0A, Right$(p01D8, 1))
    If l01E8 = 0 Then
      l01E8 = Asc(Mid$(gv0988, Asc(UCase$(p01D8)), 1))
      p01D8 = p01D8 & Mid$(gc2A0A, l01E8 And mc01FA, 1)
    ElseIf l01E8 >= mc01A4 Then
      l01E8 = l01E8 Or mc01F4
    End If
    p01DC.M36B3_count = p01DC.M36B3_count Or l01E8
  End If
  p01DE = fn0B8F(p01DC)
  p01DC.M42D9 = fn0B6A(p01DC)
  If Len(l01E6) = 0 Then
    Exit Sub
  End If
  p01DC.M36B3_count = p01DC.M36B3_count Or mc0154
  p01D8 = p01D8 & l01E6
  If l01E6 = "()" Then
    p01DE = 4
    p01DC.M42D9 = &H105
    Exit Sub
  End If
  If p01DC.M36B3_count And mc0150 Then
    If p01DC.M42D9 Then
      p01DC.M42D9 = mc0206
    Else
      p01DC.M42D9 = 0
    End If
    l01E0 = fn0B84(p01DE, 0, p01DC)
    l01E6 = Mid$(l01E6, 2, Len(l01E6) - 2) & ","
    Do
      l01E0 = InStr(l01E6, ",")
      l0204 = Left$(l01E6, l01E0 - 1): l01E6 = LTrim$(Mid$(l01E6, l01E0 + 1))
      l01E0 = InStr(l0204, "To ")
      If l01E0 Then
        l0200 = Val(l0204)
        l0202 = Val(Mid$(l0204, l01E0 + 3))
      Else
        l0202 = Val(l0204)
        l0200 = gv0984 And 1
      End If
      p01DE = p01DE * (l0202 - l0200 + 1)
    Loop While Len(l01E6)
  Else
    p01DC.M42D9 = mc0206
    p01DE = &H10
    l01E0 = 0
    Do
      p01DE = p01DE + 8
      l01E0 = InStr(l01E0 + 1, l01E6, ",")
    Loop While l01E0
  End If
Exit Sub

L15BB0:
  l01E2 = InStr(p01D8, "(")
  If l01E2 = 0 Then Return
  l01E6 = Mid$(p01D8, l01E2)
  p01D8 = Left$(p01D8, l01E2 - 1)
Return
End Sub

Sub sub0A1D(ByVal p1$, ByVal p2$)
Stop
End Sub

Sub sub0A33(p0168%, p016A$, p016C%)
Dim l016E As String
Dim l0172 As T42C7
Dim l0174%
  sub0A10 p016A, l016E, l0172, l0174
  sub0A3E p0168, l016E, l0172, l0174, p016C
End Sub

Sub sub0A4C(p0176 As Integer)
  If p0176 = gv16B6 Then Exit Sub
  sub0A5E p0176
End Sub

Sub sub0A6F(p017A%, p017C As String, p017E As Integer, p0180 As Integer)
Dim l0186 As String
Dim l018A As T42C7
Dim Size%
' Const mc018E = 6 ' &H6%
' Const mc0196 = 63 ' &H3F%
' Const VariantType = 10 ' &HA%
' Const mc019A = 1 ' &H1%
' Const mc019C = 5 ' &H5%
' Const StringType = 28 ' &H1C%
' Const mc01A0 = 29 ' &H1D%
' Const mc01A2 = 9 ' &H9%
' Const mc01A4 = 11 ' &HB%
' Const mc01A6 = 64 ' &H40%
' Const mc01A8 = 192 ' &HC0%
' Const mc01AA = 12288 ' &H3000%
' Const mc01AC = 112 ' &H70%
' Const mc01B2 = 5 ' &H5%
' Const mc01B4 = 128 ' &H80%
' Const mc01B6 = 256 ' &H100%
' Const mc01B8 = 4 ' &H4%
' Const mc01BA = 7 ' &H7%
' Const mc01BC = 8192 ' &H2000%
  Select Case p017E
  Case mc018E
If gIsVB_Ver Then
      gv371C = &HC
Else
      gv371C = &HA
End If
    gv371E = gv2FD8
    If Len(p017C) Then
      sub0A10 p017C, l0186, l018A, Size
      If p0180 = mc00F4 Then
If gIsVB_Ver Then
        Select Case l018A.M36B3_count And mc0196
        Case VariantType
          gv371C = &H10
        End Select
Else
        Select Case l018A.M36B3_count And mc0196
        Case mc019A To mc019C, StringType, mc01A0
        Case Else
          gv371C = &HC
        End Select
End If
        Select Case l018A.M36B3_count And mc0196
        Case mc019A To mc01A2
        Case StringType
          DbgVar Size = 4, "string size"
        Case VariantType
          DbgVar Size = 16, "variant size"
        Case mc01A4
        Case mc01A0
        Case Else
          l018A.M36B3_count = l018A.M36B3_count Or mc01A6
        End Select
        If (l018A.M36B3_count And mc01A8) = 0 Then
          l018A.M36B3_count = l018A.M36B3_count Or mc01AA
        ElseIf l018A.M36B3_count And mc01AC Then
          DbgVar Size = 4, "ptr size"
        End If
        sub0A7A p017A, l0186, l018A, Size
      Else
        l018A.M36B3_count = l018A.M36B3_count Or mc01A6
        sub0A7A p017A, l0186, l018A, 4
      End If
    Else
    End If
    gv3718 = gv371E + gv3718
    gv371A = gv371C + gv371A
    sub0A88 p017A, p0180
  Case mc01B2
    sub0A10 p017C, l0186, l018A, Size
    If l018A.M36B3_count And mc01B4 Then
      Select Case l018A.M36B3_count And mc0196
      Case StringType, VariantType
        l018A.M36B3_count = l018A.M36B3_count Or mc01B6
        sub0A7A p017A, l0186, l018A, Size
      Case mc01A0
        sub0A7A p017A, l0186, l018A, Size
      Case mc01A4
        DbgMileStone "UDT always ByRef"
        l018A.M36B3_count = (l018A.M36B3_count And Not mc01B4) Or mc01A6
        sub0A7A p017A, l0186, l018A, 4
      Case Else
        sub0A7A p017A, l0186, l018A, Size
      End Select
    Else
      l018A.M36B3_count = l018A.M36B3_count Or mc01A6
      sub0A7A p017A, l0186, l018A, 4
    End If
  Case mc01B8
    sub0A10 p017C, l0186, l018A, Size
    l018A.M36B3_count = l018A.M36B3_count Or mc01AA
    sub0A7A p017A, l0186, l018A, Size
  Case mc01BA
    sub0A10 p017C, l0186, l018A, Size
    l018A.M36B3_count = l018A.M36B3_count Or mc01BC
    sub0A7A p017A, l0186, l018A, Size
  End Select
End Sub

Sub sub0A88(p01BE As Integer, p01C0 As Integer)
Dim l01C2%
Dim l01C6 As T42C7
Dim l01C8 As String
  If p01C0 = mc00F4 And Not gIsVB_Ver Then
    For l01C2 = gv36E4 To 1 Step -1
      l01C6 = gv36AE(l01C2).M6894
      If l01C6.M36B3_count And mc01B6 Then
        l01C8 = fn1B6C(l01C2)
        l01C6.M36B3_count = (l01C6.M36B3_count And Not (mc01A8 Or mc01B6)) Or mc01AA
        If (l01C6.M36B3_count And mc01AC) = 0 Then
          If Left$(l01C8, 1) = "^" Then l01C8 = Mid$(l01C8, 2)
        End If
        sub0A7A p01BE%, l01C8, l01C6, fn0B8F(l01C6)
      End If
    Next
  Else
    For l01C2 = 1 To gv36E4
      l01C6 = gv36AE(l01C2).M6894
      If l01C6.M36B3_count And mc01B6 Then
        l01C8 = fn1B6C(l01C2)
        l01C6.M36B3_count = (l01C6.M36B3_count And Not (mc01A8 Or mc01B6)) Or mc01AA
        If (l01C6.M36B3_count And mc01AC) = 0 Then
          If Left$(l01C8, 1) = "^" Then l01C8 = Mid$(l01C8, 2)
        End If
        sub0A7A p01BE, l01C8, l01C6, fn0B8F(l01C6)
      End If
    Next
  End If
End Sub

Private Sub sub0AA3(p0218%)
Dim l021A As String
Dim l021C As String
  Do
    Line Input #m0060, l021A
    l021C = LTrim$(l021A)
    If Left$(l021C, 1) = "#" Then
      sub0AAF l021C, p0218
    Else
      If p0218 <= mc0068 Then
        Print #m0062, RTrim$(l021A)
      End If
    End If
  Loop While p0218
End Sub

Private Sub sub0AAF(ByVal p021E As String, p0220 As Integer)
  Select Case InStr(p021E, " ")
  Case 0, 6
    If gv01B4 Then sub0A1D p021E, "#Else"
    Select Case p0220
    Case mc0066
      MsgBox p021E, , "missing #If"
      sub0AA3 mc006C
    Case mc0068
      p0220 = mc006C
    Case mc006A
      p0220 = mc0068
    Case mc006C
    Case Else
      ErrBeep
    End Select
  Case 4
    If gv01B4 Then sub0A1D p021E, "#If "
    Select Case p0220
    Case mc0066, mc0068
      If fn0A97(Mid$(p021E, 5)) Then
        sub0AA3 mc0068
      Else
        sub0AA3 mc006A
      End If
    Case mc006A, mc006C
      sub0AA3 mc006C
    Case Else
      ErrBeep
    End Select
  Case 5
    If gv01B4 Then sub0A1D p021E, "#End If"
    p0220 = m0064
  Case 7
    If gv01B4 Then sub0A1D p021E, "#Const "
    Select Case p0220
    Case mc0066, mc0068
      sub09E8 Mid$(p021E, 8)
    End Select
  Case 8
    If gv01B4 Then sub0A1D p021E, "#ElseIf "
    Select Case p0220
    Case mc0066
      MsgBox p021E, , "missing #If"
      sub0AA3 mc006C
    Case mc0068
      p0220 = mc006C
    Case mc006A
      If fn0A97(Mid$(p021E, 10)) Then
        p0220 = mc0068
      End If
    Case mc006C
    Case Else
      ErrBeep
    End Select
  Case Else
    MsgBox p021E, , "exec?"
  End Select
End Sub

Private Sub sub0ABC(p0224 As Integer)
Dim l0228 As String
Dim l022A As String
Dim l022C As Integer
' Const mc022E = 3 ' &H3%
  m005C = m005E
  m0060 = FreeFile
  l0228 = gv094A(p0224).M3F0B
  If InStr(l0228, ":") = 0 Then l0228 = g_ExeFilePath & l0228
  Open l0228 For Input As m0060
  m0062 = fn171A(p0224, 0, mc022E)
  l022C = mc0066
  While Not EOF(m0060)
    Line Input #m0060, l0228
    l022A = LTrim$(l0228)
    If Left$(l022A, 1) = "#" Then
      sub0AAF l022A, l022C
    Else
      Print #m0062, RTrim$(l0228)
    End If
  Wend
  Close m0060
  m0062 = fn171A(p0224, 0, m015A)
End Sub

Private Sub sub0AC8(p0230 As String)
Dim l0232 As Integer
Dim l0234 As Integer
Dim l0236 As String
Dim l0238 As Integer
Dim l023A As Integer
  l0232 = InStr(gc2A12, Mid$(p0230, 4, 3)) \ 3
  DbgVar l0232, "Def<typ>"
  If l0232 = mc01F2 Then l0232 = StringType
  l0234 = InStr(p0230, " ")
  l0236 = UCase$(LTrim$(Mid$(p0230, l0234 + 1)))
  l0234 = InStr(l0236, "'")
  If l0234 Then l0236 = RTrim$(Left$(l0236, l0234 - 1))
  If l0236 = "A-Z" Then
    gv0988 = String$(255, l0232)
  Else
    Do While Len(l0236)
      l0238 = Asc(l0236)
      If Mid$(l0236, 2, 1) = "-" Then
        l023A = Asc(Mid$(l0236, 3, 1)) - l0238 + 1
        l0236 = Mid$(l0236, 5)
      Else
        l0236 = Mid$(l0236, 3)
        l023A = 1
      End If
      Mid$(gv0988, l0238, l023A) = String$(l023A, l0232)
    Loop
  End If
End Sub

Sub sub0AD4()
Dim l0240 As Integer
Dim l0244 As Integer
Dim l0246 As String
Dim l0248 As String
' Const mc026C = 32 ' &H20%
' Const mc026E = 4 ' &H4%
' Const mc0270 = 6 ' &H6%
' Const mc0272 = 4 ' &H4%
Dim l027C%
' Const mc0284 = 1 ' &H1%
' Const mc0286 = 2 ' &H2%
  frmScan.control4 = "MakeFile"
  frmScan.SetFocus
  gv_ExeFilePath = g_ExeFilePath: gExeFileNameOnly2 = ""
  m001C = g_ExeFilePath: m0018 = ""
  m0024 = g_ExeFilePath: m0020 = ""
  l0246 = gv02EC_MAKName
  l0240 = FindLastOccurenceOf(l0246, "\")
  If l0240 Then l0246 = Mid$(l0246, l0240 + 1)
  frmScan.control2 = l0246
  DoEvents
  gv16BA = 0: gv16BC = 0
  gv097E = 0: gv097C = 0: m005C = 0: gv0982 = 0
  gv0980 = 0
  gv0B1C = False
  ReDim m002A(10), m0070(10)
  ReDim m00A4(10)
  gv3676 = 0: ReDim gv3640(10)
  gv360A = 0: ReDim gv35D4(10)
  sub0AE5
  l0244 = FreeFile
  Open gv02EC_MAKName For Input As #l0244
  Do While Not EOF(l0244)
    Line Input #l0244, l0246
    Select Case Left$(l0246, 4)
    Case "Form", "Modu", "Clas"
      gv097C = gv097C + 1
    Case "Star"
      If Left$(l0246, 7) = "Startup" Then
        gv0B1C = True
        Exit Do
      ElseIf Left$(l0246, 9) = "StartMode" Then
      Else
        MsgBox l0246, , "handle?"
      End If
    End Select
  Loop
  Seek #l0244, 1
  ReDim gv094A(gv097C)
  gv097C = 0
  While Not EOF(l0244)
    Line Input #l0244, l0246
    l0240 = InStr(l0246, "=")
    If l0240 Then
      If Mid$(l0246, l0240 + 1, 1) = """" Then
        l0248 = Mid$(l0246, l0240 + 2)
        l0248 = Left$(l0248, Len(l0248) - 1)
      Else
        l0248 = Mid$(l0246, l0240 + 1)
      End If
      Select Case Left$(l0246, l0240 - 1)
      Case "ExeName"
        m0018 = l0248
      Case "Path"
        m001C = l0248
      Case "ExeName32"
        m0020 = l0248
      Case "Path32"
        m0024 = l0248
      Case "Form"
        sub09C4 l0248
      Case "Module"
        sub09D0 l0248
      Case "Class"
        sub09AB l0248
      Case "VBX"
      Case "CondComp"
        sub09E8 l0248
      End Select
    Else
      l0240 = FindLastOccurenceOf(l0246, ".")
      If l0240 > Len(l0246) - 4 Then
        Select Case UCase$(Right$(l0246, 4))
        Case ".BAS": sub09D0 l0246
        Case ".FRM": sub09C4 l0246
        Case ".EXE"
        Case ".VBX", ".OCX"
        Case ".TXT": sub09C4 l0246
        Case ".DLL"
        Case ".CLS": sub09AB l0246
        Case ".CTL"
          MsgBox l0246, , "unsupported type"
        Case Else
          MsgBox l0246, , "make what?"
        End Select
      Else
          MsgBox l0246, , "make what?"
      End If
    End If
  Wend
  Close l0244
If gv0B1C Then
Else
    gv097C = gv097E + gv0980 + gv0982
    ReDim Preserve gv094A(gv097C)
    For l0240 = 1 To gv0980
      gv094A(l0240 + gv097E).M3F0B = m0070(l0240).Name
      gv094A(l0240 + gv097E).M36B3_count = m0070(l0240).M4106
    Next
    For l0240 = 1 To gv0982
      gv094A(l0240 + gv097E + gv0980).M3F0B = m00A4(l0240).Name
      gv094A(l0240 + gv097E + gv0980).M36B3_count = m00A4(l0240).M4106
    Next
End If
  g_IsNE = fn0B77(m001C, m0018)
  gIsVB_Ver = fn0B77(m0024, m0020)
  If g_IsNE And gIsVB_Ver Then
    If MsgBox("16 Bit", mc026C Or mc026E) = mc0270 Then
      g_IsNE = True
      gIsVB_Ver = False
    Else
      g_IsNE = False
      gIsVB_Ver = True
    End If
  End If
  sub09E8 "True = -1"
  sub09E8 "False = 0"
If gIsVB_Ver Then
  sub09E8 "Win16 = False"
  sub09E8 "Win32 = True"
  gv_ExeFilePath = m0024
  gExeFileNameOnly2 = m0020
ElseIf g_IsNE Then
  sub09E8 "Win16 = True"
  sub09E8 "Win32 = False"
  gv_ExeFilePath = m001C
  gExeFileNameOnly2 = m0018
Else
  ShowMsgBox mc0272
End If
  m005E = m005C
  G_FullExeFileName = gv_ExeFilePath & gExeFileNameOnly2
  l0240 = InStr(gExeFileNameOnly2, ".")
  If l0240 Then gExeFileNameOnly2 = Left$(gExeFileNameOnly2, l0240 - 1)
  sub0AF5
  If Not gv16B0 Then Exit Sub
  gv0B08 = 0
  gv0B0A = 0: gv0B0C = 0
  ReDim gv0AD6(20)
  frmScan.control3 = "Module"
  For l027C = 1 To gv097C
    sub0B04 l027C
  Next
  sub09F4
  frmScan.control5 = "Scanning"
  frmScan.control3 = "Module"
  For l027C = 1 To gv097C
    gv0A88 = 0: gv0A8A = 0: gv36AA = 0
    frmScan.control4 = "Declarations"
    frmScan.control2 = fn0B9D(l027C)
    DoEvents
    gv094A(l027C).M3F30 = gv360A
    sub0B13 l027C, mc0284
    gv094A(l027C).M3F3B = gv360A - gv094A(l027C).M3F30
    gv094A(l027C).M3F51 = gv0A88
    gv0A8A = gv0A88
    gv094A(l027C).M3F5B = gv360A
    DbgVar gv3676 = 0, "no statics"
    gv094A(l027C).M3F8A = gv3676
    sub0B13 l027C, mc0286
    gv094A(l027C).M3F7F = gv0A8A
    gv094A(l027C).M3F67 = gv360A - gv094A(l027C).M3F5B
    gv094A(l027C).M3FA2 = gv36AA
    gv094A(l027C).M3F96 = gv3676 - gv094A(l027C).M3F8A
    sub0B23
  Next
  sub0B31
End Sub

Private Sub sub0B04(p0288%)
Dim l028A As Integer
Dim l028C As String
Dim l028E As String
Dim l0290 As String
Dim l0292 As Integer
Dim l0294 As Integer
Dim l0296 As Integer
Dim l0298 As Long
Dim l029A As Integer
Dim l029C As Integer
Dim l029E As Integer
' Const mc02A0 =-32766 ' &H8002%
' Const mc02A2 =-32767 ' &H8001%
' Const mc02A4 = 4096 ' &H1000%
' Const mc02A6 = 8192 ' &H2000%
' Const mc02A8 = 16384 ' &H4000%
  sub0B40 p0288
  l029C = 0
  l029E = False
  frmScan.control2 = fn0B9D(p0288)
  If gv094A(p0288).M3F1B = 0 Then
    frmScan.control5 = "Preprocessing"
    frmScan.control4 = "Conditions"
    DoEvents
    sub0ABC p0288
  End If
  frmScan.control5 = "Scanning"
  frmScan.control4 = "Subroutines"
  DoEvents
  l0294 = 0
  sub09DB p0288, "(form)", mc02A0
  gv094A(p0288).M3F1B = gv0B08
  l0292 = fn171A(p0288, gv0B08, mc0142)
  Select Case gv094A(p0288).M36B3_count
  Case mc00EC
    l029E = True
    Line Input #l0292, l028C
    Assume Left$(l028C, 8) = "VERSION ", "form version"
    Assume Right$(l028C, 5) <> "CLASS", "form module"
    Line Input #l0292, l028C
    l028C = Trim$(l028C)
    l028A = FindLastOccurenceOf(l028C, " ")
    If l028A Then
      gv094A(p0288).M3F23 = Mid$(l028C, l028A + 1)
    Else
      MsgBox l028C, , "expect class name"
    End If
    Do
      Line Input #l0292, l028C
      If Left$(l028C, 3) = "End" Then Exit Do
    Loop
  Case mc00DC
    l029E = True
    Line Input #l0292, l028C
    Assume Left$(l028C, 8) = "VERSION ", "class version"
    Assume Right$(l028C, 5) = "CLASS", "class module"
    Line Input #l0292, l028C
    l028C = Trim$(l028C)
    Assume l028C = "BEGIN", "BEGIN"
    Do
      Line Input #l0292, l028C
    Loop Until l028C = "END"
  End Select
  gv0AD6(gv0B08).M4045 = Seek(l0292)
  l0294 = l0294 + 1
  sub09DB p0288, "(declarations)", mc02A2
  l0298 = Seek(l0292)
  While Not EOF(l0292)
    Line Input #l0292, l028E
    l028C = LTrim$(l028E)
    If Len(l028C) = 0 Then
      l029A = False
    ElseIf Left$(l028C, 1) = "'" Then
      l029A = False
    ElseIf Len(l028C) > 3 Then
      l029A = True
      l0296 = 0
      l028A = InStr(l028C, """")
      If l028A Then l028C = Left$(l028C, l028A - 1)
      l028A = InStr(l028C, "'")
      If l028A Then l028C = Left$(l028C, l028A - 1)
      l028A = InStr("   Sub Fun Pro Pri Pub Sta Rem Typ Def Opt", Left$(l028C, 3)) \ 4
      On l028A GoSub L17DAA, L17DA2, L17D92, L17B6A, L17B6A, L17B6A, L17B06, L17B1A, L17B3A, L17B52
    ElseIf l028C = "Rem" Then
      l029A = False
    End If
    While Right$(l028E, 1) = "_"
      Line Input #l0292, l028E
    Wend
    If l029A Then l0298 = Seek(l0292)
  Wend
  gv094A(p0288).M3F13 = l0294
  If l029E Then
    gv0B0C = gv0B0C + l029C
  Else
    gv0B0A = gv0B0A + l029C
  End If
  gv0AD6(gv0B08).M4045 = Seek(l0292)
  l0292 = fn171A(p0288, gv0B08, m015A)
Exit Sub

L17B06:
  l029A = False
Return

L17B1A:
  sub0B4F p0288, l0292, l028C
Return

L17B3A:
  sub0AC8 l028C
Return

L17B52:
  sub0B5D l028C
Return

L17B6A:
  l028A = InStr(l028C, "'"): If l028A Then l028C = RTrim$(Left$(l028C, l028A - 1))
  l028A = InStr(l028C, "Declare "): If l028A Then Return
  l028A = InStr(l028C, "Const "): If l028A Then Return
  l028A = InStr(l028C, " Type ")
  If l028A Then
    sub0B4F p0288, l0292, Mid$(l028C, l028A + 1)
    Return
  End If
  l028A = InStr(l028C, "Public ")
  If l028A Then
    l0296 = mc02A4
  Else
    l028A = InStr(l028C, "Private "): If l028A Then l0296 = mc02A6
  End If
  l028A = InStr(l028C, "Sub "): If l028A Then GoTo L17DBA
  l028A = InStr(l028C, "Function ")
  If l028A Then
    GoTo L17DBA
  End If
  l028A = InStr(l028C, "Property ")
  If l028A Then
    l0296 = l0296 Or mc02A8
    GoTo L17DBA
  End If
Return

L17D92:
  l0296 = mc02A8
L17DA2:

L17DAA:
  l028A = 1

L17DBA:
  l0290 = Mid$(l028C, l028A)
  l028A = InStr(l0290, " ")
  If l028A = 0 Then Return
  If InStr("Private Public Static Property Sub Function ", Left$(l0290, l028A)) = 0 Then
    Return
  End If
  l0290 = Mid$(l0290, l028A + 1)
  l028A = InStr(l0290, "(")
  If l028A = 0 Then ErrBeep: Return
  l0290 = RTrim$(Left$(l0290, l028A - 1))
  gv0AD6(gv0B08).M4045 = l0298
  l0294 = l0294 + 1
  l029C = l029C + 1
  If (l0296 And (mc02A6 Or mc02A4)) = 0 Then
    Select Case gv094A(p0288).M36B3_count
    Case mc00EC
      l0296 = l0296 Or mc02A6
    Case mc00F4
      l0296 = l0296 Or mc02A4
    Case mc00DC
      l0296 = l0296 Or mc02A6
    Case Else
      ErrBeep
    End Select
  End If
  sub09DB p0288, l0290, l0296
  l0296 = 0
Return
End Sub

Sub sub0B13(p02EE As Integer, p02F0 As Integer)
Dim l02F2 As Integer
Dim l02F4 As String
Dim l02F8 As Long
Dim l02FA
Dim l02FC As Integer
Dim l02FE As Integer
' Const mc0302 = 3 ' &H3%
Dim l0304 As Integer
  If gv097C = 0 Then Exit Sub
  l02FC = gv094A(p02EE).M3F1B + 1
  sub0B40 p02EE
  Select Case p02F0
  Case mc0286
    gv16B4 = True
  Case mc0284
  Case mc0302
    GoTo L18DCE
  End Select
  l02F2 = fn171A(p02EE, l02FC, mc0142)
  l02F8 = gv0AD6(l02FC).M4045
  Do While Seek(l02F2) < l02F8
    Line Input #l02F2, l02F4: l02F4 = LTrim$(l02F4)
    If Len(l02F4) = 0 Then
    ElseIf InStr("#'R", Left$(l02F4, 1)) Then
      While Right$(l02F4, 1) = "_"
        Line Input #l02F2, l02F4: l02F4 = Trim$(l02F4)
      Wend
    Else
      Select Case Left$(l02F4, 3)
      Case "Dim"
        If p02F0 = mc0286 Then
          l02F4 = LTrim$(Mid$(l02F4, 5)): GoSub L18EBA
        End If
      Case "Sta"
        l02FE = mc0286
        l02F4 = LTrim$(Mid$(l02F4, 8)): GoSub L18E3A
      Case "Pri"
        l02FE = mc0286
        l02F4 = LTrim$(Mid$(l02F4, 9)): GoSub L18E3A
      Case "Glo"
        l02FE = mc0284
        l02F4 = LTrim$(Mid$(l02F4, 8)): GoSub L18E3A
      Case "Pub"
        l02FE = mc0284
        l02F4 = LTrim$(Mid$(l02F4, 8)): GoSub L18E3A
      Case "Opt"
        sub0B5D l02F4
      Case "Def"
        sub0AC8 l02F4
      Case "Dec"
      Case "Con"
        If p02F0 = mc0286 Then GoSub L18F46
      Case "Typ"
        GoSub L18F52
      Case "Att"
        Select Case Mid$(l02F4, 11, 6)
        Case "VB_Nam"
        Case "VB_Cre"
        Case "VB_Exp"
        Case "VB_Des"
        Case "VB_Glo"
        Case "VB_Pre"
        Case "VB_Tem"
        End Select
      Case Else
        MsgBox l02F4, , "Declaration?"
      End Select
    End If
    While Right$(l02F4, 1) = "_"
      Line Input #l02F2, l02F4
    Wend
  Loop
  l02F2 = fn171A(p02EE, 0, m015A)
  If p02F0 = mc0286 Then
L18DCE:
    l0304 = gv094A(p02EE).M3F1B
    For l02FA = 2 To gv094A(p02EE).M3F13
      sub0BAB l0304 + l02FA, mc0302
    Next
  End If
Exit Sub

L18E3A:
  If Left$(l02F4, 5) = "Type " Then GoTo L18F52
  If p02F0 <> l02FE Then Return
  If Left$(l02F4, 6) = "Const " Then GoTo L18F46
  If Left$(l02F4, 8) = "Declare " Then Return
L18EBA:
  l02FA = InStr(l02F4, "'")
  If l02FA Then l02F4 = RTrim$(Left$(l02F4, l02FA - 1))
  l02F4 = l02F4 & ","
  While Len(l02F4)
    sub0A33 p02EE, l02F4, p02F0
  Wend
Return

L18F46:
Return

L18F52:
  While Left$(l02F4, 4) <> "End "
    While Right$(l02F4, 1) = "_"
      Line Input #l02F2, l02F4
    Wend
    Line Input #l02F2, l02F4: l02F4 = LTrim$(l02F4)
  Wend
Return
End Sub

Private Sub sub0B40(p0306 As Integer)
  gv16B2 = p0306
  gv3676 = 0: ReDim gv3640(10)
  gv3672 = "§"
  gv0988 = String$(255, VariantType)
  gv0984 = 0
  gv36E4 = 0
End Sub

Private Sub sub0B5D(p02AA As String)
  Select Case Mid$(p02AA, 8, 4)
  Case "Expl", "Comp"
  Case "Base"
    Select Case Mid$(p02AA, 12, 1)
    Case "0"
    Case "1": gv0984 = gv0984 Or 1
    Case Else
      MsgBox p02AA, , "Option Base?"
    End Select
  Case Else
    MsgBox p02AA, , "Option???"
  End Select
End Sub

Sub sub0BAB(p030C As Integer, p030E As Integer)
Dim l0310 As Integer
Dim l0312 As String
Dim l0314 As Integer
Dim l0316 As Long
Dim l0318 As String
Dim l031A As Integer
Dim l031C As String
Dim l0320 As Integer
Dim l0322 As Integer
Dim l0324 As String
  frmScan.control4 = gv0AD6(p030C).Name
  gv16B6 = p030C
  Assume gv16B2 = gv0AD6(p030C).M403B, "current module"
  
  'sub0BB9 gv0AD6(p030C).M404D
  Stop
  gv36E0 = "§"
  If gv0AD6(p030C).Flags < 0 Then Exit Sub
  l0316 = gv0AD6(p030C).M4045
  Assume l0316 <> 0, "sub size"
  l031A = gv094A(gv16B2).M36B3_count
  l0314 = fn171A(gv16B2, p030C, mc0142)
  Do While Seek(l0314) < l0316
    Line Input #l0314, l0312
    l0318 = LTrim$(l0312)
    If Len(l0318) Then
      If Left$(l0318, 1) <> "'" And Left$(l0318, 3) <> "Rem" Then Exit Do
    End If
  Loop
  While Right$(l0318, 1) = "_"
    Line Input #l0314, l0312
    l0318 = Left$(l0318, Len(l0318) - 1) & LTrim$(l0312)
  Wend
  l0310 = InStr(l0318, "'")
  If l0310 Then l0318 = RTrim$(Left$(l0318, l0310 - 1))
  Do
    l0310 = InStr(l0318, " ")
    Select Case Left$(l0318, l0310 - 1)
    Case "Sub": Exit Do
    Case "Function", "Property"
      l0318 = Mid$(l0318, l0310 + 1)
      l0310 = InStr(l0318, "(")
      l031C = Left$(l0318, l0310 - 1)
      If Right$(l0318, 1) <> ")" Then
        l0310 = FindLastOccurenceOf(l0318, ")")
        l031C = l031C & Mid$(l0318, l0310 + 1)
        l0318 = Left$(l0318, l0310)
      End If
      Exit Do
    Case "Private": l0318 = Mid$(l0318, l0310 + 1)
    Case "Public": l0318 = Mid$(l0318, l0310 + 1)
    Case "Static"
      l0318 = Mid$(l0318, l0310 + 1)
      l0320% = True
    Case Else
      MsgBox l0318, , "parse???"
      Exit Do
    End Select
  Loop
  l0310 = InStr(l0318, "(")
  l0324 = Left$(l0318, l0310 - 1)
  l0318 = LTrim$(Mid$(l0318, l0310 + 1))
  Mid$(l0318, Len(l0318), 1) = ","
  Do While Len(l0318) > 1
    sub0A6F p030C, l0318, mc01B2, l031A
  Loop
  sub0A6F p030C, l031C, mc018E, l031A
  Do
    Line Input #l0314, l0312
    l0318 = Trim$(l0312)
    While Right$(l0318, 1) = "_"
      Line Input #l0314, l0312
      l0318 = Left$(l0318, Len(l0318) - 1) & Trim$(l0312)
    Wend
    l0310 = InStr(l0318, "'")
    If l0310 Then l0318 = Trim$(Left$(l0318, l0310 - 1))
    l0310 = InStr(l0318, " ")
    If l0310 = 0 Then l0310 = Len(l0318) + 1
    Select Case Left$(l0318, l0310 - 1)
    Case "Static"
      If p030E = mc0302 Then
        l0318 = Mid$(l0318, l0310 + 1) & ","
        Do While Len(l0318)
          sub0A33 p030C, l0318, mc0302
        Loop
      End If
    Case "Dim"
      If l0320 Then
        If p030E = mc0302 Then
          l0318 = Mid$(l0318, l0310 + 1) & ","
          Do While Len(l0318)
            sub0A33 p030C, l0318, mc0302
          Loop
        End If
      ElseIf l0322 Then
        MsgBox l0324 & Chr$(10) & l0318, , "move to begin of sub:"
        l0318 = Mid$(l0318, l0310 + 1) & ","
        Do While Len(l0318)
          sub0A6F p030C, l0318, mc01BA, l031A
        Loop
      Else
        l0318 = Mid$(l0318, l0310 + 1) & ","
        Do While Len(l0318)
          sub0A6F p030C, l0318, mc01B8, l031A
        Loop
      End If
    Case "Const"
    Case "Rem"
    Case "End": Exit Do
    Case Else
      If Len(l0318) Then
        If Left$(l0318, 1) <> "'" Then
          l0322 = True
        End If
      End If
    End Select
  Loop
  sub0BC7 p030C, l031A
  gv16B8 = True
  l0314 = fn171A(gv16B2, p030C, m015A)
  sub0BD8
End Sub

