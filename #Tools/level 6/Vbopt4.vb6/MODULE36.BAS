Attribute VB_Name = "MODULE36"
' Module36
Option Explicit
Const mc00B4 = 3 ' &H3%
Const mc00EE = 128 ' &H80%

Sub sub21F3(p00A2 As T6407, p00A4 As String)
Dim l00A6 As Integer
Dim l00A8
Dim l00AA As Integer
Dim l00AC As String
Dim l00AE As Integer
Dim l00B0 As String
Dim l00B2 As String
' Const mc00B4 = 3 ' &H3%
  Do While Len(p00A4)
    l00AA = InStr(p00A4, ",")
    p00A2.M6414 = p00A2.M6414 + 1
    l00AC = Trim$(Left$(p00A4, l00AA - 1))
    If UCase(l00AC) = l00AC Then
      If Len(l00AC) = 2 Then
        For l00A8 = 3 To 1 Step -1
          l00A6 = InStr(gv2422(l00A8), l00AC)
          If l00A6 Then Exit For
        Next
        Select Case l00A8
        Case 0: ErrBeep
        Case 3
          l00AE = 2
        Case 1 To 4
          l00AE = l00A8
        Case 5
          l00AE = 0
        Case Else: ErrBeep
        End Select
        l00A6 = l00A8 * 8 + (l00A6 \ 2)
        l00A8 = gc2478
      ElseIf l00AC = "ST0" Then
        l00A6 = gc23D2
        l00A8 = gc2478
        l00AE = 0
      Else
        l00A6 = InStr(gv2422(gc23DA \ 8), l00AC)
  If l00A6 = 0 Then ErrBeep
        l00A6 = gc23DA + l00A6 \ 3
        Select Case l00A6
        Case gc245C: l00AE = 2:
        Case gc246A: l00AE = 2
        Case gc246C, gc2470: l00AE = 14
        Case gc246E, gc2472: l00AE = 94
        Case Else: l00AE = 0
        End Select
        If l00A6 <= gc2464 Then
          l00A8 = gc2478 Or gc2476
        Else
          l00A8 = gc2478
        End If
      End If
    Else
      l00B0 = Left$(l00AC, 1)
      If p00A2.M641D = 0 Then
        If InStr("ESRCDTFGM", l00B0) Then
          p00A2.M641D = mc00B4
        End If
      End If
      l00B2 = Right$(l00AC, 1)
      l00A8 = InStr("EeSRFCDTGMAIJO", l00B0)
      If l00A8 = 0 Then ErrBeep
      l00AE = InStr("bwvlaspqntfd", l00B2)
      l00A6 = l00AE
      Select Case l00B0
      Case "e"
        l00A8 = gc2478 Or gc2476
        l00A6 = InStr(gv2422(2), Mid$(l00AC, 2))
        If l00A6 = 0 Then ErrBeep
        l00A6 = gc23CC + (l00A6 \ 2)
        l00AE = 2
      Case "S"
        l00A6 = l00A8 * 8
        l00A8 = gc248E
        l00AE = 2
      Case "F"
        l00A6 = l00A8 * 8
        l00A8 = gc248E
        l00AE = 0
      Case "C", "D", "T"
        l00A6 = l00A8 * 8
        l00A8 = gc248E
        l00AE = 4
      Case "G"
        l00A8 = gc248E
        Select Case l00AE
        Case 3
          l00A8 = l00A8 Or gc2476
          l00A6 = gc23CC
          l00AE = 2
        Case 1 To 4
          l00A6 = l00AE * 8
        Case Else
          ErrBeep
        End Select
      Case "R"
        l00A8 = gc2492
        If l00AE = 3 Then
          l00A8 = l00A8 Or gc2476
          l00A6 = gc23CC
          l00AE = 2
        Else
          l00A6 = l00AE * 8
        End If
      Case "E"
        l00A8 = gc2494
        If l00AE <> 3 Then
          l00A6 = l00AE * 8
        End If
      Case "M"
        l00A8 = gc2490
      Case "A"
        l00A8 = gc2486
        l00AE = 4
      Case "I"
        l00A8 = gc247A
      Case "J"
        l00A8 = gc248C
      Case "O"
        l00A8 = gc248A
      Case Else
        ErrBeep
      End Select
    End If
    If l00AE = 3 Then
      l00AE = 2
      l00A8 = l00A8 Or gc2476
    End If
    If l00A8 = 0 Or l00A6 = 0 Then
      MsgBox p00A4, , "unknown argument"
      Exit Sub
    End If
    p00A2.M6441(p00A2.M6414).M63E9 = l00A8
    p00A2.M6441(p00A2.M6414).TypeCode_M5F96 = l00A6
    p00A2.M6441(p00A2.M6414).Size = l00AE
    p00A4 = Trim$(Mid$(p00A4, l00AA + 1))
  Loop
End Sub

Sub sub2201(p00B6 As Integer, ByVal p00B8 As String)
Dim l00BA
  For l00BA = 0 To 7
    gv23E0(p00B6 + l00BA) = p00B8 & Format$(l00BA)
  Next
End Sub

Sub sub220D(p00BC As Integer, ByVal p00BE As String)
Dim l00C0
  For l00C0 = 0 To 7
    gv23E0(p00BC + l00C0) = Mid$(p00BE, l00C0 * 2 + 1, 2)
  Next
  gv2422(p00BC \ 8) = p00BE
End Sub

Sub sub2218(p00C4 As T63DC, p00C8 As T63F1, p00CA As Integer)
Dim l00CC As Integer
Dim l00CE As Integer
Dim l00D0 As Integer
Dim l00E8 As Long
' Const mc00EE = 128 ' &H80%
  l00CE = p00C4.M63E9
  l00CC = p00C4.TypeCode_M5F96
  l00D0 = p00C4.Size
  If l00CE And gc2476 Then
    If gAsmNew.M65EC Then
      l00CC = l00CC + &H10
      l00D0 = l00D0 + 2
    End If
    l00CE = l00CE And &H1F
  End If
  p00C8.M63E9 = l00CE
  Select Case l00CE
  Case gc2478
    If p00CA < 0 Then
      gDisOutputTxt = gDisOutputTxt & " " & gv23E0(l00CC)
    End If
    GoTo L49FBA
  Case gc247A
    GoSub L4A054
    If p00CA < 0 Then
      If p00C8.subRef_M3C6A Then
        gDisOutputTxt = gDisOutputTxt & " &" & fn2023(p00C8.subRef_M3C6A)
      Else
        gDisOutputTxt = gDisOutputTxt & " " & Hex$(p00C8.M4106)
      End If
    End If
    GoTo L49FBA
  Case gc248C
    GoSub L49FF2
    p00C8.M4106 = p00C8.M4106 + gVACmdEnd
If g_IsNE Then
    p00C8.M3C63 = fn087D(gVACmdEnd)
Else
End If
    p00C8.M63E9 = gc247E
    sub2224 p00C8
  Case gc248A
    If gAsmNew.M65F7 Then
      p00C8.M4106 = fn22E6(4, p00C8.subRef_M3C6A)
    Else
      p00C8.M4106 = fn22E6(2, p00C8.subRef_M3C6A)
    End If
    sub2224 p00C8
    If p00CA < 0 Then
      gDisOutputTxt = gDisOutputTxt & " ["
      If gAsmNew.M6603 Then
        gDisOutputTxt = gDisOutputTxt & gv23E0(gAsmNew.M6603) & ":"
      End If
      If p00C8.subRef_M3C6A Then
        gDisOutputTxt = gDisOutputTxt & fn2023(p00C8.subRef_M3C6A)
      Else
If g_IsDemoVersion Then
          gDisOutputTxt = gDisOutputTxt & "pppp"
Else
        If p00C8.M3C63 Then
          gDisOutputTxt = gDisOutputTxt & Format$(p00C8.M3C63) & ":"
        End If
        gDisOutputTxt = gDisOutputTxt & Hex$(p00C8.M4106)
End If
      End If
      gDisOutputTxt = gDisOutputTxt & "]"
    End If
    GoTo L49FBA
  Case gc2486
    If gAsmNew.M65F7 Then
      If gIsVB_Ver Then
        sub2232
        Exit Sub
      End If
      l00D0 = 6
    Else
      l00D0 = 5
    End If
    p00C8.M4106 = fn22E6(l00D0, p00C8.subRef_M3C6A)
    sub2224 p00C8
  Case gc248E
    l00CC = l00CC + gAsmNew.M661F

L49A2C:
    l00CE = gc2478
    If p00CA < 0 Then
      gDisOutputTxt = gDisOutputTxt & " " & gv23E0(l00CC)
    End If
    GoTo L49FBA
  Case gc2490
    If gAsmNew.M6616 = 3 Then
      sub2232
      Exit Sub
    End If
    GoTo L49B1C
  Case gc2492
    If gAsmNew.M6616 <> 3 Then
      sub2232
      Exit Sub
    End If
    l00CC = l00CC + gAsmNew.M6627
    GoTo L49A2C
  Case gc2494
L49B1C:
    If p00CA < 0 Then
      If gAsmNew.M6616 = 3 Then
        l00CC = gAsmNew.M6627 + l00D0 * 8
        GoTo L49A2C
      Else
        l00E8 = gAsmNew.M6656
        gDisOutputTxt = gDisOutputTxt & " "
        If gAsmNew.subRef_M3C6A Then
          If l00E8 Then
            ErrBeep
          End If
          gDisOutputTxt = gDisOutputTxt & fn2023(gAsmNew.subRef_M3C6A)
          If gAsmNew.M664D Then
            gDisOutputTxt = gDisOutputTxt & "+" & gv23E0(gAsmNew.M664D)
          End If
        ElseIf gAsmNew.M6603 Then
          If gAsmNew.M664D Then
            gDisOutputTxt = gDisOutputTxt & gv23E0(gAsmNew.M6603) & ":" & gv23E0(gAsmNew.M664D)
          Else
            If l00E8 Then
              p00C8.M63E9 = gc247E
              sub2224 p00C8
            End If
            gDisOutputTxt = gDisOutputTxt & gv23E0(gAsmNew.M6603) & ":" & Hex$(l00E8)
            l00E8 = 0
          End If
        ElseIf gAsmNew.M664D Then
          gDisOutputTxt = gDisOutputTxt & gv23E0(gAsmNew.M664D)
        End If
        gDisOutputTxt = gDisOutputTxt & "["
        If gAsmNew.M6643 Then
          gDisOutputTxt = gDisOutputTxt & gv23E0(gAsmNew.M6643)
          If gAsmNew.M663A > 1 Then
            gDisOutputTxt = gDisOutputTxt & "*" & Format$(gAsmNew.M663A)
          End If
        End If
        If l00E8 Then
          If l00E8 > 0 Then
            gDisOutputTxt = gDisOutputTxt & "+" & Hex$(l00E8) & "h"
          Else
            gDisOutputTxt = gDisOutputTxt & Str$(l00E8)
          End If
        End If
        gDisOutputTxt = gDisOutputTxt & "]"
        p00C8.subRef_M3C6A = gAsmNew.subRef_M3C6A
        p00C8.M4106 = l00E8
      End If
    Else
    End If
    If gAsmNew.M664D = gc2442 Then
      l00CE = gc2484
    ElseIf gAsmNew.M664D Or gAsmNew.M6643 Then
      l00CE = gc2482
    Else
      l00CE = gc247E
    End If
    GoTo L49FBA
  Case Else
    ErrBeep
  End Select
  If p00CA < 0 Then
    If gCmdByteStruc.M642E And mc00EE Then
      gDisOutputTxt = gDisOutputTxt & "<error>"
    ElseIf p00C8.subRef_M3C6A Then
      gDisOutputTxt = gDisOutputTxt & " " & fn2023(p00C8.subRef_M3C6A)
    Else
        gDisOutputTxt = gDisOutputTxt & " [" & Hex$(p00C8.M4106) & "]"
    End If
  Else
  End If

L49FBA:
  p00C8.M63E9 = l00CE
  p00C8.TypeCode_M5F96 = l00CC
  p00C8.Size = l00D0
Exit Sub

L49FF2:
  If l00D0 = 3 Then
    ErrBeep
    If gAsmNew.M65EC Then l00D0 = 4 Else l00D0 = 2
  End If
  p00C8.M4106 = fn22D8(l00D0, p00C8.subRef_M3C6A)
Return

L4A054:
  If l00D0 = 3 Then
    ErrBeep
    If gAsmNew.M65EC Then l00D0 = 4 Else l00D0 = 2
  End If
  p00C8.M4106 = fn22E6(l00D0, p00C8.subRef_M3C6A)
Return
End Sub

Sub sub2240()
Dim l00F2
Dim l00F4 As String
  sub220D gc23CA, "ALCLDLBLAHCHDHBH"
  sub220D gc23CC, "AXCXDXBXSPBPSIDI"
  sub220D gc23CE, "ESCSSSDSFSGS"
  For l00F2 = 0 To 7
    gv23E0(gc23D0 + l00F2) = "E" & gv23E0(gc23CC + l00F2)
  Next
  sub2201 gc23D2, "ST"
  sub2201 gc23D4, "CR"
  sub2201 gc23D6, "DR"
  sub2201 gc23D8, "TR"
  l00F4 = "FLGSP+SP-SRCDSTST+ST-CWDENVFSTM14M94"
  gv2422(gc23DA \ 8) = l00F4
  gv23E0(gc245C) = "FLAGS"
  gv23E0(gc245E) = "[SP+]"
  gv23E0(gc2460) = "[-SP]"
  gv23E0(gc2462) = "DS:SI"
  gv23E0(gc2464) = "ES:DI"
  gv23E0(gc245C + &H10) = "EFLAGS"
  gv23E0(gc245E + &H10) = "[ESP+]"
  gv23E0(gc2460 + &H10) = "[-ESP]"
  gv23E0(gc2462 + &H10) = "DS:ESI"
  gv23E0(gc2464 + &H10) = "ES:EDI"
  gv23E0(gc2466) = "[ST+]"
  gv23E0(gc2468) = "[-ST]"
  gv23E0(gc246A) = "CW87"
  gv23E0(gc246C) = "ENV87"
  gv23E0(gc246E) = "STATE87"
  gv23E0(gc2470) = "M14"
  gv23E0(gc2472) = "M94"
End Sub

Sub sub2252(p00F8 As T63DC)
Dim l00FA As Integer
  Select Case p00F8.M63E9 And &H1F
  Case gc248A
    If gAsmNew.M65F7 Then l00FA = 4 Else l00FA = 2
    If gVACmdEnd = gVAModuleEnd Then sub019B
    gVACmdEnd = gVACmdEnd + l00FA
  Case gc247A, gc248C
    l00FA = p00F8.Size
    If p00F8.M63E9 And gc2476 Then
      If gAsmNew.M65EC Then l00FA = l00FA + 2
    ElseIf l00FA = 3 Then
      ErrBeep
      If gAsmNew.M65EC Then l00FA = 4 Else l00FA = 2
    End If
    If gVACmdEnd = gVAModuleEnd Then sub019B
    gVACmdEnd = gVACmdEnd + l00FA
    p00F8.Size = l00FA
  Case gc2490
    If gAsmNew.M6616 = 3 Then sub2232
  Case gc2492
    If gAsmNew.M6616 <> 3 Then sub2232
  Case gc2486
    sub2232
  End Select
End Sub
