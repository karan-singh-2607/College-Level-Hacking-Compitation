Attribute VB_Name = "MODULE2"
' Module2
Option Explicit
Dim m004A As Integer
Dim m004C As Long
Const mc0058 = 2 ' &H2%
Const mc005A = 8 ' &H8%
Const mc005C = 31 ' &H1F%
Const mc005E = 5 ' &H5%
Const mc00C8 = 33 ' &H21%
Const mc00D4 = 6 ' &H6%
Const mc00DA = 1 ' &H1%
Const mc00DC = 2 ' &H2%
Const mc00E2 = 3 ' &H3%
Const mc00E4 = 4 ' &H4%
Const mc011C = 38 ' &H26%
Const mc012C = 1 ' &H1%
Const mc014C = 3 ' &H3%
Const mc0168 = 32 ' &H20%
Const mc01B8 = 18 ' &H12%
Const mc01DA = 4096 ' &H1000%
Const mc01E0 = 63 ' &H3F%
Const mc01E2 = 32 ' &H20%
Const mc01E4 = 1 ' &H1%
Const mc01E6 = 2 ' &H2%
Const mc01E8 = 18 ' &H12%
Const mc01EA = 3 ' &H3%
Const mc01EC = 4 ' &H4%
Const mc01EE = 5 ' &H5%
Const mc01F0 = 6 ' &H6%
Const mc01F2 = 7 ' &H7%
Const mc01F4 = 8 ' &H8%
Const mc01F6 = 9 ' &H9%
Const mc01F8 = 10 ' &HA%
Const mc01FA = 11 ' &HB%
Const mc01FC = 1 ' &H1%
Const mc01FE = 2 ' &H2%
Const mc0200 = 3 ' &H3%
Const mc0204 = 10 ' &HA%
Const mc0206 = 4 ' &H4%
Const mc0208 = 5 ' &H5%
Const mc021A = 261 ' &H105%
Const mc021C = 9 ' &H9%
Const mc0240 = 28 ' &H1C%
Const mc0242 = 29 ' &H1D%
Const mc024C = 256 ' &H100%

Function ReadZeroString(hFile As Integer, Offset As Long) As String
Dim ReadBuff As String * 64
Dim ZeroPos As Integer
Dim CurFilePos As Long
  
 'Retain current FilePos
  CurFilePos = Seek(hFile)
  
   ' Read and seekOver
    Get hFile, Offset, ReadBuff
  
  Seek hFile, CurFilePos
  
 'Cut at ZeroChar
  ZeroPos = InStr(ReadBuff, Chr$(0))
  ReadZeroString = Left$(ReadBuff, ZeroPos - 1)
  
End Function

Function fn0218(p019A As Long) As String
  fn0218 = ReadZeroString(g_hFileExe, VAtoRaw(p019A))
End Function

Function fn0226(p01A0 As Integer) As String
  fn0226 = "#" & Format$(p01A0)
End Function

Sub sub00AC()
LB300:
  If gv014C Then
    gVAModuleEnd = gv161C(gv014C).VOffset
    gv014E = gv161C(gv014C).M36B3_count
      If gv014E = mc0058 Then
        Select Case gv013E
        Case mc005A To mc005C
        Case Else
          sub00B9 gv161C(gv014C)
          gv014E = gv161C(gv014C).M36B3_count
          If gv014E = mc005E Then
            If gReloc_Arr(gv013C).Count <> gv0140 Then
              gv0140 = gReloc_Arr(gv013C).Count
              gv0142 = gReloc_Arr(gv0140).M36B3_count
              If gIsVB5 >= 0 Then
                gv0144 = fn202F(gv0140)
              End If
            End If
            GoTo LB300
          End If
        End Select
      End If
    gv0150 = gv161C(gv014C).subRef_M3C6A
  Else
    gVAModuleEnd = gv0144
    gv014E = 0
  End If
End Sub

Sub sub00C9()
Dim l006C As Integer
Dim l006E As Long
Dim l0070 As String
Dim l0072 As String
If gIsVB5 Then
  DbgMileStone "decomp sub"
  Exit Sub
End If
  l006C = fn1D2E()
  l006E = gBuggy(l006C).VOffset
  gVACmdEnd = gModuleVAStart
  Assume gVACmdEnd <> 0, "Sub loaded"
  gModuleVACur = gModuleVAEnd
  gv0144 = gModuleVACur
  gv014C = fn1624(gVACmdEnd)
  sub00AC
  sub00E9
  Do
    l0072 = fn18AB(gVACmdEnd)
    If l006C > 1 Then
    Do While gBuggy(l006C).VOffset < gVACmdEnd
      l0070 = fn1D20(l006C)
      If Len(l0070) Then DisOutput l0070 & ":": WriteOutputTxt
      l006C = fn1D5C(l006C)
      If l006C <= 1 Then
        Exit Do
      End If
    Loop
    End If
    If Len(l0072) Then
      If gv34F4 Then
        Exit Do
      Else
        DisOutput "  " & Trim$(l0072): WriteOutputTxt
        If Left$(l0072, 1) <> "'" Then sub00E9
      End If
    End If
  Loop Until gv34F4
End Sub

Sub sub0107(p0090 As Integer, ByVal pv0092 As Integer)
Dim l0094 As Integer
Dim l0096 As Long
Dim l0098 As Long
Dim l009A As String
  WriteOutputTxt
  l0096 = fn202F(p0090)
  sub0115 l0096
  l0094 = gReloc_Arr(p0090).Count
  If l0094 Then
    l0098 = fn202F(l0094)
  Else
    l0098 = gv1066
  End If
  If l0096 = l0098 Then
    DisOutput "<end segment>": WriteOutputTxt
  Else
    l009A = fn2023(p0090)
    DisOutput l009A: WriteOutputTxt
    If pv0092 = 0 Then pv0092 = gReloc_Arr(p0090).M36B3_count
    sub0126 p0090, l0096, l0098, pv0092
  End If
End Sub

Sub sub0126(p00A4 As Integer, p00A6 As Long, p00A8 As Long, ByVal pv00AA As Integer)
Dim l00AC As String
Dim l00AE As Integer
Dim l00B2 As Integer
Dim l00B4%
Dim l00B6 As Integer
Dim l00B8 As Integer 'or Long?

Dim l00BA As Long
Dim l00BC%
Dim l00BE As Long
Dim l00C0 As Integer
Dim l00C2 As Long
Dim l00C4%
Dim l00C6 As Integer
' Const mc00C8 = 33 ' &H21%
' Const mc00D4 = 6 ' &H6%
Static m00D8 As Integer
' Const mc00DA = 1 ' &H1%
' Const mc00DC = 2 ' &H2%
' Const mc00E2 = 3 ' &H3%
' Const mc00E4 = 4 ' &H4%
Dim l00F8 As Integer
Dim l00FA As Integer
' Const mc011C = 38 ' &H26%
Dim l0122 As Integer 'or Long?

' Const mc012C = 1 ' &H1%
Dim l0138 As Integer 'or Long?
' Const mc014C = 3 ' &H3%
If gIsVB5 < 0 Then
Else
  If pv00AA = mc00C8 Then
    If p00A6 = gModuleVAStart Then
      If gv16BC > 0 Then
        Assume gEXEdef_Subs(gv16BC).lOffsetSym_M3C6A = p00A4, "sub loaded"
      End If
      l00C0 = fn1D2E()
      If l00C0 Then
        l00C2 = gBuggy(l00C0).VOffset
      Else
        l00C2 = &H7FFFFFFF
      End If
    Else
      l00C2 = &H7FFFFFFF
      If pv00AA = mc00C8 Then
        pv00AA = mc00D4
      End If
    End If
  End If
End If
  sub014F p00A4, 0
  gv013E = pv00AA
  gVACmdEnd = p00A6
  gv0144 = p00A8
  Do While gVACmdEnd < gv0144
    gVACmdStart = gVACmdEnd
    If gVACmdEnd >= gVAModuleEnd Then
      sub015B True
    Else
      Select Case gv013E
      Case m00D8
      Case mc00DA
      Case mc00DC
        While gVACmdEnd < gVAModuleEnd
          If gIsVB_Ver Then
            sub0167 g_hFileExe, gVACmdEnd, "Z"
          Else
            sub0167 g_hFileExe, gVACmdEnd, "z"
          End If
        Wend
      Case mc00E2
        While gVACmdEnd < gVAModuleEnd
          sub0167 g_hFileExe, gVACmdEnd, "z"
        Wend
      Case mc00E4
        While gVACmdEnd < gVAModuleEnd
          sub0167 g_hFileExe, gVACmdEnd, "Z"
        Wend
        If gIsVB5 Then
          sub0167 g_hFileExe, gVACmdEnd - 4, " (§l )"
        End If
      Case mc00C8
        Do While gVACmdEnd < gVAModuleEnd And gVACmdEnd < gv0144
          While gVACmdEnd >= l00C2
            If gVACmdEnd > l00C2 Then
              ErrBeep
            Else
              l00B2 = gBuggy(l00C0).M4C81
              If l00B2 Then
                DisOutput Format$(ToUnSigned(l00B2))
              ElseIf gBuggy(l00C0).M36B3_count Then
                DisOutput fn1D20(l00C0) & ":"
              Else
                DisOutput "<nl>"
              End If
              WriteOutputTxt
            End If
            l00C0 = fn1D5C(l00C0)
            If l00C0 > 1 Then
              l00C2 = gBuggy(l00C0).VOffset
            Else
              l00C2 = &H7FFFFFFF
            End If
          Wend
If gIsVB5 Then
          l00B2 = fn1EA9(gVAToRaw + gVACmdEnd)
Else
          Get g_hFileExe, gVAToRaw + gVACmdEnd, gv34F6
          l00B2 = 2
End If
          If gv34F6 < 0 Or gv34F6 > UBound(gDatArrLinks) Then
            l00AE = 0
          Else
            gVACmdStart = gVACmdEnd
            l00AE = gDatArrLinks(gv34F6)
          End If
          If l00AE Then
            l00AC = fn04E4(l00AE)
            If Len(l00AC) = 0 Then l00AC = fn0842(gv34F6) & "?"
            l00AC = " " & l00AC
            gVACmdEnd = gVACmdEnd + l00B2
            l00F8 = gDatArr1(l00AE).M3818
            l00FA = 0
            Do
              Select Case Mid$(gDatStrings1, l00F8, 1)
              Case "}"
                Do While l00FA > 0
                  Get g_hFileExe, , l00B2: gVACmdEnd = gVACmdEnd + 2
                  l00AC = l00AC & fn085C(l00B2)
                  l00FA = l00FA - 1
                Loop
                If gVACmdEnd + 4 >= gv0144 Then
                  gVACmdEnd = (gVACmdEnd + 3) And Not 3
                End If
                Exit Do
              Case "§"
                Do While l00FA > 0
                  Get g_hFileExe, , l00B2: gVACmdEnd = gVACmdEnd + 2
                  l00AC = l00AC & fn085C(l00B2)
                  l00FA = l00FA - 1
                Loop
                Exit Do
              Case "ß"
                gVACmdEnd = gVACmdEnd - 2
                l00FA = l00FA + 1
              Case "~"
                Get g_hFileExe, , l00FA
                l00FA = (l00FA \ 2) + 1
                l00AC = l00AC & "[" & Hex$(l00FA - 1) & "]"
              Case "D", "T"
                gVACmdEnd = gVACmdEnd + l00FA * 2 - 2
                While l00FA > 0
                  GoSub LD596
                  l00FA = l00FA - 1
                Wend
              Case "L"
                gVACmdEnd = gVACmdEnd + l00FA * 2 - 2
                While l00FA > 0
                  Get g_hFileExe, , l00B2
                  l00AC = l00AC & " " & fn100C(l00B2, gv013C)
                  l00FA = l00FA - 1
                Wend
              Case "a"
                Get g_hFileExe, , l00B2
                If l00B2 <= gv356C Then
                  l00AC = l00AC & " " & gv353A(l00B2).Name
                Else
                  l00AC = l00AC & " <------- NO arg!"
                End If
                l00F8 = l00F8 + 1
              Case "P"
                GoSub LD652
                If l00AE = 0 Then Exit Do
              Case "p"
                GoSub LD652
                If l00AE = 0 Then Exit Do
                l00F8 = l00F8 + 1
              Case "."
                Get g_hFileExe, , l00B8
                If Mid$(gDatStrings1, l00F8, 3) = ".pc" Then
                  gVACmdEnd = gVACmdEnd + 2
                  l00FA = l00FA - 1
                  If gVACmdEnd <> gVAModuleEnd Then
                    GoSub LD652
                  ElseIf gv0150 Then
                    Get g_hFileExe, fn1483(gv0150), gv0B84
                    l00B6 = fn124C(gv0B84, False)
                    If l00B6 = 0 Then
                      l00B6 = Asc(Left$(gv0B84.M3CE1, 1))
                      If l00B6 Then
                        Mid$(gv0B84.M3CE1, 1, 1) = Chr$(l00B6 - 1)
                        l00B6 = fn124C(gv0B84, True)
                        Assume l00B6 > 0, "class defined"
                      End If
                    End If
                    If gReloc_Arr(gv0150).M36B3_count <> mc011C Then
                      gReloc_Arr(gv0150).M36B3_count = mc011C
                    End If
                    GoSub LD652
                    l00AC = l00AC & "." & fn12C6(l00B6, l00B8)
                  Else
                    GoSub LD652
                  End If
                  If l00AE = 0 Then Exit Do
                  l00F8 = l00F8 + 2
                Else
                  Stop
                  l00AC = l00AC & " ." & toHex16(CLng(l00B8)) 'CLng added
                End If
              Case "d", "b", "j", "t", "f", "N"
                GoSub LD596
              Case "e"
                Get g_hFileExe, , l00BC
                If l00BC = -1 Then
                  l00AC = l00AC & " Resume Next"
                ElseIf l00BC = -2 Then
                  l00AC = l00AC & " GoTo 0"
                Else
                  l00AC = l00AC & " GoTo"
                  GoSub LD5AC
                End If
              Case "c"
                Get g_hFileExe, , l0122
                l0122 = (l0122 - &H2AC) \ 4
                DebugPrintVar l0122 >= 0, "vtbl base"
                l00B2 = gEXEdef_Subs(gv16BC).ModuleIndex
                If g_decl_objs(l00B2).M36B3_count <> mc012C Then
                  l00B2 = g_decl_objs(l00B2).M41E0 + l0122
                  If l00B2 <= UBound(gv0D3E) Then
                    l00AC = l00AC & " " & gv0D3E(l00B2).M3F23
                  Else
                    Stop
                    l00AC = l00AC & "------------------ ctl#?" & toHex16(CLng(l0122))
                  End If
                Else
                  l00AC = l00AC & "----------------------" & Hex$(l0122)
                End If
              Case "m"
                DbgMileStone "split prop/meth"
                Get g_hFileExe, , l0122
                l00B2 = (l0122 \ 8) - 5
                l00AC = l00AC & " " & fn20F2(&HD, l00B2)
              Case "l", "i"
                Get g_hFileExe, , l00B2
                l00AC = l00AC & " " & fn100C(l00B2, gv013C)
              Case "M"
                Get g_hFileExe, , l0138
                Get g_hFileExe, , l00B2: l00FA = l00FA - 1: gVACmdEnd = gVACmdEnd + 2
                l00AC = l00AC & " " & fn103E(l0138, l00B2, gv013C)
              Case "s"
                Get g_hFileExe, , l00B2
                l00AC = l00AC & " " & fn1182(gv2FDE, l00B2)
              Case "G"
                Get g_hFileExe, , l00B4
                gVACmdEnd = gVACmdEnd + 2
                l00FA = l00FA - 1
                l00AC = l00AC & " " & fn0FD9(l00B4)
              Case "n"
                Get g_hFileExe, , l00B2
                Stop
                l00AC = l00AC & " [" & toHex16(CLng(l00B2)) & "]"
              Case "+"
                Get g_hFileExe, , l00B2
                Stop
                l00AC = l00AC & " +" & toHex16(CLng(l00B2))
              Case "B", "°"
                Get g_hFileExe, , gSingleChar: l00B2 = Asc(gSingleChar)
                If l00B2 And &H80 Then l00B2 = l00B2 Or &HFF00
                l00AC = l00AC & fn085C(l00B2)
                gVACmdEnd = gVACmdEnd - 1
              Case "%", "2"
                Get g_hFileExe, , l00B2
                l00AC = l00AC & Str$(l00B2) & "%"
              Case "&", "4"
                Get g_hFileExe, , gv0120
                gVACmdEnd = gVACmdEnd + 2: l00FA = l00FA - 1
                l00AC = l00AC & Str$(gv0120) & "&"
              Case "!"
                Get g_hFileExe, , gv0124
                gVACmdEnd = gVACmdEnd + 2
                l00AC = l00AC & Str$(gv0124) & "!"
              Case "#", "8"
                Get g_hFileExe, , gv0128
                gVACmdEnd = gVACmdEnd + 6
                l00AC = l00AC & Str$(gv0128) & "#"
              Case "@"
                Get g_hFileExe, , gv0130
                gVACmdEnd = gVACmdEnd + 6
                l00AC = l00AC & Str$(gv0130) & "@"
              Case "u"
                Get g_hFileExe, , gv0120
                If gv0120 Then
                  gv0120 = gv0120 \ 2
                  gv0138 = Space$(gv0120)
                  For l00C4 = 1 To gv0120
                    Get g_hFileExe, , l00B2
                    DbgVar l00B2, "strlen>0"
                    If l00B2 = 0 Then
                      gv0138 = Left$(gv0138, l00C4 - 1)
                      Exit For
                    Else
                      Mid$(gv0138, l00C4, 1) = Chr$(l00B2)
                    End If
                  Next
                Else
                  gv0138 = ""
                End If
                l00FA = l00FA - gv0120 - 2
                gVACmdEnd = gVACmdEnd + gv0120 * 2 + 4
                If Len(gv0138) = 1 Then
                  If Asc(gv0138) = 0 Then
                    l00AC = l00AC & " vbNullChar"
                  Else
                    l00AC = l00AC & " """ & gv0138 & """"
                  End If
                Else
                  l00AC = l00AC & " """ & gv0138 & """"
                End If
              Case "$"
                Get g_hFileExe, , gv0120
If gIsVB_Ver Then
                If gv0120 Then
                  DebugPrintVar gv0120 < l00FA, "strlen>args"
                  gv0138 = Space$(gv0120 - 1)
                  For l00C4 = 1 To gv0120 - 1
                    Get g_hFileExe, , l00B2
                    If l00B2 = 0 Then
                      gv0138 = Left$(gv0138, l00C4 - 1)
                      Exit For
                    Else
                      Mid$(gv0138, l00C4, 1) = Chr$(l00B2)
                    End If
                  Next
                Else
                  gv0138 = ""
                  gv0120 = 1
                End If
                l00FA = l00FA - gv0120 - 1
                gVACmdEnd = gVACmdEnd + gv0120 * 2
Else
                gv0138 = Space$(gv0120)
                Get g_hFileExe, , gv0138
                gv0120 = gv0120 And -2
                gVACmdEnd = gVACmdEnd + gv0120 + 4
                l00FA = l00FA - 2 - (gv0120 \ 2)
End If
                If Len(gv0138) = 1 Then
                  If Asc(gv0138) = 0 Then
                    l00AC = l00AC & " vbNullChar"
                  Else
                    l00AC = l00AC & " """ & gv0138 & """"
                  End If
                Else
                  l00AC = l00AC & " """ & gv0138 & """"
                End If
              Case ":"
                DisOutput "nl": WriteOutputTxt
                DisOutput l00AC: WriteOutputTxt: l00AC = ""
                If l00FA > 0 Then gVACmdEnd = gVACmdEnd + l00FA * 2
                Exit Do
              Case "_", "=", "1"
                DisOutput l00AC: WriteOutputTxt
                If l00FA > 0 Then gVACmdEnd = gVACmdEnd + l00FA * 2
                l00AC = "nl"
                Exit Do
              Case "g"
                gVACmdEnd = gVACmdEnd + l00FA * 2
                l00FA = 1
                gVACmdEnd = gVACmdEnd - 2
              Case "w"
                Get g_hFileExe, , l00B2
                l00AC = l00AC & fn085C(l00B2)
              Case "C"
                l00F8 = l00F8 + 1
                Select Case Mid$(gDatStrings1, l00F8, 1)
                Case "%"
                  l00B2 = Val(Mid$(gDatStrings1, l00F8 + 1, 1))
                  l00AC = l00AC & fn085C(l00B2)
                  If l00B2 = 10 Or l00B2 < 0 Then
                    l00F8 = l00F8 + 2
                  Else
                    l00F8 = l00F8 + 1
                  End If
                Case Else
                  DbgMileStone "unhandled const type"
                End Select
              Case Else
                Get g_hFileExe, , l00B2
                l00AC = l00AC & fn085C(l00B2)
              End Select
              gVACmdEnd = gVACmdEnd + 2
              l00FA = l00FA - 1
              l00F8 = l00F8 + 1
            Loop
          End If
          If gVACmdEnd > gVAModuleEnd Then l00AE = 0
          If l00AE = 0 Then
            gVACmdEnd = gVACmdStart
            gv013E = mc00DA
            DisOutput "<error abort>": WriteOutputTxt
            Exit Do
          End If
          DisOutput l00AC: WriteOutputTxt
        Loop
      Case mc005A To mc005C
        
        
        Do While gVACmdEnd < gv0144
          If gVACmdEnd = gVAModuleEnd Then Exit Do
          
          DisasmLine
          
          If gVACmdEnd = gVACmdStart Then
            Get g_hFileExe, gVAToRaw + gVACmdEnd, gSingleChar
            l00AC = Hex$(gVACmdEnd + gVANegOffsetCurMod) & ConCatHex8(Asc(gSingleChar))
            DisOutput l00AC: WriteOutputTxt
            gVACmdEnd = gVACmdEnd + 1
          End If
          
         'Reached End of Module?
          If gVACmdEnd > gVAModuleEnd Then
            
            If gv161C(gv014C).VOffset <> gVAModuleEnd Then
              sub00AC
              If gVAModuleEnd > gv0144 Then gVAModuleEnd = gv0144: gv014C = 0
            Else
              Exit Do
            End If
          
          End If
        Loop
      End Select
      
      If gVACmdEnd >= gv0144 Then Exit Do
      If gVACmdEnd < gVAModuleEnd Then
        WriteOutputTxt
        Exit Do
      End If
    End If
  Loop
  p00A6 = gv0144
Exit Sub

LD596:
  Get g_hFileExe, , l00BC

LD5AC:
  l00BE = ToUnSigned(l00BC)
If gIsVB_Ver Then
  l00BA = l00BE + gModuleVAStart
Else
  l00BA = gVACmdEnd + l00BE + 2
  l00BE = l00BA - gModuleVAStart
End If
  l00C4 = fn1D12(l00BA)
  l00AC = l00AC & " " & fn1D20(l00C4)
Return

LD652:
  If gVACmdEnd = gVAModuleEnd Then
    l00C6 = gv0150
    Select Case gv014E
    Case mc005E, mc014C
      l00AC = l00AC & " " & fn2023(gv0150)
    Case Else
      MsgBox "unexpected fixup"
      l00AE = 0: Return
    End Select
    gVACmdEnd = gVACmdEnd + 2
    Seek g_hFileExe, gVAToRaw + gVACmdEnd + 2
    l00FA = l00FA - 1
    sub019B
    If gVAModuleEnd > gv0144 Then gVAModuleEnd = gv0144: gv014C = 0
  Else
    l00C6 = 0
    Get g_hFileExe, , gv0120
    gVACmdEnd = gVACmdEnd + 2: l00FA = l00FA - 1
    If gv0120 = 0 Then
      l00AC = l00AC & " NULL"
    ElseIf fn087D(gv0120) = 0 Then
      l00AC = l00AC & " 0:" & toHex16(fn088F(gv0120))
    Else
      ErrBeep
      l00AE = 0
      l00AC = l00AC & "?" & Hex$(gv0120) & "?"
    End If
  End If
Return

LD876:
  If gVACmdEnd = p00A6 + 4 Then
    sub0167 g_hFileExe, gVACmdEnd, " statics:§l"
  End If
LD8C4:
  sub0167 g_hFileExe, gVACmdEnd, "w statics§"
  sub0167 g_hFileExe, gVACmdEnd, "w globals§"
  sub0167 g_hFileExe, gVACmdEnd, "l global base§"
Return
End Sub

Sub sub0134()
  ShowMsgBox "Feature not available in this version"
End Sub

Sub sub014F(p0152 As Integer, p0154 As Integer)
  gv013C = p0152
  gVACmdEnd = fn202F(p0152)
  gv0140 = gReloc_Arr(gv013C).Count
  If p0154 = 0 Then
    p0154 = gv0140
  End If
  gv0144 = fn202F(p0154)
  sub0115 gVACmdEnd
  gVAToRaw = gVAToRaw2
  gVANegOffsetCurMod = -gVACmdEnd
  gv014C = fn1624(gVACmdEnd)
  sub00AC
  If gVAModuleEnd > gv0144 Then gVAModuleEnd = gv0144: gv014C = 0
End Sub

Sub sub015B(p015C As Integer)
Dim l015E As String
' Const mc0168 = 32 ' &H20%
  If p015C < 0 Then
    If Len(gDisOutputTxt) = 0 Then DisOutput Hex$(gVACmdEnd + gVANegOffsetCurMod)
  End If
  If gVACmdEnd < gVAModuleEnd Then
    If p015C < 0 Then
      Get g_hFileExe, gVAToRaw + gVACmdEnd, gv0120
      If gv0120 Then
        DisOutput " non-null: " & fn086D(gv0120)
      Else
        DisOutput " NULL"
      End If
      WriteOutputTxt
    End If
    gVACmdEnd = gVACmdEnd + 4
    Exit Sub
  End If
  If p015C < 0 Then
    Select Case gv014E
    Case mc005E
      l015E = " ptr "
    Case mc0058
      l015E = " SEG "
    Case mc014C
      l015E = " PTR "
    End Select
    l015E = l015E & fn2023(gv0150)
    DisOutput l015E: WriteOutputTxt
  End If
  Select Case gv014E
  Case mc005E, mc014C: gVACmdEnd = gVAModuleEnd + 4
  Case mc0058: gVACmdEnd = gVAModuleEnd + 2
  End Select
  sub019B
  If gVAModuleEnd > gv0144 Then gVAModuleEnd = gv0144: gv014C = 0
End Sub

Sub sub0167(p0172 As Integer, p0174 As Long, ByVal p0176 As String)
Dim l0178
Dim l017A
Dim l017C As Integer
Dim l017E%
Dim l0180 As Long
Dim l0182 As String
  Seek p0172, p0174 + gVAToRaw
  For l0178 = 1 To Len(p0176)
    Select Case Mid$(p0176, l0178, 1)
    Case "b"
      Get p0172, , gSingleChar: l017C = Asc(gSingleChar): p0174 = p0174 + 1
      DisOutput ConCatHex8(l017C)
    Case "B"
      Get p0172, , l017E: p0174 = p0174 + 2 + l017E
      DisOutput fn085C(l017E) & ":"
      For l017A = 1 To l017E
        Get p0172, , gSingleChar: DisOutput ConCatHex8(Asc(gSingleChar))
      Next
    Case "w"
      Get p0172, , l017C: p0174 = p0174 + 2
      DisOutput fn085C(l017C)
    Case "l"
      Get p0172, , l0180: p0174 = p0174 + 4
      l0182 = Space$(8)
      RSet l0182 = Hex$(l0180)
      DisOutput " " & l0182
    Case "n"
      WriteOutputTxt
    Case "G"
      Get p0172, , gv0B84: p0174 = p0174 + Len(gv0B84)
      DisOutput " " & fn1285(gv0B84)
    Case "p"
      gVACmdEnd = p0174
      sub015B True
      p0174 = gVACmdEnd
      Seek p0172, p0174 + gVAToRaw
    Case "x"
      Exit Sub
    Case "z"
      l0182 = Space$(100)
LE148:
      Get p0172, , l0182
      l017E = InStr(l0182, Chr$(0))
      If l017E = 0 Then l017E = 1
      DisOutput " """ & Left$(l0182, l017E - 1) & """"
      p0174 = p0174 + l017E
      Seek p0172, p0174 + gVAToRaw
    Case "Z"
      l0182 = Space$(100)
      For l017E = 1 To 100
        Get p0172, , l017C
        If l017C And &HFF00 Then
          Seek p0172, p0174
          GoTo LE148
        End If
        If l017C = 0 Then Exit For
        Mid$(l0182, l017E, 1) = Chr$(l017C)
      Next
      DisOutput " """ & Left$(l0182, l017E - 1) & """L"
      p0174 = p0174 + l017E * 2
    Case "$"
      Get p0172, , l017E: p0174 = p0174 + 2
      If l017E > 0 Then
        l0182 = Space$(l017E)
        Get p0172, , l0182: p0174 = p0174 + l017E
        DisOutput " """ & StrCutZero(l0182) & """"
      Else
        DisOutput fn085C(l017E)
      End If
    Case "s"
      l017E = Val("&h" & Mid$(p0176, l0178 + 1, 2))
      l0178 = l0178 + 2
      l0182 = Space$(l017E)
      Get p0172, , l0182: p0174 = p0174 + l017E
      l0182 = "'" & l0182 & Chr$(0)
      LSet l0182 = StrCutZero(l0182) & "'"
      DisOutput " " & l0182
    Case "S"
      Get p0172, , gSingleChar: l017E = Asc(gSingleChar): p0174 = p0174 + l017E + 1
      l0182 = Space$(l017E)
      Get p0172, , l0182
      l017E = InStr(l0182, Chr$(0))
      If l017E Then l0182 = Left$(l0182, l017E - 1)
      DisOutput " '" & l0182 & "'"
    Case "U"
      Get p0172, , l017E: p0174 = p0174 + 2 + 2 * l017E
      l0182 = ""
      While l017E > 0
        Get p0172, , l017C
        l0182 = l0182 & Chr$(l017C)
        l017E = l017E - 1
      Wend
      DisOutput " '" & l0182 & "'"
    Case "a"
      If p0174 And 1 Then
        p0174 = p0174 + 1
        Seek p0172, p0174 + gVAToRaw
      End If
    Case "A"
        p0174 = (p0174 + 3) And Not 3
        Seek p0172, p0174 + gVAToRaw
    Case "§"
      l0178 = l0178 + 1
      l017A = InStr(l0178, p0176, "§")
      DisOutput Mid$(p0176, l0178, l017A - l0178)
      l0178 = l017A
    Case Else
      l017A = InStr(l0178, p0176, "§")
      If l017A Then
        DisOutput Mid$(p0176, l0178, l017A - l0178)
        l0178 = l017A
      Else
        DisOutput Mid$(p0176, l0178, 1)
      End If
    End Select
  Next
  WriteOutputTxt
Exit Sub
End Sub

Sub sub019B()
  If gv014C Then
    gv014C = fn164F(gv014C)
    sub00AC
  Else
    gVAModuleEnd = gv0144
  End If
End Sub

Sub sub01AB(p0158 As Integer)
  If p0158 Then
    sub01BA g_hFileExe, gVAToRaw, gVACmdEnd, gVAModuleEnd
  Else
    gVACmdEnd = gVAModuleEnd
  End If
  sub015B p0158
  sub015B p0158
  sub015B p0158
  If p0158 Then
    sub01BA g_hFileExe, gVAToRaw, gVACmdEnd, gVACmdEnd + 4
  Else
    gVACmdEnd = gVACmdEnd + 4
  End If
If gIsVB_Ver Then
  sub015B p0158
End If
End Sub

Sub sub01CA()
End Sub

Sub sub01D5(p1%, p2%)
Stop
  ShowMsgBox "Feature not available in this version"
End Sub

Sub sub01DF()
  ShowMsgBox mc0168
End Sub

Sub sub01E9()
  ShowMsgBox mc0168
End Sub

Sub sub01F3()
  ShowMsgBox mc0168
End Sub

Sub sub0200()
  ShowMsgBox mc0168
End Sub

Sub MakeProject(Dummy%)
Stop
End Sub

Sub sub0243()
End Sub

Sub sub0251(p01A8 As Integer, p01AA As Integer)
' Const mc01B8 = 18 ' &H12%
If g_IsDemoVersion Then
  ShowMsgBox "Feature not available in demo version"
Else
  sub025C p01A8
  gVANegOffsetCurMod = -g_FileExeOffset
  Select Case p01AA
  Case gv011A
    sub0268 g_hFileExe, g_FileExeOffset, gv02C0
  Case gc011C
    sub0272 g_hFileExe, g_FileExeOffset, gv02C0
  Case gc011E
    Select Case gv02A2.M36B3_count
    Case mc01B8
      Get g_hFileExe, g_FileExeOffset, gv0120
      DebugPrintVar gv0120 = &H10, "GUID size"
      Get g_hFileExe, , gv0B84
      DisOutput fn1285(gv0B84): WriteOutputTxt
    Case Is > 0
      If Left$(fn040F(p01A8), 8) = "_IID_FRM" Then
        Get g_hFileExe, g_FileExeOffset, gv0120
        DisOutput Hex$(gv0120)
        Get g_hFileExe, , gSingleChar: DisOutput ConCatHex8(Asc(gSingleChar))
        Get g_hFileExe, , gSingleChar: DisOutput ConCatHex8(Asc(gSingleChar))
        DisOutput " " & fn040F(p01A8)
        WriteOutputTxt
        sub0268 g_hFileExe, g_FileExeOffset + 6, g_FileExeOffset + 4 + gv0120
      Else
        sub0268 g_hFileExe, g_FileExeOffset, gv02C0
      End If
    Case Else
      sub0268 g_hFileExe, g_FileExeOffset, gv02C0
    End Select
  End Select
End If
End Sub

Sub sub027C(p01BC As Integer, ByVal p01BE As String, p01C0 As Integer)
Dim l01C2 As Integer
Dim l01C4 As Long
Dim l01C6 As Integer
Dim l01C8 As Integer
Dim l01CA As String
Dim l01CC
Dim l01CE As Integer
Dim l01D0 As Integer
' Const mc01DA = 4096 ' &H1000%
Static m01DC As Integer
' Const mc01E0 = 63 ' &H3F%
' Const mc01E2 = 32 ' &H20%
' Const mc01E4 = 1 ' &H1%
' Const mc01E6 = 2 ' &H2%
' Const mc01E8 = 18 ' &H12%
' Const mc01EA = 3 ' &H3%
' Const mc01EC = 4 ' &H4%
' Const mc01EE = 5 ' &H5%
' Const mc01F0 = 6 ' &H6%
' Const mc01F2 = 7 ' &H7%
' Const mc01F4 = 8 ' &H8%
' Const mc01F6 = 9 ' &H9%
' Const mc01F8 = 10 ' &HA%
' Const mc01FA = 11 ' &HB%
' Const mc01FC = 1 ' &H1%
' Const mc01FE = 2 ' &H2%
' Const mc0200 = 3 ' &H3%
' Const mc0204 = 10 ' &HA%
' Const mc0206 = 4 ' &H4%
' Const mc0208 = 5 ' &H5%
' Const mc021A = 261 ' &H105%
' Const mc021C = 9 ' &H9%
  l01C4 = gEXEdefEntries(p01BC).M689A
  If l01C4 < 0 Then
    If (gEXEdefEntries(p01BC).M6894.M36B3_count And mc01DA) = 0 Then Exit Sub
  End If
  l01C2 = gEXEdefEntries(p01BC).M6894.M42D9
  Select Case l01C2 And &H1F
  Case m01DC
    l01C6 = gEXEdefEntries(p01BC).Size
    If l01C6 = 0 Then
      l01D0 = fn10AB(p01BC)
    Else
      l01D0 = l01C6
    End If
    l01C2 = gEXEdefEntries(p01BC).M6894.M36B3_count And mc01E0
    If l01C2 And mc01E2 Then
      DbgMileStone "prefix array?"
    End If
    Select Case l01C2 And &H1F
    Case mc01E4: l01CA = "Byte"
    Case mc01E6, mc01E8
      If l01C6 = 0 Then l01C6 = l01D0
      l01CA = "String *" & Str$(l01C6)
    Case mc01EA: l01CA = "Boolean"
    Case mc01EC: l01CA = "Integer"
    Case mc01EE: l01CA = "Long"
    Case mc01F0: l01CA = "Single"
    Case mc01F2: l01CA = "Double"
    Case mc01F4: l01CA = "Currency"
    Case mc01F6: l01CA = "Date"
    Case mc01F8: l01CA = "Variant"
    Case mc01FA
      DbgMileStone "type name?"
      l01CA = "Type#" & Hex$(gEXEdefEntries(p01BC).M6894.M42D1)
    End Select
    If Len(l01CA) Then
      l01CA = " As " & l01CA
      If l01C2 And mc01E2 Then
        l01CA = "()" & l01CA
      End If
    Else
      If p01C0 Then
        If l01C6 = 0 Then l01C6 = l01D0
        Select Case l01C6
        Case 1: l01CA = "Byte"
        Case 2: l01CA = "Integer"
        Case 4: l01CA = "Long"
        Case 8: l01CA = "Double"
        Case Else
          l01CA = "any" & Format$(l01C6)
        End Select
        l01CA = " As " & l01CA
      Else
        l01CA = " As *" & Format$(l01C6)
        If l01C6 = 0 Then
          l01CA = l01CA & "<=" & Format$(l01D0)
        End If
      End If
    End If
  Case mc01FC
    l01CA = "$"
  Case mc01FE
    l01CA = " As Variant"
  Case mc0200
    l01CA = " As Object"
    l01C2 = gEXEdefEntries(p01BC).M6894.M42D1
    If l01C2 < 0 Then
      l01CA = l01CA & fn104E(-l01C2)
    End If
  Case mc0204, mc0206
    l01CA = " As String *" & Str$(gEXEdefEntries(p01BC).M6894.M42D1)
  Case mc0208
    l01CA = "("
    l01C2 = gEXEdefEntries(p01BC).M6894.M42D1
    l01C8 = gEXEdef_Arr1(l01C2).M34F5
    l01CE = gEXEdef_Arr1(l01C2).M55C2
    For l01CC = 1 To l01C8
      If l01CC > 1 Then l01CA = l01CA & ", "
      l01C6 = gEXEdef_Arr2(l01CE + l01CC).M6968
      l01CA = l01CA & l01C6 & " To "
      l01CA = l01CA & gEXEdef_Arr2(l01CE + l01CC).M54E9 + l01C6 - 1
    Next
    l01CA = l01CA & ") As *" & Format$(gEXEdef_Arr1(l01C2).M3503)
  Case mc021A
    l01CA = "() As any?"
  Case mc021C
    l01C2 = gEXEdefEntries(p01BC).M6894.M42D1
    If l01C2 < 0 Then
      l01CA = " As " & fn104E(-l01C2)
    Else
      Stop
      DisOutput " As Type" & toHex16(CLng(p01BC)) 'CLng(
    End If
  End Select
  If Len(l01CA) Then
    l01CA = " " & fn1B19(p01BC) & l01CA
    DisOutput p01BE & l01CA
  End If
End Sub

Sub sub028D(p0220 As Integer, p0222 As Integer)
Dim l0224 As Integer
Dim l0226 As Integer
Dim l0228 As Integer
Dim l022A As Integer
Dim l022C As Long
Dim l022E As Long
' Const mc0240 = 28 ' &H1C%
' Const mc0242 = 29 ' &H1D%
' Const mc024C = 256 ' &H100%
  If p0220 < 0 Then Exit Sub
  sub0299 p0220, True
  m004A = gL2Max + p0220
  m004C = 0
  If hFile_Out Then
    gDisOutputTxt = " ": WriteOutputTxt
  End If
  
  l022A = fn106F(p0220, l022C, l022E)
  l0226 = gEXEdef_Subs(p0220).M4277
  l0228 = l0226 And &H3F
  Select Case l0226
  Case 0
    DisOutput "Sub "
  Case Else
    DisOutput "Function "
  End Select
  
  If Not p0222 Then
    DisOutput "@" & toHex16(CLng(p0220)) & " " 'CLng(
    Stop
  End If
  
  DisOutput fn0C98(p0220) & " ("
  l0224 = g_gbl(m004A).M4950
  If l0224 Then
    DisOutput "_": WriteOutputTxt
    Do While l0224
      If p0222 And gEXEdefEntries(l0224).M689A = l022C Then
        l0224 = gEXEdefEntries(l0224).Count
      Else
        sub027C l0224, "  ", p0222
        l0224 = gEXEdefEntries(l0224).Count
        If l0224 Then
            DisOutput ",_": WriteOutputTxt
        End If
      End If
    Loop
  End If
  DisOutput ")"
  If l0226 Then
    DisOutput " As "
    Select Case l022A
    Case mc01E4: DisOutput "Byte"
    Case mc01EA: DisOutput "Boolean"
    Case mc01EC: DisOutput "Integer"
    Case mc0240: DisOutput "String"
    Case mc0242: DisOutput "Object"
    Case mc01FA
      If l022C Then
        l0224 = gv3126(l022C)
        If l0224 Then
          l0224 = gEXEdefEntries(l0224).M6894.M42D1
        End If
      End If
      If l0224 = 0 Then
        If l022E Then
          l0224 = gv3126(l022E)
          If l0224 Then
            l0224 = gEXEdefEntries(l0224).M6894.M42D1
          End If
        End If
      End If
      If l0224 Then
        DisOutput fn104E(-l0224)
      Else
        DisOutput "Type...": DbgMileStone "which type?"
      End If
    Case mc01F0: DisOutput "Single"
    Case mc01EE: DisOutput "Long"
    Case mc01F4: DisOutput "Currency"
    Case mc01F6: DisOutput "Date"
    Case mc01F2: DisOutput "Double"
    Case mc01F8: DisOutput "Variant"
    Case Else
      DisOutput "any" & CStr(l0228 And &H1F)
    End Select
  End If
  WriteOutputTxt
  m004A = gL3Max + p0220
  l0224 = g_gbl(m004A).M4950
  If l0228 Then
    If l0224 > 0 And gEXEdefEntries(l0224).M689A = l022E Then
      If p0222 Then
        sub027C l0224, "'fn", p0222
      Else
        sub027C l0224, "Dim", p0222
      End If
      WriteOutputTxt
      l0224 = gEXEdefEntries(l0224).Count
    End If
  End If
  While l0224 > 0
    If gEXEdefEntries(l0224).M6894.M36B3_count And mc024C Then
      sub027C l0224, "'ByVal", p0222
      WriteOutputTxt
    Else
      sub027C l0224, "Dim", p0222
      WriteOutputTxt
    End If
    l0224 = gEXEdefEntries(l0224).Count
  Wend
  m004A = gL1Max + gEXEdef_Subs(p0220).ModuleIndex
  l0224 = g_gbl(m004A).M4950
  While l0224
    If gEXEdefEntries(l0224).M6877 = p0220 Then
      sub027C l0224, "Static", p0222
      WriteOutputTxt
    End If
    l0224 = gEXEdefEntries(l0224).Count
  Wend
  If p0222 Then
    sub00C9
  End If
  If l0226 Then
    DisOutput "End Function"
  Else
    DisOutput "End Sub"
  End If
  WriteOutputTxt
End Sub

Sub sub02A8(p0250%, p0252 As Integer)
Dim l0254 As Integer
  DisOutput " ": WriteOutputTxt
  DisOutput "Type " & fn1C33(p0250) & " ' *" & Str$(gEXEdefEntries2(p0250).Size)
  WriteOutputTxt
  m004A = fn1C9B(p0250, False)
  l0254 = g_gbl(m004A).M4950
  While l0254
    sub027C l0254, "  ", p0252
    WriteOutputTxt
    l0254 = gEXEdefEntries(l0254).Count
  Wend
  DisOutput "End Type": WriteOutputTxt
End Sub

