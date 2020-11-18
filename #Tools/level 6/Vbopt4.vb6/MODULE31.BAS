Attribute VB_Name = "MODULE31"
' Module31
Option Explicit
Dim m000C As Integer
Dim m000E As Integer
Dim m0010 As Integer
Dim m0012 As Long
Dim m0018() As T588E
Dim m004A As String
Dim m004E As String
Dim m0052 As String
Dim MForms As T58CA
Dim m00C2 As T594D
Dim m0120 As T59F0
Dim m0172 As T5A33
Dim m01A6 As T5AAE
Dim m01E0 As T5B23
Dim m021A As T5BAB
Dim m025C As T5BFF
Dim m027A As T5C30
Dim m02A0 As T5CA0
Dim m02FC As T5D6E
Dim m0352 As T5DC9
Dim m035C As T5DD5
Dim m0366 As T5DE3
Dim m0378 As T5E0D
Dim m03A2 As T5E30
Dim m03B0 As T5E49
Const mc043A = 34 ' &H22%
Const mc045E = 4 ' &H4%
Const mc04EA = 8 ' &H8%
Const mc0508 = 16 ' &H10%
Const mc050A = 31 ' &H1F%
Const mc050C = 16 ' &H10%
Const mc0510 = 9 ' &H9%
Const mc0518 = 11 ' &HB%
Const mc051A = 11 ' &HB%
Const mc051C = 12 ' &HC%
Const mc051E = 12 ' &HC%
Const mc054A = 2 ' &H2%
Const mc054C = 4 ' &H4%
Const mc054E = 11 ' &HB%
Const mc0550 = 17 ' &H11%
Const mc0552 = 15 ' &HF%
Const mc055E = 5 ' &H5%
Const mc0560 = 12 ' &HC%
Const mc0562 = 6 ' &H6%
Const mc0566 = 9 ' &H9%
Const mc0578 = 1 ' &H1%
Const mc057A = 2 ' &H2%
Const mc057E = 3 ' &H3%
Const mc0582 = 37 ' &H25%
Const mc0584 = 7 ' &H7%
Const mc0586 = 50 ' &H32%
Const mc0588 = 35 ' &H23%
Const mc05DA = 32 ' &H20%

Sub sub1DC1(p0406 As Long)
Dim l0408 As Integer
  If p0406 <= 0 Then Exit Sub
  m0052 = ReadZeroString(g_hFileExe, VAtoRaw(p0406))
  l0408 = RetOffsetSym(p0406, "'" & m0052 & "'", 0)
End Sub

Sub sub1DCE(p0412 As Long)
Dim l0414 As Integer
Dim l0416 As String
  If p0412 <= 0 Then Exit Sub
  l0416 = ReadZeroString(g_hFileExe, VAtoRaw(p0412))
  l0414 = RetOffsetSym(p0412, "'" & l0416 & "'", 0)
  p0412 = p0412 + Len(l0416) + 1
  If p0412 And 3 Then
    p0412 = (p0412 + 3) And Not 3
  End If
End Sub

Sub sub1DDC(VA As Long)
Dim l041A As Integer
Dim l041C As Long
Dim l041E As Long
Dim l0420 As Long
Dim l0422 As Long
Dim l042E As Long
Dim l0430 As String
Dim l0432 As Long
Dim l0438 As Integer
' Const mc043A = 34 ' &H22%
  l041A = RetOffsetSym(VA, m004A & ".Class3", 0)
sub0C11 g_decl_num, l041A
  Get g_hFileExe, gVAToRaw2 + VA, m021A
  l0422 = m021A.M5BDC
  l041C = m021A.M5BC8
  If l041C <> l0422 Then
    l041A = RetOffsetSym(l041C, m004A & ".Params", 0)
sub0C11 g_decl_num, l041A
m0018(g_decl_num).M58B2 = l041A
    l0420 = l041C + fn203B(l041A)
    Do While l041C < l0420
      Get g_hFileExe, gVAToRaw2 + l041C, l041E
      If l041E > 0 And l041E > gImageBase1 And l041E < gVOffset1 Then
        l041A = RetOffsetSym(l041E, m004A & ".para" & Hex$(l041E), 0)
        If l041A = 0 Then Exit Do
        Get g_hFileExe, gVAToRaw2 + l041E, m027A
        l0430 = "("
        l042E = m027A.M5C73
        l041A = CreateSym(l042E)
        If l041A Then
          l0432 = l042E + fn203B(l041A)
          Do While l042E < l0432
            Get g_hFileExe, gVAToRaw2 + l042E, l041E
            If l041E = 0 Then Exit Do
            sub1DC1 l041E
            l0430 = l0430 & m0052 & ","
            l042E = l042E + 4
          Loop
          If Len(l0430) > 1 Then
            Mid$(l0430, Len(l0430), 1) = ")"
          Else
            l0430 = "()"
          End If
          l041A = fn1E30(m027A.M5C73, l0430)
        End If
      ElseIf l041E = 0 Then
      Else
        Exit Do
      End If
      l041C = l041C + 4
    Loop
  End If
  l041C = m021A.M5BD0
  If l041C <> l0422 Then
    l041A = fn1E30(l041C, m004A & ".PubVars[]")
    If fn2023(l041A) <> "pNull" Then
sub0C11 g_decl_num, l041A
m0018(g_decl_num).M58BF = l041A
    End If
    l0420 = l041C + fn203B(l041A)
    Do While l041C < l0420
      Get g_hFileExe, gVAToRaw2 + l041C, l041E
      If l041E > 0 Then
        Get g_hFileExe, gVAToRaw2 + l041E, m025C
        sub1DC1 m025C.M5C0C
        l041A = RetOffsetSym(l041E, m004A & "." & m0052, 0)
      End If
      l041C = l041C + 4
    Loop
    l041C = m021A.M5BDC
    Get g_hFileExe, gVAToRaw2 + l041C, l041E
    If l041E Then
      l041A = RetOffsetSym(l041C, m004A & ".p9", 0)
      l0438 = fn203B(l041A) \ 4
      For l0438 = 1 To l0438
        l041A = RetOffsetSym(l041E, m004A & ".p9." & CStr(l0438), 0)
        Get g_hFileExe, , l041E
      Next
    Else
      l041A = fn1E30(l041C, m004A & "()")
    End If
  End If
  l041C = m021A.M5BE2
  If l041C <> l0422 Then
    Get g_hFileExe, gVAToRaw2 + l041C, l041E
    If l041E Then
      l041A = RetOffsetSym(l041C, m004A & ".mem10", 0)
      l041A = RetOffsetSym(l041E, m004A & ".mem10.1", 0)
      gReloc_Arr(l041A).M36B3_count = mc043A
      Get g_hFileExe, , l041E
      l041A = RetOffsetSym(l041E, m004A & ".mem10.2", 0)
      gReloc_Arr(l041A).M36B3_count = mc043A
    Else
      l041A = fn1E30(l041C, m004A & "()")
    End If
  End If
End Sub

Sub sub1DEA()
  DbgMileStone "see add_exeform4"
End Sub

Sub sub1DFA(p0444 As Long)
Dim l0446 As Integer
  Assume p0444 > 0, "GUID"
  Get g_hFileExe, gVAToRaw2 + p0444, gv0B84
  l0446 = RetOffsetSym(p0444, fn1285(gv0B84), 0)
  If fn203B(l0446) <> Len(gv0B84) Then
    l0446 = CreateSym(p0444 + Len(gv0B84))
  End If
End Sub

Sub sub1E07(ByVal p044E As Long, ByVal p0450 As Long)
Dim l0452 As Integer
  Assume p044E > 0, "modhdr4"
  While p0450 > 0
    Get g_hFileExe, gVAToRaw2 + p044E, m02A0
    Assume m02A0.M5CAF = &H40, "40"
    If m02A0.M5CAF <> &H40 Then
      Exit Sub
    End If
    sub1DC1 m02A0.M5CE5
    m004E = m004A & "." & m0052
    l0452 = RetOffsetSym(p044E, m004E, 0)
    sub1DFA m02A0.M5B37
    l0452 = RetOffsetSym(m02A0.M5CCB, m004E & ".Events", 0)
    p0450 = p0450 - 1
    p044E = p044E + Len(m02A0)
  Wend
End Sub

Sub sub1E13(p0456 As Long)
Dim l0458 As String
Dim l045A As Integer
' Const mc045E = 4 ' &H4%
  Assume p0456 > 0, "unistring"
  l0458 = fn1561(p0456)
  l045A = RetOffsetSym(p0456, l0458, mc045E)
  l045A = RetOffsetSym(p0456 - 4, "len", 0)
End Sub

Sub sub1E21(p0460 As Long)
Dim l0462 As Integer
Dim l0464 As Integer
Dim l0466 As Long
  Assume p0460 > 0, "ppGuid"
  Get g_hFileExe, gVAToRaw2 + p0460, l0466
  Assume l0466 > 0, "pGuid"
  Get g_hFileExe, gVAToRaw2 + l0466, gv0B84
  l0462 = RetOffsetSym(l0466, fn1285(gv0B84), 0)
  Assume l0462, "Guid"
  If fn203B(l0462) > Len(gv0B84) Then
    l0464 = CreateSym(l0466 + Len(gv0B84))
  End If
  l0462 = RetOffsetSym(p0460, "->" & fn2023(l0462), 0)
End Sub

Function fn1E30(p046A As Long, frmName As String) As Integer
Dim l046E As Integer
Dim l0470 As Long
  If p046A = 0 Then Exit Function
  l046E = CreateSym(p046A)
  fn1E30 = l046E
  
  If gReloc_Arr(l046E).subRef_M3C6A Then Exit Function
  Get g_hFileExe, gVAToRaw2 + p046A, l0470
  If l0470 = 0 Then
    sub0D9B l046E, "pNull"
  Else
    sub0D9B l046E, frmName
  End If
End Function

Sub sub1E3B(p0472 As Integer)
Dim l0474 As Integer
Dim l0478 As Long
Dim l047A As Long
  l047A = fn202F(p0472)
  l0478 = fn202F(gReloc_Arr(p0472).Count)
  l0474 = fn1624(l047A)
  While gv161C(l0474).VOffset < l0478
    sub1DC1 fn202F(gv161C(l0474).subRef_M3C6A)
    l0474 = fn164F(l0474)
  Wend
End Sub

Sub sub1E49(p048A%, ByVal p048C As String)
Dim l048E As Integer
Dim l0490 As Long
Dim l0492 As Long
Dim l0494 As Long
Dim l0496 As Long
Dim l0498 As Integer
Dim l049C As String
Dim l04A0 As Integer
  l0492 = fn202F(p048A)
  l04A0 = fn2052(p048A)
  If l04A0 Then
    l0496 = fn202F(l04A0)
  Else
    l0496 = gv1066
  End If
  If l0496 < l0492 Then
    Exit Sub
  End If
  l0494 = l0492
  l048E = fn1624(l0494)
  If l048E = 0 Then Exit Sub
  l0490 = gv161C(l048E).VOffset
  Do While l0490 < l0496
    l0498 = gv161C(l048E).subRef_M3C6A
    l049C = p048C & fn0842((l0490 - l0492) \ 4)
      l0498 = fn1E30(fn202F(l0498), l049C)
      If fn2023(l0498) = l049C Then
        sub1E49 l0498, fn2023(l0498)
      End If
    l048E = fn164F(l048E)
    Assume l048E, "pat follows"
    l0490 = gv161C(l048E).VOffset
  Loop
End Sub

Sub sub1E55(p04A8&)
Dim l04AA As String
Dim l04AC As Integer
  l04AA = fn1561(p04A8)
  l04AC = RetOffsetSym(p04A8, l04AA, mc045E)
End Sub

Sub sub1E61()
  DebugPrintVar False, "unused"
End Sub

Sub sub1E75()
Dim l04B4 As Integer
Dim l04B6 As Integer
Dim l04BA As Long
Dim l04BC As Integer
Dim l04C0%
Dim l04C2
Dim l04C4 As Integer
Dim l04C6 As Long
Dim l04C8 As Integer
Dim l04CA As Long
Dim l04CC As Long
Dim l04CE As String
Dim l04D0 As Integer
Dim l04D2 As Long
Dim l04D4 As Long
Dim l04D6 As Long
' Const mc04EA = 8 ' &H8%
  frmScan.control4 = "Code"
  frmScan.control2 = ".text"
  DoEvents
Screen.MousePointer = 11
  l04CC = fn202F(m000E)
  sub0115 l04CC
  sub014F m000E, m0010
  sub1E89 l04CC, gv0144, False
  l04B6 = m000E
  While l04B6 <> m0010
    Get g_hFileExe, fn1483(l04B6), l04C0
    If l04C0 = &H8B55 Then
      l04BC = fn0DF5(l04B6, 0)
      l04B4 = fn1624(fn202F(l04B6))
      For l04C2 = 1 To 3
        If l04B4 = 0 Then
          Exit For
        Else
          l04BA = gv161C(l04B4).VOffset
          If l04BA > fn202F(gReloc_Arr(l04B6).Count) Then
            Exit For
          End If
          Get g_hFileExe, fn15A9(l04BA) - 9, l04D4
        End If
        If l04D4 = &H89575653 Then
          l04C4 = gv161C(l04B4).subRef_M3C6A
          l04C6 = fn202F(l04C4)
          l04CE = "sub" & CStr(gEXEdef_Arr5Count)
          sub0D9B l04C4, l04CE & ".vtbl"
          l04C8 = fn2052(l04C4)
          l04CA = fn202F(l04C8)
          l04B4 = fn1624(l04C6)
          Do While l04B4
            l04D6 = gv161C(l04B4).VOffset
            If l04D6 >= l04CA Then Exit Do
            l04D0 = gv161C(l04B4).subRef_M3C6A
            l04C0 = (l04D6 - l04C6) \ 4
            sub0D9B l04D0, l04CE & ".p" & CStr(l04C0)
            If l04C0 = 5 Then
              l04D2 = fn202F(l04D0)
              l04C8 = fn2052(l04D0)
              l04CA = fn202F(l04C8)
              l04B4 = fn1624(l04D2)
              l04C0 = 0
              Do While l04B4
                l04D6 = gv161C(l04B4).VOffset
                If l04D6 >= l04CA Then Exit Do
                l04C0 = l04C0 + 1
                sub0D9B gv161C(l04B4).subRef_M3C6A, "line" & CStr(l04C0)
                l04B4 = fn164F(l04B4)
              Loop
              Exit For
            End If
            l04B4 = fn164F(l04B4)
          Loop
          Exit For
        End If
        l04B4 = fn164F(l04B4)
      Next
    Else
      gReloc_Arr(l04B6).M36B3_count = mc04EA
    End If
    l04B6 = fn2052(l04B6)
  Wend
Screen.MousePointer = 0
End Sub

Sub sub1E96()
Dim l04EC As Long
Dim l04EE As Integer
Dim l04F0 As Integer
Dim l04F2 As Integer
Dim l04F4 As Integer
Dim l04F6 As Long
Dim l04F8 As String
' Const mc0508 = 16 ' &H10%
' Const mc050A = 31 ' &H1F%
' Const mc050C = 16 ' &H10%
' Const mc0510 = 9 ' &H9%
Dim l0512&
Dim l0514&
' Const mc0518 = 11 ' &HB%
' Const mc051A = 11 ' &HB%
' Const mc051C = 12 ' &HC%
' Const mc051E = 12 ' &HC%
  l04EC = fn202F(gEntryPoint_Page)
  l04EE = fn1624(l04EC)
  Do While l04EE
    l04EC = gv161C(l04EE).VOffset
    l04F0 = gv161C(l04EE).subRef_M3C6A
    l04F2 = fn2047(l04F0)
    If l04F2 Then
      l04F4 = gv161C(gv161C(l04EE).M54BB).subRef_M3C6A
      Select Case gEXEdef_Arr4_imports(l04F2).TypeCode_M5F96
      Case mc0508 To mc050A
        Get g_hFileExe, gVAToRaw2 + l04EC - 9, l04F6
        If l04F6 = &H4244481 Then
          l04F0 = CreateSym(l04EC - 9)
          Get g_hFileExe, , l04F6
          l04F4 = sub09DB(l04F0, 0, l04F2, mc050C + (gEXEdef_Arr4_imports(l04F2).TypeCode_M5F96 - mc0508), l04F6)
          l04F0 = CreateSym(l04EC + 6)
        ElseIf l04F6 Then
        End If
      Case mc0510
        If gEXEdef_Arr4_imports(l04F2).TypeCode_M5F96 = 9 Then
          Get g_hFileExe, gVAToRaw2 + l04EC - 12, l04F6
          If l04F6 = &H274C00B Then
            l04F0 = CreateSym(l04EC - 17)
            gReloc_Arr(l04F0).M36B3_count = mc04EA
            Get g_hFileExe, gVAToRaw2 + l04EC - 5, l04F6
            Get g_hFileExe, gVAToRaw2 + l04F6, l0512
            Get g_hFileExe, , l0514
            sub1DC1 l0512
            l04F8 = m0052
            sub1DC1 l0514
            l04F8 = l04F8 & ":" & m0052
            l04F6 = fn0DE7(l04F4, l04F8, l04F0)
            gReloc_Arr(l04F4).subRef_M3C6A = -l04F0
          End If
        End If
      Case mc0518
        If gReloc_Arr(l04F4).subRef_M3C6A Then
        Else
          l04F0 = CreateSym(l04EC - 6)
          l04F4 = sub09DB(l04F4, 0, l04F0, mc051A, 0)
          l04F0 = CreateSym(l04EC + 6)
        End If
      Case mc051C
        If gReloc_Arr(l04F4).subRef_M3C6A Then
            Exit Do
        Else
          l04F0 = CreateSym(l04EC - 8)
          Get g_hFileExe, gVAToRaw2 + l04EC - 14, l04F6
          l04F4 = sub09DB(l04F4, 0, l04F0, mc051E, l04F6)
          l04F2 = CreateSym(l04EC - 15)
          gReloc_Arr(l04F2).M36B3_count = gReloc_Arr(l04F0).M36B3_count
          gReloc_Arr(l04F2).subRef_M3C6A = gReloc_Arr(l04F0).subRef_M3C6A
          l04F0 = CreateSym(l04EC + 5)
        End If
      End Select
    End If
    l04EE = fn164F(l04EE)
  Loop
End Sub

Function fn1EA9(p0522 As Long) As Integer
  Get g_hFileExe, p0522, gSingleChar: gv34F6 = Asc(gSingleChar)
  Select Case gv34F6
  Case &HFF: gv34F6 = &H100
  Case &HFE: gv34F6 = &H200
  Case &HFD: gv34F6 = &H300
  Case &HFC: gv34F6 = &H400
  Case &HFB: gv34F6 = &H500
  Case Else
    fn1EA9 = 1
    Exit Function
  End Select
  Get g_hFileExe, , gSingleChar
  gv34F6 = Asc(gSingleChar) + gv34F6
  fn1EA9 = 2
End Function

Sub sub0D0B()
  DebugPrintVar False, "unused"
End Sub

Sub sub0CEE()
Dim l052E As Integer
Dim l0530 As Integer
Dim l0536 As Long
Dim l0538 As Integer
Dim l053A As Integer
Dim l053C As Long
Dim l053E As String
' Const mc054A = 2 ' &H2%
' Const mc054C = 4 ' &H4%
' Const mc054E = 11 ' &HB%
' Const mc0550 = 17 ' &H11%
' Const mc0552 = 15 ' &HF%
Dim l0556 As Long
Dim i
Dim l055A As Integer
Dim l055C As Integer
' Const mc055E = 5 ' &H5%
' Const mc0560 = 12 ' &HC%
' Const mc0562 = 6 ' &H6%
' Const mc0566 = 9 ' &H9%
' Const mc0578 = 1 ' &H1%
' Const mc057A = 2 ' &H2%
' Const mc057E = 3 ' &H3%
' Const mc0582 = 37 ' &H25%
' Const mc0584 = 7 ' &H7%
' Const mc0586 = 50 ' &H32%
' Const mc0588 = 35 ' &H23%
  gv3754 = Len(gv20E4)
  load_400ss
  l052E = gv2EF8(0)
  l0536 = fn202F(l052E)
  sub0115 l0536
  l0538 = 0
  Get g_hFileExe, fn1483(l052E), MForms
  l0530 = fn0E06(1, l052E, &H30, l0538)
  l0530 = fn0E06(mc054A, l052E, &H4C, l0538)
  l0530 = fn0E06(mc054C, l052E, &H54, l0538)
  l052E = gv2EF8(1)
  Get g_hFileExe, fn1483(l052E), m00C2
  l0538 = 0
  l0530 = fn0E06(mc054E, l052E, 4, l0538)
  Assume gv2EF8(mc054E), "Tab11"
  l0530 = fn0E06(mc0550, l052E, &H18, l0538)
  m0012 = fn202F(gv2EF8(mc0550))
  l0530 = fn0E06(mc0552, l052E, &H234, l0538)
  l0530 = RetOffsetSym(fn202F(l052E) + m00C2.M59A8, "path.vbp", 0)
  m000C = RetOffsetSym(m00C2.M596A, "CodeStart", 0)
  m000E = CreateSym(m00C2.M596A + 4)
  m0010 = RetOffsetSym(m00C2.M5978, "CodeEnd", 0)
  If m00C2.M599E Then
'    MsgBox "assume native?"
    gIsVB_Ver = -5
  Else
    gIsVB5 = 5
  End If
  
  l052E = gv2EF8(mc0552)
  If l052E And l052E <> gv2EF8(0) Then
    l0556 = gVAToRaw2 + fn202F(l052E)
    Do
      Get g_hFileExe, l0556, m0352
      l0556 = Seek(g_hFileExe)
      Select Case m0352.ResObjName
      Case 6
        Get g_hFileExe, gVAToRaw2 + m0352.VOffset, m035C
        sub1DFA m035C.M5B37
        l0530 = RetOffsetSym(m035C.M598A, "Global.data", 0)
      Case 7
        Get g_hFileExe, gVAToRaw2 + m0352.VOffset, m0366
        sub1DC1 m0366.M5DF5
        l053E = m0052
        sub1DC1 m0366.M5E01
        l053E = l053E & ":" & m0052
        l0530 = RetOffsetSym(m0352.VOffset, l053E, 0)
        l0530 = RetOffsetSym(m0366.M598A, l053E & ".data", 0)
        If gIsVB_Ver < 0 Then
          l0530 = RetOffsetSym(m0352.VOffset + Len(m0366), l053E & ".code", 0)
          gReloc_Arr(l0530).M36B3_count = mc04EA
          l0530 = CreateSym(m0352.VOffset + Len(m0366) + 11)
          gReloc_Arr(l0530).M36B3_count = mc04EA
          l0530 = CreateSym(m0352.VOffset + Len(m0366) + &H19)
        End If
      Case Else
        Exit Do
      End Select
    Loop
  End If
  l052E = gv2EF8(mc054E)
  Get g_hFileExe, fn1483(l052E), m02FC
  ReDim m0018(m02FC.ModuleTableEntryCount)
  sub1DC1 m02FC.M5DB6
  l0536 = m02FC.M5D7B
  Assume l0536 > 0, "PrjRec"
  l0530 = fn0E06(mc055E, l052E, &H8, 0)
  Get g_hFileExe, fn1483(l0530), m0378
  l0536 = m0378.M5905
  If l0536 <> -1 Then
    Assume l0536 > 0, "Forms Table"
    l0530 = fn0E06(mc0560, l0530, &H10, 0)
  End If
  If m0378.M5E18 > 0 Then
    l0536 = m0378.M5E18
    If l0536 > 0 Then
      sub1DCE l0536
    Else
      l0536 = 0
    End If
  Else
    l0536 = m02FC.M5D7B + Len(m0378)
  End If
  If l0536 Then
    Get g_hFileExe, VAtoRaw(l0536), m03A2
    l052E = fn0C32(mc0562, l0536)
    sub1DFA m03A2.M5B37
    l0530 = fn0E06(mc0566, l052E, 0, 0)
    If l0530 Then
      Get g_hFileExe, fn1483(l0530), m03B0
      sub1DFA m03B0.M5B37
      sub1DC1 m03B0.M5E62
      sub1DC1 m03B0.M5E6A
      Do
        l0536 = l0536 + Len(m03A2)
        Get g_hFileExe, VAtoRaw(l0536), m03A2
        If m03A2.M5E3E <> fn202F(l0530) Then Exit Do
        Assume RetOffsetSym(l0536, "PrjPtrs...", 0), "void"
        sub1DFA m03A2.M5B37
      Loop
    Else
      MsgBox "expected olbrec"
    End If
  End If
  l0536 = m02FC.M5DAD
  Assume m02FC.ModuleTableEntryCount = m02FC.M5DA0, "dup modcount"
  gEXEdef1 = m02FC.ModuleTableEntryCount
  ReDim g_decl_objs(gEXEdef1), gv28BA(gEXEdef1)
  g_decl_num = 0
  Do
    Get g_hFileExe, gVAToRaw2 + l0536, m0172
    Select Case m0172.Flags And &HFFFFFF83
    Case &H18001
      l055A = mc0578
      l053E = ".BAS"
    Case &H18003
      l055A = mc057A
      l053E = ".CLS"
      gEXEdef2 = gEXEdef2 + 1
    Case &H18083
      l055A = mc057E
      l053E = ".FRM"
      gFormsCount = gFormsCount + 1
    Case Else
      Assume gEXEdef1 = g_decl_num, "all mods found"
      Exit Do
    End Select
    g_decl_num = g_decl_num + 1
    sub1DC1 m0172.M5A71
    m004A = m0052
g_decl_objs(g_decl_num).M3F23 = m004A
g_decl_objs(g_decl_num).Name_M411A = m004A & l053E
g_decl_objs(g_decl_num).M36B3_count = l055A
    l052E = RetOffsetSym(l0536, m004A & ".Hdr2", 0)
    gReloc_Arr(l052E).M36B3_count = mc0582
sub0C11 g_decl_num, l052E
    
    Assume m0172.M5A51 > 0, "gbl layout"
    l052E = fn1E30(m0172.M5A51, m004A & ".Gbl")
    gReloc_Arr(l052E).M36B3_count = mc043A
g_decl_objs(g_decl_num).M4131 = l052E
    If m0172.M5A60 > 0 Then
      l052E = RetOffsetSym(m0172.M5A60, m004A & ".gbldata", 0)
      Assume m0012 > 0, "preset .data"
g_decl_objs(g_decl_num).M4150 = m0172.M5A60 - m0012
g_decl_objs(g_decl_num).M4126 = m0172.M5A60 - m0012
    End If
    Get g_hFileExe, gVAToRaw2 + m0172.M5A51, gv20E4
    l055C = gv20E4.M5EBD
    l052E = CreateSym(m0172.M5A51 + l055C)
    l055C = gv20E4.M5EC8
    If l055C And 2 Then l055C = l055C + 2
g_decl_objs(g_decl_num).Size = l055C
    If m0172.M5A58 Then
      l052E = fn1E30(m0172.M5A58, m004A & ".Stat")
      gReloc_Arr(l052E).M36B3_count = mc043A
g_decl_objs(g_decl_num).M415B = l052E
g_decl_objs(g_decl_num).M416F = m0172.M5A91
      Get g_hFileExe, gVAToRaw2 + m0172.M5A58, gv20E4
      l055C = gv20E4.M5EBD
      l052E = CreateSym(m0172.M5A58 + l055C)
      l055C = gv20E4.M5EC8
g_decl_objs(g_decl_num).statSize = l055C
    End If
    If m0172.M5A68 > 0 Then
      l052E = RetOffsetSym(m0172.M5A68, m004A & ".statdata", 0)
g_decl_objs(g_decl_num).M4178 = m0172.M5A68 - m0012
    End If
    If m0172.M5A86 And m0172.M5A7D > 0 Then
      l052E = RetOffsetSym(m0172.M5A86, m004A & ".PubSubs[]", 0)
m0018(g_decl_num).M58A6 = l052E
      gReloc_Arr(l052E).M36B3_count = mc0584
      sub1E3B l052E
      l0530 = CreateSym(m0172.M5A86 + m0172.M5A7D * 4)
    End If
    Get g_hFileExe, gVAToRaw2 + m0172.M5A42, m01A6
    l052E = RetOffsetSym(m0172.M5A42, m004A & ".Hdr3", 0)
    gReloc_Arr(l052E).M36B3_count = mc0586 + 3
sub0C11 g_decl_num, l052E
If gIsVB5 < 0 Then
  g_decl_objs(g_decl_num).M3F13 = m0172.M5A7D - m01A6.M5AEB
Else
  g_decl_objs(g_decl_num).M3F13 = m01A6.M5AEB
End If
g_decl_objs(g_decl_num).M3F1B = g_decl_objs(g_decl_num - 1).M3F1B + g_decl_objs(g_decl_num - 1).M3F13
m0018(g_decl_num).M3F13 = g_decl_objs(g_decl_num).M3F13
    If l055A = mc0578 Then
    Else
      Get g_hFileExe, , m01E0
      Assume m01E0.M5A9A = 1, "l12"
      l053C = m01E0.M5B37
      Assume l053C > 0, "pGuid"
      sub1DFA l053C
      Assume m01E0.M5B40 = 0, "l14"
      l053C = m01E0.M5B47
      Assume l053C > 0, "ppGuid"
      sub1E21 l053C
      Assume m01E0.M5B51 = 1, "l16"
      l053C = m01E0.M5B58
      Assume l053C > 0, "pHdr4"
      If l053C <> m01E0.M5B7C Then
        DebugPrintVar l053C = m01E0.M5B7C, "always same ref"
        sub1E07 l053C, m01E0.M5B51
      End If
      Assume m01E0.M5B61 = 0, "l18"
      l053C = m01E0.M5B68
      Assume l053C > 0, "ppGuid2"
      sub1E21 l053C
      l053C = m01E0.M5B73
      Assume l053C > 0, "#classes"
      l053C = m01E0.M5B7C
        sub1E07 l053C, m01E0.M5B73
      l053C = m01E0.M5B97
      Assume l053C > 0, "pMeths"
      l0530 = RetOffsetSym(l053C, m004A & ".methods", 0)
sub0C11 g_decl_num, l0530
m0018(g_decl_num).M589B = l0530
      If m01E0.M5B86 > 0 Then
        sub1EB7 l053C, m01E0.M5B86
      End If
      l053C = m01E0.M5BA1
      If l053C > 0 Then
        l0530 = RetOffsetSym(l053C, m004A & ".data3class", 0)
sub0C11 g_decl_num, l0530
      End If
    End If
    If m01A6.M5AD2 > 0 Then
      sub1DDC m01A6.M5AD2
    End If
    If m01A6.M598A > 0 Then
      l0530 = RetOffsetSym(m01A6.M598A, m004A & ".oGbl", 0)
g_decl_objs(g_decl_num).gbls_sorted = m01A6.M598A - m0012
    End If
    If m01A6.M5AFC > 0 Then
      l052E = RetOffsetSym(m01A6.M5B1A, m004A & ".Args", 0)
g_decl_objs(g_decl_num).M41F0 = l052E
      gReloc_Arr(l052E).M36B3_count = mc0584
    End If
    l0536 = l0536 + Len(m0172)
  Loop
  
  l052E = gv2EF8(2)
  For i = 1 To MForms.MFormsCount
    gReloc_Arr(l052E).M36B3_count = mc0586 + mc054A
    l0536 = fn202F(l052E)
    l0530 = RetOffsetSym(l0536, "Form#" & CStr(i), mc0582)
    Get g_hFileExe, fn1483(l0530), m0120
    
    l053A = m0120.M5A03 + 1
    sub0C11 l053A, l0530
    
    l0530 = RetOffsetSym(m0120.M5A18, "Form" & CStr(i) & ".RC", mc0588)
    
    gReloc_Arr(l0530).M36B3_count = mc0588
    g_decl_objs(l053A).M41D7 = l0530
    l052E = CreateSym(l0536 + &H50)
  Next
  
  Assume gFormsCount = MForms.MFormsCount, "forms done"
  gEXEdef2 = gEXEdef2 + gFormsCount
  If gIsVB5 < 0 Then
    sub1E75
  Else
    sub1E96
  End If
  
  l052E = gv2EF8(0)
  sub1E49 l052E, "Tab?"
  If gEXEdef_Arr5Count Then
    If gIsVB_Ver < 0 Then
      sub1EC7
    Else
      sub1EDC
    End If
    
    If gv097C Then
      Assume gv097C = g_decl_num, "prj match"
      sub0DAA
    End If
  Else
    MsgBox "no subs found!"
  End If
  sub0DB8
  If gEXEdef_Arr5Count Then sub0DC9
End Sub

Sub sub1EB7(p058C As Long, p058E As Integer)
Dim l0590 As Integer
Dim l0592 As Integer
Dim l0594 As Integer
Dim l0596 As Long
Dim l0598 As Long
Dim l059A As Integer
Dim l059C
  If gIsVB5 >= 0 Then Exit Sub
  l059A = fn1624(p058C)
  For l059C = 1 To p058E
    l0590 = gv161C(l059A).subRef_M3C6A
    l0596 = fn202F(l0590)
    Get g_hFileExe, fn15A9(l0596), l0598
    If (l0598 And &HFF) <> &HE9 Then GoTo L40E0E
    Get g_hFileExe, fn15A9(l0596) + 1, l0598
    l0598 = l0598 + l0596 + 5
    gReloc_Arr(l0590).M36B3_count = mc04EA
    l0594 = CreateSym(l0598)
    gReloc_Arr(l0590).subRef_M3C6A = -l0594
    l0592 = CreateSym(l0596 - 8)
    gReloc_Arr(l0592).M36B3_count = mc04EA
    gReloc_Arr(l0592).subRef_M3C6A = -l0594
    l059A = fn164F(l059A)
  Next
Exit Sub
L40E0E:
  MsgBox "not method/event!"
End Sub

Private Sub sub1EC7()
Dim l05A0
Dim l05A2 As Integer
Dim l05A4 As Integer
Dim l05A8 As Integer
Dim l05AA
Dim l05AC As Integer
Dim l05AE As Long
Dim l05B0 As String
Dim l05B2 As Integer
Dim l05B4 As Integer
Dim l05B6 As Integer
Dim l05B8() As Long
Dim l05BE() As Integer
Dim l05C4() As Long
Dim l05CA() As Long
' Const mc05DA = 32 ' &H20%
  For l05AA = 1 To g_decl_num
    l05A2 = m0018(l05AA).M589B
    l05AC = m0018(l05AA).M3F13
    If l05A2 <> 0 And l05AC > 0 Then
      ReDim l05B8(l05AC) As Long, l05BE(l05AC) As Integer
      ReDim l05C4(l05AC) As Long, l05CA(l05AC) As Long
      l05A2 = m0018(l05AA).M589B
      Seek g_hFileExe, fn1483(l05A2)
      For l05A0 = 1 To l05AC
        Get g_hFileExe, , l05B8(l05A0)
      Next
      l05B4 = gEXEdef_Arr5Count: l05B6 = 0
      For l05A0 = 1 To l05AC
        l05AE = l05B8(l05A0)
        If l05AE Then
          l05A2 = SomeAdrConv(l05AE)
          l05A4 = fn2047(l05A2)
          l05A8 = -gEXEdef_Arr4_imports(l05A4).TypeCode_M5F96
          gEXEdef_Subs(l05A8).ModuleIndex = l05AA
          l05BE(l05A0) = l05A8
          If l05A8 > l05B6 Then l05B6 = l05A8
          If l05A8 < l05B4 Then l05B4 = l05A8
        Else
          ShowMsgBox mc05DA
        End If
      Next
      l05B4 = l05B4 - 1
      Assume l05B6 - l05B4 = g_decl_objs(l05AA).M3F13, "method count"
      g_decl_objs(l05AA).M3F1B = l05B4
      l05B2 = l05B4
      l05A2 = m0018(l05AA).M58A6
      Seek g_hFileExe, fn1483(l05A2)
      For l05A0 = 1 To l05AC
        Get g_hFileExe, , l05C4(l05A0)
      Next
    l05A2 = m0018(l05AA).M58B2
    If l05A2 Then
      Seek g_hFileExe, fn1483(l05A2)
      For l05A0 = 1 To l05AC
        Get g_hFileExe, , l05CA(l05A0)
      Next
      For l05A0 = 1 To l05AC
        l05AE = l05CA(l05A0)
        If l05AE Then
          l05A2 = SomeAdrConv(l05AE)
        Else
          l05A2 = 0
        End If
        If l05A2 Then
          l05A2 = fn1624(l05AE)
          l05B0 = fn2023(gv161C(l05A2).subRef_M3C6A)
          If Left$(l05B0, 1) <> "(" Then l05B0 = "()"
        Else
          l05B0 = "(?)"
        End If
        l05AE = l05C4(l05A0)
        If l05AE Then
          l05B0 = "=" & ReadZeroString(g_hFileExe, fn15A9(l05AE)) & l05B0
        Else
        End If
        l05A8 = gEXEdef_Subs(l05BE(l05A0)).lOffsetSym_M3C6A
        l05AE = fn202F(l05A8)
        Assume l05AE <> 0, "sub found"
          l05B0 = fn2023(l05A8) & l05B0
          l05A4 = RetOffsetSym(l05AE, l05B0, 0)
      Next
    End If
    Else
      g_decl_objs(l05AA).M3F1B = l05B2
      If l05AC Then
        gEXEdef_Subs(l05B2 + 1).ModuleIndex = l05AA
      End If
    End If
    l05B2 = l05B2 + l05AC
    If l05B2 > gEXEdef_Arr5Count Then
      g_decl_objs(l05AA).M3F13 = gEXEdef_Arr5Count - g_decl_objs(l05AA).M3F1B
      l05B2 = gEXEdef_Arr5Count
    End If
  Next
End Sub

Private Sub sub1EDC()
Dim l05DC As Long
Dim l05DE
Dim l05E4 As Integer
Dim l05E6 As Integer
Dim l05E8
Dim l05EE As Integer
  If gIsVB_Ver < 0 Then
    Exit Sub
  End If
If gIsVB_Ver > 0 Then
  For l05E8 = 1 To gEXEdef1
    g_decl_objs(l05E8).M3F1B = 0
  Next
  For l05DC = 1 To gEXEdef_Arr5Count
    Select Case gEXEdef_Subs(l05DC).M36B3_count
    Case mc051E, mc051A
      l05E6 = gEXEdef_Subs(l05DC).lOffsetSym_M3C6A
      l05E4 = fn14AE(l05E6, 0, 0, True)
      Assume l05E4, "subhdr"
      If l05E4 Then
        Get g_hFileExe, fn1483(l05E4), m01A6
        Assume m01A6.M34FD = 1, "hdr3"
        l05E8 = m01A6.M5ABD + 1
        gEXEdef_Subs(l05DC).ModuleIndex = l05E8
        If g_decl_objs(l05E8).M3F1B = 0 Then
          g_decl_objs(l05E8).M3F1B = l05DC
        Else
          Assume l05DC > g_decl_objs(l05E8).M3F1B And l05DC <= g_decl_objs(l05E8).M3F1B + g_decl_objs(l05E8).M3F13, "sub order"
        End If
      End If
    End Select
  Next
  For l05E8 = 1 To gEXEdef1
    l05EE = g_decl_objs(l05E8).M3F1B
    If l05EE Then
      g_decl_objs(l05E8).M3F1B = l05EE - 1
      For l05DE = 1 To g_decl_objs(l05E8).M3F13
        If gEXEdef_Subs(l05DE + l05EE).ModuleIndex <> l05E8 Then
          DbgMileStone "dlls?"
        End If
      Next
    Else
      MsgBox g_decl_objs(l05E8).Name_M411A, , "no subs in"
    End If
  Next
Else
  l05DC = 1
  For l05DE = 1 To g_decl_num
    If g_decl_objs(l05DE).M3F13 Then
      l05E6 = gEXEdef_Subs(l05DC).lOffsetSym_M3C6A
      l05E4 = fn14AE(l05E6, 0, 0, True)
      Assume l05E4, "subhdr"
      If l05E4 Then
        Get g_hFileExe, fn1483(l05E4), m01A6
        Assume m01A6.M34FD = 1, "hdr3"
        l05E8 = m01A6.M5ABD + 1
        g_decl_objs(l05E8).M3F1B = l05DC - 1
        l05DC = l05DC + g_decl_objs(l05E8).M3F13
      Else
        MsgBox fn2023(l05E6), , "wrong sub"
        Exit For
      End If
    End If
  Next
End If
  Assume l05DC = gEXEdef_Arr5Count + 1, "sub count"
End Sub

Sub sub1EF0()
Dim l05EE As Integer
  DebugPrintVar False, "unused"
End Sub
