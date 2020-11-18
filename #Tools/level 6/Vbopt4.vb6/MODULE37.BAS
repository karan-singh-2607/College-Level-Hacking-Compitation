Attribute VB_Name = "MODULE37"
' Module37
Option Explicit
Dim m0014 As Integer
Const mc0050 = 224 ' &HE0%
Dim m0060(7) As Integer
Dim m0076(7) As Integer
Const mc00BE = 3 ' &H3%
Const mc00C0 = 2 ' &H2%
Const mc00C2 = 5 ' &H5%
Const mc00E2 = 9 ' &H9%
Const mc00E4 = 4 ' &H4%
Const mc00E6 = 5 ' &H5%
Const mc00E8 = 8 ' &H8%
Const mc00F2 = 15 ' &HF%
Const mc00F6 = 25 ' &H19%
Const mc00F8 = 26 ' &H1A%
Const mc00FA = 7 ' &H7%
Const mc012A = 24 ' &H18%
Const mc012C = 27 ' &H1B%
Const mc012E = 28 ' &H1C%
Const mc0130 = 29 ' &H1D%
Const mc0140 = 32 ' &H20%
Const mc0142 = 37 ' &H25%
Const mc0144 = 36 ' &H24%
Const mc0146 = 21 ' &H15%
Const mc0158 = 30 ' &H1E%

Sub DisasmLine()
Dim VAOffsetInFile As Long
Dim codebyte As Integer
  m0014 = True
  
  DisAsmCmd
  
  ' Add Offset
  gDisOutputTxt = Hex$(gVACmdStart + gVANegOffsetCurMod) & " " & gDisOutputTxt
  
  ' Add Space
  If Len(gDisOutputTxt) < 50 Then
    gDisOutputTxt = gDisOutputTxt & Space$(50 - Len(gDisOutputTxt))
  End If
  
  ' Add Hex CodeBytes
  VAOffsetInFile = gVACmdStart
  While VAOffsetInFile < gVACmdEnd
    
    Get g_hFileExe, VAOffsetInFile + gVAToRaw, gSingleChar
    
    codebyte = Asc(gSingleChar): VAOffsetInFile = VAOffsetInFile + 1
    gDisOutputTxt = gDisOutputTxt & ConCatHex8((codebyte))
    
  Wend
  
  WriteOutputTxt
  
End Sub

Sub sub22BE(p1&, p00B4 As String)
' called (p1&, ByVal p2$)
  MsgBox p00B4
End Sub

Private Sub sub22C9(p00B6 As Integer, p00B8 As Integer)
' Const mc00BE = 3 ' &H3%
' Const mc00C0 = 2 ' &H2%
' Const mc00C2 = 5 ' &H5%
  If gVACmdEnd = gVAModuleEnd Then
    Select Case gv014E
    Case mc00BE
      If p00B6 <> 4 Then
        ErrBeep
        sub2232
      End If
    Case mc00C0
      If p00B6 <> 2 Then ErrBeep
    Case mc00C2
      If p00B6 <> 4 And p00B6 <> 2 Then ErrBeep
    End Select
    p00B8 = gv0150
    sub019B
  End If
  gVACmdEnd = gVACmdEnd + p00B6
End Sub

Function fn22D8(p00C8%, p00CA%) As Long
Dim l00CC As Integer
Dim l00CE As Long
  If gVACmdEnd >= gVAModuleEnd Then
    sub22C9 p00C8, p00CA
  Else
    On p00C8 GoTo L4A944, L4A9CE, L4A9BE, L4AA10, L4AA50, L4AA94
    ErrBeep
  End If
Exit Function

L4A944:
  Get g_hFileExe, gVACmdEnd + gVAToRaw, gSingleChar
  l00CC = Asc(gSingleChar): gVACmdEnd = gVACmdEnd + 1
  If l00CC And &H80 Then
    fn22D8 = l00CC Or &HFF00
  Else
    fn22D8 = l00CC
  End If
Exit Function

L4A9BE:
  ErrBeep

L4A9CE:
  Get g_hFileExe, gVAToRaw + gVACmdEnd, l00CC: gVACmdEnd = gVACmdEnd + 2
  fn22D8 = l00CC
Exit Function

L4AA10:
  Get g_hFileExe, gVAToRaw + gVACmdEnd, l00CE
  fn22D8 = l00CE: gVACmdEnd = gVACmdEnd + 4
Exit Function

L4AA50:
  ErrBeep
  GoSub L4A9CE
  If gVACmdEnd <> gVAModuleEnd Then ErrBeep
  sub22C9 2, p00CA
Exit Function

L4AA94:
  MsgBox "unexpected ptr32"
  GoSub L4AA10
  If gVACmdEnd <> gVAModuleEnd Then ErrBeep
  sub22C9 2, p00CA
Exit Function
End Function

Function fn22E6(p00D2 As Integer, p00D4%) As Long
Dim l00D6 As Integer
Dim l00D8 As Long
  If gVACmdEnd = gVAModuleEnd Then
    If gv0150 Then
      If p00D2 = 5 Then p00D2 = 4
      sub22C9 p00D2, p00D4
      Exit Function
    End If
  End If
  On p00D2 GoTo L4AB6A, L4ABCE, L4ABB8, L4AC18, L4AC58, L4ACD2
  ErrBeep
Exit Function

L4AB6A:
  Get g_hFileExe, gVACmdEnd + gVAToRaw, gSingleChar
  l00D6 = Asc(gSingleChar): gVACmdEnd = gVACmdEnd + 1
  fn22E6 = l00D6
Exit Function

L4ABB8:
  ErrBeep
  p00D2 = 2
L4ABCE:
  Get g_hFileExe, gVAToRaw + gVACmdEnd, l00D6: gVACmdEnd = gVACmdEnd + 2
  fn22E6 = CLng(l00D6) And &HFFFF&
Exit Function

L4AC18:
  Get g_hFileExe, gVAToRaw + gVACmdEnd, l00D8
  fn22E6 = l00D8: gVACmdEnd = gVACmdEnd + 4
Exit Function

L4AC58:
  Get g_hFileExe, gVAToRaw + gVACmdEnd, l00D6: gVACmdEnd = gVACmdEnd + 2
  fn22E6 = CLng(l00D6) And &HFFFF&
  If gVACmdEnd <> gVAModuleEnd Then ErrBeep
  sub22C9 2, p00D4
  p00D2 = 4
Exit Function

L4ACD2:
  MsgBox "unexpected ptr32"
  Get g_hFileExe, gVAToRaw + gVACmdEnd, l00D8
  fn22E6 = l00D8: gVACmdEnd = gVACmdEnd + 4
  If gVACmdEnd <> gVAModuleEnd Then ErrBeep
  sub22C9 2, p00D4
Exit Function
End Function

Sub sub2232()
  sub22BE gVACmdStart, "illegal op"
  gCmdByteStruc.M642E = gCmdByteStruc.M642E Or gc25A2
End Sub

Sub sub22F6()
  sub2305
  sub2240
  sub2314
  gv25D8 = &H1FF
  ReDim gCmdByteArr(gv25D8)
sub22AD &H0, "ADD.b, Eb, Gb", gc256E, 0, 0
sub22AD &H1, "ADD.?, Ev, Gv", gc256E, 0, gc2596
sub22AD &H2, "ADD.b, Gb, Eb", gc256E, 0, 0
sub22AD &H3, "ADD.?, Gv, Ev", gc256E, 0, gc2596
sub22AD &H4, "ADD.b, AL, Ib", gc256A, 0, 0
sub22AD &H5, "ADD.?, eAX, Iv", gc256A, 0, gc2596
sub22AD &H6, "PUSH.s, SP-,ES", 0, gc258A, 0
sub22AD &H7, "POP.s, ES,SP+", 0, gc258C, 0
sub22AD &H8, "OR.b, Eb, Gb", gc256E, 0, 0
sub22AD &H9, "OR.?, Ev, Gv", gc256E, 0, gc2596
sub22AD &HA, "OR.b, Gb, Eb", gc256E, 0, 0
sub22AD &HB, "OR.?, Gv, Ev", gc256E, 0, gc2596
sub22AD &HC, "OR.b, AL, Ib", 0, 0, 0
sub22AD &HD, "OR.?, eAX, Iv", 0, 0, gc2596
sub22AD &HE, "PUSH.s, SP-,CS", 0, gc258A, 0
sub22AD &HF, "2byte*", gc256C, 0, 0
sub22AD &H10, "ADC.b, Eb, Gb", gc256E, 0, 0
sub22AD &H11, "ADC.?, Ev, Gv", gc256E, 0, gc2596
sub22AD &H12, "ADC.b, Gb, Eb", gc256E, 0, 0
sub22AD &H13, "ADC.?, Gv, Ev", gc256E, 0, gc2596
sub22AD &H14, "ADC.b, AL, Ib", 0, 0, 0
sub22AD &H15, "ADC.?, eAX, Iv", 0, 0, gc2596
sub22AD &H16, "PUSH.s, SP-,SS", 0, gc258A, 0
sub22AD &H17, "POP.s, SS,SP+", 0, gc258C, 0
sub22AD &H18, "SBB.b, Eb, Gb", gc256E, 0, 0
sub22AD &H19, "SBB.?, Ev, Gv", gc256E, 0, gc2596
sub22AD &H1A, "SBB.b, Gb, Eb", gc256E, 0, 0
sub22AD &H1B, "SBB.?, Gv, Ev", gc256E, 0, gc2596
sub22AD &H1C, "SBB.b, AL, Ib", 0, 0, 0
sub22AD &H1D, "SBB.?, eAX, Iv", 0, 0, gc2596
sub22AD &H1E, "PUSH.s, SP-,DS", 0, gc258A, 0
sub22AD &H1F, "POP.s, DS,SP+", 0, gc258C, 0
sub22AD &H20, "AND.b, Eb, Gb", gc256E, 0, 0
sub22AD &H21, "AND.?, Ev, Gv", gc256E, 0, gc2596
sub22AD &H22, "AND.b, Gb, Eb", gc256E, 0, 0
sub22AD &H23, "AND.?, Gv, Ev", gc256E, 0, gc2596
sub22AD &H24, "AND.b, AL, Ib", 0, 0, 0
sub22AD &H25, "AND.?, eAX, Iv", 0, 0, gc2596
sub22AD &H26, "ES:", gc257C, 0, 0
sub22AD &H27, "DAA,Ib", 0, 0, 0
sub22AD &H28, "SUB.b, Eb, Gb", gc256E, 0, 0
sub22AD &H29, "SUB.?, Ev, Gv", gc256E, 0, gc2596
sub22AD &H2A, "SUB.b, Gb, Eb", gc256E, 0, 0
sub22AD &H2B, "SUB.?, Gv, Ev", gc256E, 0, gc2596
sub22AD &H2C, "SUB.b, AL, Ib", 0, 0, 0
sub22AD &H2D, "SUB.?, eAX, Iv", 0, 0, gc2596
sub22AD &H2E, "CS:", gc257C, 0, 0
sub22AD &H2F, "DAS", 0, 0, 0
sub22AD &H30, "XOR.b, Eb, Gb", gc256E, 0, 0
sub22AD &H31, "XOR.?, Ev, Gv", gc256E, 0, gc2596
sub22AD &H32, "XOR.b, Gb, Eb", gc256E, 0, 0
sub22AD &H33, "XOR.?, Gv, Ev", gc256E, 0, gc2596
sub22AD &H34, "XOR.b, AL, Ib", 0, 0, 0
sub22AD &H35, "XOR.?, eAX, Iv", 0, 0, gc2596
sub22AD &H36, "SS:", gc257C, 0, 0
sub22AD &H37, "AAA", 0, 0, 0
sub22AD &H38, "CMP.b, Eb, Gb", gc256E, 0, 0
sub22AD &H39, "CMP.?, Ev, Gv", gc256E, 0, gc2596
sub22AD &H3A, "CMP.b, Gb, Eb", gc256E, 0, 0
sub22AD &H3B, "CMP.?, Gv, Ev", gc256E, 0, gc2596
sub22AD &H3C, "CMP.b, AL, Ib", 0, 0, 0
sub22AD &H3D, "CMP.?, eAX, Iv", 0, 0, gc2596
sub22AD &H3E, "DS:", gc257C, 0, 0
sub22AD &H3F, "AAS", 0, 0, 0
sub22AD &H40, "INC.?, eAX", 0, 0, gc2596
sub22AD &H41, "INC.?, eCX", 0, 0, gc2596
sub22AD &H42, "INC.?, eDX", 0, 0, gc2596
sub22AD &H43, "INC.?, eBX", 0, 0, gc2596
sub22AD &H44, "INC.?, eSP", 0, gc2588, gc2596
sub22AD &H45, "INC.?, eBP", 0, gc2592, gc2596
sub22AD &H46, "INC.?, eSI", 0, 0, gc2596
sub22AD &H47, "INC.?, eDI", 0, 0, gc2596
sub22AD &H48, "DEC.?, eAX", 0, 0, gc2596
sub22AD &H49, "DEC.?, eCX", 0, 0, gc2596
sub22AD &H4A, "DEC.?, eDX", 0, 0, gc2596
sub22AD &H4B, "DEC.?, eBX", 0, 0, gc2596
sub22AD &H4C, "DEC.?, eSP", 0, gc2586, gc2596
sub22AD &H4D, "DEC.?, eBP", 0, gc2592, gc2596
sub22AD &H4E, "DEC.?, eSI", 0, 0, gc2596
sub22AD &H4F, "DEC.?, eDI", 0, 0, gc2596
sub22AD &H50, "PUSH.?, SP-,eAX", 0, gc2586, gc2596
sub22AD &H51, "PUSH.?, SP-,eCX", 0, gc2586, gc2596
sub22AD &H52, "PUSH.?, SP-,eDX", 0, gc2586, gc2596
sub22AD &H53, "PUSH.?, SP-,eBX", 0, gc2586, gc2596
sub22AD &H54, "PUSH.?, SP-,eSP", 0, gc2586, gc2596
sub22AD &H55, "PUSH.?, SP-,eBP", 0, gc2586, gc2596
sub22AD &H56, "PUSH.?, SP-,eSI", 0, gc2586, gc2596
sub22AD &H57, "PUSH.?, SP-,eDI", 0, gc2586, gc2596
sub22AD &H58, "POP.?, eAX,SP+", 0, gc2588, gc2596
sub22AD &H59, "POP.?, eCX,SP+", 0, gc2588, gc2596
sub22AD &H5A, "POP.?, eDX,SP+", 0, gc2588, gc2596
sub22AD &H5B, "POP.?, eBX,SP+", 0, gc2588, gc2596
sub22AD &H5C, "POP.?, eSP,SP+", 0, gc2588, gc2596
sub22AD &H5D, "POP.?, eBP,SP+", 0, gc2588, gc2596
sub22AD &H5E, "POP.?, eSI,SP+", 0, gc2588, gc2596
sub22AD &H5F, "POP.?, eDI,SP+", 0, gc2588, gc2596
sub22AD &H60, "PUSHA.?", 0, gc2586, gc2596
sub22AD &H61, "POPA.?", 0, gc2588, gc2596
sub22AD &H62, "BOUND.?, Gv, Ma", gc256E, 0, gc2596
sub22AD &H63, "ARPL.w, Ew, Rw", gc256E, 0, 0
sub22AD &H64, "FS:", gc257C, 0, 0
sub22AD &H65, "GS:", gc257C, 0, 0
sub22AD &H66, "op*", gc257A, 0, 0
sub22AD &H67, "ad*", gc2578, 0, 0
sub22AD &H68, "PUSH.?, SP-,Iv", 0, gc2586, gc2596
sub22AD &H69, "IMUL.?, Gv, Ev, Iv", gc256E, 0, gc2596
sub22AD &H6A, "PUSH.b, SP-,Ib", 0, gc2586, 0
sub22AD &H6B, "IMUL.?, Gv, Ev, Ib", gc256E, 0, gc2596
sub22AD &H6C, "INS.b", 0, 0, gc259E
sub22AD &H6D, "INS.?", 0, 0, gc259E Or gc2596
sub22AD &H6E, "OUTS.b", 0, 0, gc259E
sub22AD &H6F, "OUTS.?", 0, 0, gc259E Or gc2596
sub22AD &H70, "JO, Jb", 0, gc257E, 0
sub22AD &H71, "JNO, Jb", 0, gc257E, 0
sub22AD &H72, "JB, Jb", 0, gc257E, 0
sub22AD &H73, "JNB, Jb", 0, gc257E, 0
sub22AD &H74, "JZ, Jb", 0, gc257E, 0
sub22AD &H75, "JNZ, Jb", 0, gc257E, 0
sub22AD &H76, "JBE, Jb", 0, gc257E, 0
sub22AD &H77, "JNBE, Jb", 0, gc257E, 0
sub22AD &H78, "JS, Jb", 0, gc257E, 0
sub22AD &H79, "JNS, Jb", 0, gc257E, 0
sub22AD &H7A, "JP, Jb", 0, gc257E, 0
sub22AD &H7B, "JNP, Jb", 0, gc257E, 0
sub22AD &H7C, "JL, Jb", 0, gc257E, 0
sub22AD &H7D, "JNL, Jb", 0, gc257E, 0
sub22AD &H7E, "JLE, Jb", 0, gc257E, 0
sub22AD &H7F, "JNLE, Jb", 0, gc257E, 0
sub22AD &H80, "ADD.b,Eb,Ib/OR.b,Eb,Ib/ADC.b,Eb,Ib/SBB.b,Eb,Ib/AND.b,Eb,Ib/SUB.b,Eb,Ib/XOR.b,Eb,Ib/CMP.b,Eb,Ib/", gc256E, 0, 0
sub22AD &H81, "ADD.?,Ev,Iv/OR.?,Ev,Iv/ADC.?,Ev,Iv/SBB.?,Ev,Iv/AND.?,Ev,Iv/SUB.?,Ev,Iv/XOR.?,Ev,Iv/CMP.?,Ev,Iv/", gc256E, 0, gc2596
sub22AD &H83, "ADD.?,Ev,Ib/OR.?,Ev,Ib/ADC.?,Ev,Ib/SBB.?,Ev,Ib/AND.?,Ev,Ib/SUB.?,Ev,Ib/XOR.?,Ev,Ib/CMP.?,Ev,Ib/", gc256E, 0, gc2596
sub22AD &H84, "TEST.b, Eb, Gb", gc256E, 0, 0
sub22AD &H85, "TEST.?, Ev, Gv", gc256E, 0, gc2596
sub22AD &H86, "XCHG.b, Eb, Gb", gc256E, 0, 0
sub22AD &H87, "XCHG.?, Ev, Gv", gc256E, 0, gc2596
sub22AD &H88, "MOV.b, Eb, Gb", gc256E, 0, 0
sub22AD &H89, "MOV.?, Ev, Gv", gc256E, 0, gc2596
sub22AD &H8A, "MOV.b, Gb, Eb", gc256E, 0, 0
sub22AD &H8B, "MOV.?, Gv, Ev", gc256E, 0, gc2596
sub22AD &H8C, "MOV.s, Ew, Sw", gc256E, gc2590, 0
sub22AD &H8D, "LEA.?, Gv, Mv", gc256E, 0, gc2596 Or gc2594
sub22AD &H8E, "MOV.s, Sw, Ew", gc256E, gc258E, 0
sub22AD &H8F, "POP.?, Ev,SP+", gc256E, gc2588, gc2596
sub22AD &H90, "NOP", 0, 0, 0
sub22AD &H91, "XCHG.?, eAX, eCX", 0, 0, gc2596
sub22AD &H92, "XCHG.?, eAX, eDX", 0, 0, gc2596
sub22AD &H93, "XCHG.?, eAX, eBX", 0, 0, gc2596
sub22AD &H94, "XCHG.?, eAX, eSP", 0, gc2592, gc2596 Or gc25A0
sub22AD &H95, "XCHG.?, eAX, eBP", 0, gc2592, gc2596 Or gc25A0
sub22AD &H96, "XCHG.?, eAX, eSI", 0, 0, gc2596
sub22AD &H97, "XCHG.?, eAX, eDI", 0, 0, gc2596
sub22AD &H98, "CBW\CWDE", 0, 0, 0
sub22AD &H99, "CWD\CDQ", 0, 0, 0
sub22AD &H9A, "CALL.F, Ap", 0, gc2582, 0
sub22AD &H9B, "WAIT", 0, 0, 0
sub22AD &H9C, "PUSH.?,SP-,FLG", 0, gc2586, gc2596
sub22AD &H9D, "POP.?, FLG,SP+", 0, gc2588, gc2596
sub22AD &H9E, "SAHF.b,FLG,AH", 0, 0, 0
sub22AD &H9F, "LAHF.b,AH,FLG", 0, 0, 0
sub22AD &HA0, "MOV.b, AL, Ob", 0, 0, 0
sub22AD &HA1, "MOV.?, eAX, Ov", 0, 0, gc2596
sub22AD &HA2, "MOV.b, Ob, AL", 0, 0, 0
sub22AD &HA3, "MOV.?, Ov, eAX", 0, 0, gc2596
sub22AD &HA4, "MOVS.b,DST,SRC", 0, 0, 0
sub22AD &HA5, "MOVS.?,DST,SRC", 0, 0, gc2596
sub22AD &HA6, "CMPS.b,DST,SRC", 0, 0, 0
sub22AD &HA7, "CMPS.?,DST,SRC", 0, 0, gc2596
sub22AD &HA8, "TEST.b, AL, Ib", 0, 0, 0
sub22AD &HA9, "TEST.?, eAX, Iv", 0, 0, gc2596
sub22AD &HAA, "STOS.b,DST,AL", 0, 0, 0
sub22AD &HAB, "STOS.?,DST,eAX", 0, 0, gc2596
sub22AD &HAC, "LODS.b,AL,SRC", 0, 0, 0
sub22AD &HAD, "LODS.?,eAX,SRC", 0, 0, gc2596
sub22AD &HAE, "SCAS.b,DST,AL", 0, 0, 0
sub22AD &HAF, "SCAS.?,DST,eAX", 0, 0, gc2596
sub22AD &HB0, "MOV.b, AL, Ib", 0, 0, 0
sub22AD &HB1, "MOV.b, CL, Ib", 0, 0, 0
sub22AD &HB2, "MOV.b, DL, Ib", 0, 0, 0
sub22AD &HB3, "MOV.b, BL, Ib", 0, 0, 0
sub22AD &HB4, "MOV.b, AH, Ib", 0, 0, 0
sub22AD &HB5, "MOV.b, CH, Ib", 0, 0, 0
sub22AD &HB6, "MOV.b, DH, Ib", 0, 0, 0
sub22AD &HB7, "MOV.b, BH, Ib", 0, 0, 0
sub22AD &HB8, "MOV.?, eAX, Iv", 0, 0, gc2596
sub22AD &HB9, "MOV.?, eCX, Iv", 0, 0, gc2596
sub22AD &HBA, "MOV.?, eDX, Iv", 0, 0, gc2596
sub22AD &HBB, "MOV.?, eBX, Iv", 0, 0, gc2596
sub22AD &HBC, "MOV.?, eSP, Iv", 0, gc2592, gc2596 Or gc25A0
sub22AD &HBD, "MOV.?, eBP, Iv", 0, gc2592, gc2596
sub22AD &HBE, "MOV.?, eSI, Iv", 0, 0, gc2596
sub22AD &HBF, "MOV.?, eDI, Iv", 0, 0, gc2596
sub22AD &HC0, "ROL.b,Eb,Ib/ROR.b,Eb,Ib/RCL.b,Eb,Ib/RCR.b,Eb,Ib/SHL.b,Eb,Ib/SHR.b,Eb,Ib//SAR.b,Eb,Ib/", gc2570, 0, 0
sub22AD &HC1, "ROL.b,Eb,Ib/ROR.b,Eb,Ib/RCL.b,Eb,Ib/RCR.b,Eb,Ib/SHL.b,Eb,Ib/SHR.b,Eb,Ib//SAR.b,Eb,Ib/", gc2570, 0, 0
sub22AD &HC2, "RET.N,Iw", 0, gc2584, 0
sub22AD &HC3, "RET.N", 0, gc2584, 0
sub22AD &HC4, "LES.?, ES,Gv, Mp", gc256E, gc258E, gc2596 Or gc2594
sub22AD &HC5, "LDS.?, DS,Gv, Mp", gc256E, gc258E, gc2596 Or gc2594
sub22AD &HC6, "MOV.b, Eb, Ib", gc256E, 0, 0
sub22AD &HC7, "MOV.?, Ev, Iv", gc256E, 0, gc2596
sub22AD &HC8, "ENTER, Iw, Ib", 0, gc2592, 0
sub22AD &HC9, "LEAVE", 0, gc2592, 0
sub22AD &HCA, "RET.F, Iw", 0, gc2584, 0
sub22AD &HCB, "RET.F", 0, gc2584, 0
sub22AD &HCC, "INT3", 0, 0, 0
sub22AD &HCD, "INT, Ib", 0, 0, 0
sub22AD &HCE, "INTO", 0, 0, 0
sub22AD &HCF, "IRET", 0, 0, gc259E Or gc2584
sub22AD &HD0, "ROL.b,Eb/ROR.b,Eb/RCL.b,Eb/RCR.b,Eb/SHL.b,Eb/SHR.b,Eb//SAR.b,Eb/", gc2570, 0, 0
sub22AD &HD1, "ROL.?,Ev/ROR.?,Ev/RCL.?,Ev/RCR.?,Ev/SHL.?,Ev/SHR.?,Ev//SAR.?,Ev/", gc2570, 0, gc2596
sub22AD &HD2, "ROL.b,Eb,CL/ROR.b,Eb,CL/RCL.b,Eb,CL/RCR.b,Eb,CL/SHL.b,Eb,CL/SHR.b,Eb,CL//SAR.b,Eb,CL/", gc2570, 0, 0
sub22AD &HD3, "ROL.?,Ev,CL/ROR.?,Ev,CL/RCL.?,Ev,CL/RCR.?,Ev,CL/SHL.?,Ev,CL/SHR.?,Ev,CL//SAR.?,Ev,CL/", gc2570, 0, gc2596
sub22AD &HD4, "AAM.b, Ib", 0, 0, 0
sub22AD &HD5, "AAD.b, Ib", 0, 0, 0
sub22AD &HD6, "SALC", 0, 0, gc259E
sub22AD &HD7, "XLATB", 0, 0, 0
sub22AD &HE0, "LOOPNE,Jb", 0, gc257E, 0
sub22AD &HE1, "LOOPE,Jb", 0, gc257E, 0
sub22AD &HE2, "LOOP,Jb", 0, gc257E, 0
sub22AD &HE3, "JZ.?,eCX,Jb", 0, gc257E, gc2596
sub22AD &HE4, "IN.b,AL,Ib", 0, 0, gc259E
sub22AD &HE5, "IN.?,eAX,Ib", 0, 0, gc2596 Or gc259E
sub22AD &HE6, "OUT.b,Ib, AL", 0, 0, gc259E
sub22AD &HE7, "OUT.?, Ib, eAX", 0, 0, gc2596
sub22AD &HE8, "CALL.N, Jv", 0, gc2582, 0
sub22AD &HE9, "JMP.N, Jv", 0, gc2580, gc2596
sub22AD &HEA, "JMP.F, Ap", 0, gc2580, 0
sub22AD &HEB, "JMP.N, Jb", 0, gc2580, 0
sub22AD &HEC, "IN.b,AL,DX", 0, 0, gc259E
sub22AD &HED, "IN.?,eAX,DX", 0, 0, gc259E Or gc2596
sub22AD &HEE, "OUT.b,DX,AL", 0, 0, gc259E
sub22AD &HEF, "OUT.?,DX,eAX", 0, 0, gc259E Or gc2596
sub22AD &HF0, "LOCK:", gc2576, 0, 0
sub22AD &HF2, "REPNE:", gc2576, 0, 0
sub22AD &HF3, "REPE:", gc2576, 0, 0
sub22AD &HF4, "HLT", 0, 0, gc259E
sub22AD &HF5, "CMC", 0, 0, 0
sub22AD &HF6, "TEST.b,Eb,Ib//NOT.b,Eb/NEG.b,Eb/MUL.b,AX,AL,Eb/IMUL.b,AX,AL,Eb/DIV.b,AH,AL,Eb/IDIV.b,AH,AL,Eb/", gc2570, 0, 0
sub22AD &HF7, "TEST.?,Ev,Iv//NOT.?,Ev/NEG.?,Ev/MUL.?,DX,eAX,Ev/IMUL.?,DX,eAX,Ev/DIV.?,eDX,eAX,Ev/IDIV.?,eDX,eAX,Ev/", gc2570, 0, gc2596
sub22AD &HF8, "CLC", 0, 0, 0
sub22AD &HF9, "STC", 0, 0, 0
sub22AD &HFA, "CLI", 0, 0, gc25A0
sub22AD &HFB, "STI", 0, 0, gc25A0
sub22AD &HFC, "CLD", 0, 0, 0
sub22AD &HFD, "STD", 0, 0, 0
sub22AD &HFE, "INC.b,Eb/DEC.b,Eb///////", 0, 0, 0
sub22AD &HFF, "INC.?,Ev/DEC.?,Ev/CALL.N,Ev/CALL.F,Ep/JMP.N,Ev/JMP.F,Ep/PUSH.?,Ev/", 0, 0, 0
sub22AD &H100, "SLDT.w,Ew/STR.w,Ew/LDT.w,Ew/LTR.w,Ew/VERR.w,Ew/VERW.w,Ew///", 0, 0, gc259E
sub22AD &H101, "SQDT.w,Ms/SIDT.w,Ms/LQDT.6,Ms/LIDT.6,Ms/SMSW.w,Ew//LMSW.w,Ew/INVLPG/", 0, 0, gc259E
sub22AD &H102, "LAR.w,Gv,Ew", gc256E, 0, gc2596
sub22AD &H103, "LSL.w,Gv,Ew", gc256E, 0, gc2596
sub22AD &H106, "CLTS", 0, 0, gc259E
sub22AD &H108, "INVD", 0, 0, gc259E
sub22AD &H109, "WBINVD", 0, 0, gc259E
sub22AD &H120, "MOV.l, Cl,Rl", gc256E, 0, gc259E
sub22AD &H121, "MOV.l, Dl,Rl", gc256E, 0, gc259E
sub22AD &H122, "MOV.l, Rl,Cl", gc256E, 0, gc259E
sub22AD &H123, "MOV.l, Rl,Dl", gc256E, 0, gc259E
sub22AD &H124, "MOV.l, Tl,Rl", gc256E, 0, gc259E
sub22AD &H126, "MOV.l, Rl,Tl", gc256E, 0, gc259E
sub22AD &H140, "CMOVO.?,Gv,Ev", gc256E, 0, gc2596
sub22AD &H141, "CMOVNO.?,Gv,Ev", gc256E, 0, gc2596
sub22AD &H142, "CMOVB.?,Gv,Ev", gc256E, 0, gc2596
sub22AD &H143, "CMOVAE.?,Gv,Ev", gc256E, 0, gc2596
sub22AD &H144, "CMOVE.?,Gv,Ev", gc256E, 0, gc2596
sub22AD &H145, "CMOVNE.?,Gv,Ev", gc256E, 0, gc2596
sub22AD &H146, "CMOVBE.?,Gv,Ev", gc256E, 0, gc2596
sub22AD &H147, "CMOVA.?,Gv,Ev", gc256E, 0, gc2596
sub22AD &H148, "CMOVS.?,Gv,Ev", gc256E, 0, gc2596
sub22AD &H149, "CMOVNS.?,Gv,Ev", gc256E, 0, gc2596
sub22AD &H14A, "CMOVP.?,Gv,Ev", gc256E, 0, gc2596
sub22AD &H14B, "CMOVNP.?,Gv,Ev", gc256E, 0, gc2596
sub22AD &H14C, "CMOVL.?,Gv,Ev", gc256E, 0, gc2596
sub22AD &H14D, "CMOVGE.?,Gv,Ev", gc256E, 0, gc2596
sub22AD &H14E, "CMOVLE.?,Gv,Ev", gc256E, 0, gc2596
sub22AD &H14F, "CMOVG.?,Gv,Ev", gc256E, 0, gc2596
sub22AD &H180, "JO,Jv", 0, gc257E, gc2596
sub22AD &H181, "JNO,Jv", 0, gc257E, gc2596
sub22AD &H182, "JB,Jv", 0, gc257E, gc2596
sub22AD &H183, "JAE,Jv", 0, gc257E, gc2596
sub22AD &H184, "JE,Jv", 0, gc257E, gc2596
sub22AD &H185, "JNE,Jv", 0, gc257E, gc2596
sub22AD &H186, "JBE,Jv", 0, gc257E, gc2596
sub22AD &H187, "JA,Jv", 0, gc257E, gc2596
sub22AD &H188, "JS,Jv", 0, gc257E, gc2596
sub22AD &H189, "JNS,Jv", 0, gc257E, gc2596
sub22AD &H18A, "JP,Jv", 0, gc257E, gc2596
sub22AD &H18C, "JL,Jv", 0, gc257E, gc2596
sub22AD &H18D, "JGE,Jv", 0, gc257E, gc2596
sub22AD &H18E, "JLE,Jv", 0, gc257E, gc2596
sub22AD &H18F, "JG,Jv", 0, gc257E, gc2596
sub22AD &H190, "SETO.b,Eb", gc256E, 0, 0
sub22AD &H191, "SETNO.b,Eb", gc256E, 0, 0
sub22AD &H192, "SETB.b,Eb", gc256E, 0, 0
sub22AD &H193, "SETNB.b,Eb", gc256E, 0, 0
sub22AD &H194, "SETZ.b,Eb", gc256E, 0, 0
sub22AD &H195, "SETNZ.b,Eb", gc256E, 0, 0
sub22AD &H196, "SETBE.b,Eb", gc256E, 0, 0
sub22AD &H197, "SETNBE.b,Eb", gc256E, 0, 0
sub22AD &H198, "SETS.b,Eb", gc256E, 0, 0
sub22AD &H199, "SETNS.b,Eb", gc256E, 0, 0
sub22AD &H19A, "SETP.b,Eb", gc256E, 0, 0
sub22AD &H19B, "SETNP.b,Eb", gc256E, 0, 0
sub22AD &H19C, "SETL.b,Eb", gc256E, 0, 0
sub22AD &H19D, "SETNL.b,Eb", gc256E, 0, 0
sub22AD &H19E, "SETLE.b,Eb", gc256E, 0, 0
sub22AD &H19F, "SETNLE.b,Eb", gc256E, 0, 0
sub22AD &H1A0, "PUSH.s,SP-,FS", 0, gc258A, 0
sub22AD &H1A1, "POP.s,FS,SP+", 0, gc258C, 0
sub22AD &H1A2, "CPUID", 0, 0, gc259E
sub22AD &H1A3, "BT.?,Ev,Gv", gc256E, 0, gc2596
sub22AD &H1A4, "SHLD.b,Ev,Gv,Ib", gc256E, 0, gc2596
sub22AD &H1A5, "SHLD.?,Ev,Gv,CL", gc256E, 0, gc2596
sub22AD &H1A6, "CMPXCHG.b,Eb,Gb", gc256E, 0, 0
sub22AD &H1A7, "CMPXCHG.?,Ev,Gv", gc256E, 0, gc2596
sub22AD &H1A8, "PUSH.s,SP-,GS", 0, gc258A, 0
sub22AD &H1A9, "POP.s,GS,SP+", 0, gc258C, 0
sub22AD &H1AB, "BTS.?,Gv,Ev", gc256E, 0, gc2596
sub22AD &H1AC, "SHRD.?, Ev,Gv,Ib", gc256E, 0, gc2596
sub22AD &H1AD, "SHRD.?, Ev,Gv,CL", gc256E, 0, gc2596
sub22AD &H1AF, "IMUL.?,Gv,Ev", gc256E, 0, gc2596
sub22AD &H1B2, "LSS.?,SS,Gv,Mp", gc256E, gc258E, gc2596 Or gc2594
sub22AD &H1B3, "BTR.?,Gv,Ev", gc256E, 0, gc2596
sub22AD &H1B4, "LFS.?,FS,Gv,Mp", gc256E, gc258E, gc2596 Or gc2594
sub22AD &H1B5, "LGS.?,GS,Gv,Mp", gc256E, gc258E, gc2596 Or gc2594
sub22AD &H1B6, "MOVZX.b, Gv,Eb", gc256E, 0, gc2596
sub22AD &H1B7, "MOVZX.w, Gv,Ew", gc256E, 0, gc2596
sub22AD &H1BA, "////BT.w,Ew/BS.w,Ew/BTR.w,Ew/BTC.w,Ew/", 0, 0, 0
sub22AD &H1BB, "BTC.?,Ev,Gv", gc256E, 0, gc2596
sub22AD &H1BC, "BSF.?,Gv,Ev", gc256E, 0, gc2596
sub22AD &H1BD, "BSR.?,Gv,Ev", gc256E, 0, gc2596
sub22AD &H1BE, "MOVSX.?,Gv,Eb", gc256E, 0, gc2596
sub22AD &H1BF, "MOVSX.?,Gv,Ew", gc256E, 0, gc2596
sub22AD &H1C0, "XADD.b,Eb,Gb", gc256E, 0, 0
sub22AD &H1C1, "XADD.?,Ev,Gv", gc256E, 0, gc2596
sub22AD &H1C8, "BSWAP.?,eAX", 0, 0, gc2596
sub22AD &H1C9, "BSWAP.?,eCX", 0, 0, gc2596
sub22AD &H1CA, "BSWAP.?,eDX", 0, 0, gc2596
sub22AD &H1CB, "BSWAP.?,eBX", 0, 0, gc2596
sub22AD &H1CC, "BSWAP.?,eSP", 0, 0, gc2596
sub22AD &H1CD, "BSWAP.?,eBP", 0, 0, gc2596
sub22AD &H1CE, "BSWAP.?,eSI", 0, 0, gc2596
sub22AD &H1CF, "BSWAP.?,eDI", 0, 0, gc2596
sub22A3 &HD8, 0, "FADD.f,ST0,Mf", "FADD,ST0,Fi"
sub22A3 &HD8, 1, "FMUL.f,ST0,Mf", "FMUL,ST0,Fi"
sub22A3 &HD8, 2, "FCOM.f,ST0,Mf", "FCOM,ST0,Fi"
sub22A3 &HD8, 3, "FCOMP.f,ST+,Mf", "FCOMP,ST+,Fi"
sub22A3 &HD8, 4, "FSUB.f,ST0,Mf", "FSUB,ST0,Fi"
sub22A3 &HD8, 5, "FSUBR.f,ST0,Mf", "FSUBR,ST0,Fi"
sub22A3 &HD8, 6, "FDIV.f,ST0,Mf", "FDIV,ST0,Fi"
sub22A3 &HD8, 7, "FDIVR.f,ST0,Mf", "FDIVR,ST0,Fi"
sub22A3 &HD9, 0, "FLD.f,ST-,Mf", "FLD.f,ST-,Fi"
sub22A3 &HD9, 1, "", "FXCH,ST0,Fi"
sub22A3 &HD9, 2, "FST.f,Mf,ST0", "FNOP////////"
sub22A3 &HD9, 3, "FSTP.f,Mf,ST+", "FSTP,Fi,ST+"
sub22A3 &HD9, 4, "FLDENV,ENV,M14", "FCHS/FABS///FTST/FXAM///"
sub22A3 &HD9, 5, "FLDCW,CWD,Mw", "FLD1/FLD2T/FLD2E/FLDPI/FLDLG2/FLDLN2/FLDZ//"
sub22A3 &HD9, 6, "FSTENV,M14,ENV", "F2XM1/FYL2X/FPTAN/FPATAN/FXTRACT//FDECSTP/FINCSTP/"
sub22A3 &HD9, 7, "FSTCW,Mw,CWD", "FPREM/FYL2XP1/FSQRT/FSINCOS/FRNDINT/FSCALE/FSIN/FCOS/"
sub22A3 &HDA, 0, "FADD.l,ST0,Ml", ""
sub22A3 &HDA, 1, "FMUL.l,ST0,Ml", ""
sub22A3 &HDA, 2, "FCOM.l,ST0,Ml", ""
sub22A3 &HDA, 3, "FCOMP.l,ST+,Ml", ""
sub22A3 &HDA, 4, "FSUB.l,ST0,Ml", ""
sub22A3 &HDA, 5, "FSUBR.l,ST0,Ml", "/FUCOMPP///////"
sub22A3 &HDA, 6, "FDIV.l,ST0,Ml", ""
sub22A3 &HDA, 7, "FDIVR.l,ST0,Ml", ""
sub22A3 &HDB, 0, "FLD.l,ST-,Ml", ""
sub22A3 &HDB, 2, "FST.l,Ml,ST0", ""
sub22A3 &HDB, 3, "FSTP.l,Ml,ST+", ""
sub22A3 &HDB, 4, "", "FENI/FDISI/FCLEX/FINIT/FSETPM////"
sub22A3 &HDB, 5, "FLD.t,ST-,Mt", ""
sub22A3 &HDB, 7, "FST.t,Mt,ST+", ""
sub22A3 &HDC, 0, "FADD.d,ST0,Md", "FADD,Fi,ST0"
sub22A3 &HDC, 1, "FMUL.d,ST0,Md", "FMUL,Fi,ST0"
sub22A3 &HDC, 2, "FCOM.d,ST0,Md", "FCOM,Fi,ST0"
sub22A3 &HDC, 3, "FCOMP.d,ST+,Md", "FCOMP,Fi,ST+"
sub22A3 &HDC, 4, "FSUB.d,ST0,Md", "FSUB,Fi,ST0"
sub22A3 &HDC, 5, "FSUBR.d,ST0,Md", "FSUBR,Fi,ST0"
sub22A3 &HDC, 6, "FDIV.d,ST0,Md", "FDIV,Fi,ST0"
sub22A3 &HDC, 7, "FDIVR.d,ST0,Md", "FDIVR,Fi,ST0"
sub22A3 &HDD, 0, "FLD.d,ST-,Md", "FFREE,Fi"
sub22A3 &HDD, 1, "", "FXCH,Fi,ST0"
sub22A3 &HDD, 2, "FST.d,Md,ST0", "FST,Fi,ST0"
sub22A3 &HDD, 3, "FSTP.d,Md,ST+", "FSTP,Fi,ST+"
sub22A3 &HDD, 4, "FRSTOR,FST,M94", "/FUCOM///////"
sub22A3 &HDD, 5, "", "/FUCOMP///////"
sub22A3 &HDD, 6, "FSAVE,M94,FST", ""
sub22A3 &HDD, 7, "FSTSW,Mw,FS", ""
sub22A3 &HDE, 0, "FADD.w,ST0,Mw", "FADDP,ST0,Fi"
sub22A3 &HDE, 1, "FMUL.w,ST0,Mw", "FMUL,ST0,Fi"
sub22A3 &HDE, 2, "FCOM.w,ST0,Mw", "FCOM,ST0,Fi"
sub22A3 &HDE, 3, "FCOMP.w,ST+,Mw", "FCOMP,ST+,Fi"
sub22A3 &HDE, 4, "FSUB.w,ST0,Mw", "FSUB,ST0,Fi"
sub22A3 &HDE, 5, "FSUBR.w,ST0,Mw", "FSUBR,ST0,Fi"
sub22A3 &HDE, 6, "FDIV.w,ST0,Mw", "FDIV,ST0,Fi"
sub22A3 &HDE, 7, "FDIVR.w,ST0,Mw", "FDIVR,ST0,Fi"
sub22A3 &HDF, 0, "FLD.w,ST-,Mw", "FFREE,Fi"
sub22A3 &HDF, 1, "", "FXCH,ST0,Fi"
sub22A3 &HDF, 2, "FST.w,Mw,ST0", "FSTP,ST+,Fi"
sub22A3 &HDF, 3, "FSTP.w,Mw,ST+", "FSTP,ST+,Fi"
sub22A3 &HDF, 4, "FLD.n,ST-,Mn", "FCOMPP,ST+,ST+////////"
sub22A3 &HDF, 5, "FLD.q,ST-,Mq", ""
sub22A3 &HDF, 6, "FSTP.n,Mn,ST+", ""
sub22A3 &HDF, 7, "FSTP.q,Mq,ST+", ""
End Sub

Sub sub2324()
  sub2305
  If gv24A4 = 0 Then
    gv24A4 = mc0050
  End If
  If gv24A6 = 0 Then
    gv24A6 = gc2598
  End If
  gv24A8 = gv24A4 Or gv24A6
End Sub

Sub sub2305()
  gAsmOld.M65EC = gIsVB_Ver <> 0
  gAsmOld.M65F7 = gAsmOld.M65EC
End Sub

Sub sub2224(p00DE As T63F1)
Dim l00E0 As Integer
' Const mc00E2 = 9 ' &H9%
' Const mc00E4 = 4 ' &H4%
' Const mc00E6 = 5 ' &H5%
' Const mc00E8 = 8 ' &H8%
' Const mc00F2 = 15 ' &HF%
' Const mc00F6 = 25 ' &H19%
' Const mc00F8 = 26 ' &H1A%
' Const mc00FA = 7 ' &H7%
  If m0014 = 0 Then Exit Sub
  l00E0 = p00DE.subRef_M3C6A
  If l00E0 = 0 Then
    Select Case p00DE.M63E9
    Case mc00E2
      GoSub L50AE2
      p00DE.M63E9 = mc00E4
    Case mc00E6
      GoSub L50AE2
    Case mc00E4, mc00E8
      GoSub L50AE2
    Case Else
      Exit Sub
    End Select
    If l00E0 Then
      p00DE.subRef_M3C6A = l00E0
    Else
      Exit Sub
    End If
  End If
    If gCmdByteStruc.M642E And gc259A Then
      If gReloc_Arr(l00E0).M36B3_count = mc00F2 Then
        MsgBox "found __vbaErrorOverflow"
        gCmdByteStruc.M642E = gCmdByteStruc.M642E Or gc2598
      End If
    End If
Exit Sub

L50AE2:
    If p00DE.M3C63 = 0 Then
      If p00DE.M4106 >= &H10000 Then
        l00E0 = CreateSym(p00DE.M4106)
        If l00E0 Then
          p00DE.M4106 = 0
        Else
          MsgBox "sym not created?"
        End If
      Else
        Select Case gAsmNew.M6603
        Case mc00F6
          l00E0 = CreateSym(p00DE.M4106)
        Case mc00F8
          p00DE.M63E9 = mc00FA
        Case 0
          MsgBox Hex$(p00DE.M4106)
        Case Else
          Return
        End Select
      End If
    End If
    If l00E0 = 0 Then Return
  If p00DE.M4106 Then
    l00E0 = CreateSym(fn202F(l00E0) + p00DE.M4106)
    If l00E0 Then
      p00DE.M4106 = 0
    Else
      ErrBeep
    End If
  End If
Return
End Sub

Sub sub0CB5()
  If gv25D8 = 0 Then sub22F6
  sub2324
End Sub

Private Function fn2288(p0100 As Integer) As Integer
  fn2288 = gv25D8 + 1
  gv25D8 = gv25D8 + p0100
  ReDim Preserve gCmdByteArr(gv25D8)
End Function

Sub sub2295(p0102 As Integer, p0104 As Integer, p0106 As Integer)
  m0060(p0102) = p0104
  m0076(p0102) = p0106
End Sub

Sub sub22A3(p0108 As Integer, p010A As Integer, ByVal p010C As String, ByVal p010E As String)
Dim l0110 As Integer
  l0110 = gCmdByteArr(p0108).M6426
  If l0110 = 0 Then
    l0110 = fn2288(16)
    gCmdByteArr(p0108).M641D = gc2574
    gCmdByteArr(p0108).M6426 = l0110
  End If
  l0110 = l0110 + p010A
  If Len(p010C) Then
    sub22AD l0110, p010C, gc256A, 0, 0
  End If
  l0110 = l0110 + 8
  If Len(p010E) Then
    If InStr(p010E, "/") Then
      sub22AD l0110, p010E, gc2572, 0, 0
    Else
      sub22AD l0110, p010E, gc256A, 0, 0
    End If
  End If
End Sub

Sub sub22AD(ByVal pv0112 As Integer, ByVal p0114 As String, ByVal pv0116 As Integer, ByVal pv0118 As Integer, ByVal pv011A As Integer)
Dim l011E As T6407
Dim l0120
Dim l0122 As String
Dim l0124 As Integer
Dim l0126 As String
Dim l0128 As Integer
' Const mc012A = 24 ' &H18%
' Const mc012C = 27 ' &H1B%
' Const mc012E = 28 ' &H1C%
' Const mc0130 = 29 ' &H1D%
  l011E.M642E = pv011A
  If InStr(p0114, "/") Then
    l0128 = fn2288(8)
    l011E.M6426 = l0128
    If pv0116 = gc2572 Then
      l011E.M641D = pv0116
    Else
      l011E.M641D = gc2570
    End If
    gCmdByteArr(pv0112) = l011E
    l0122 = p0114
    For l0120 = 0 To 7
      l0124 = InStr(l0122, "/")
      If l0124 > 1 Then sub22AD l0128 + l0120, Left$(l0122, l0124 - 1), gc256A, 0, pv011A
      l0122 = Mid$(l0122, l0124 + 1)
    Next
    Exit Sub
  End If
  If pv0118 = 0 Then
    If pv0116 = gc257C Then
      Select Case pv0112
      Case &H26: pv0118 = mc012A
      Case &H2E: pv0118 = mc00F6
      Case &H36: pv0118 = mc00F8
      Case &H3E: pv0118 = mc012C
      Case &H64: pv0118 = mc012E
      Case &H65: pv0118 = mc0130
      End Select
    Else
      Select Case Left$(p0114, 3)
      Case "CAL": pv0118 = gc2582
      Case "JMP": pv0118 = gc2580
      Case "LOO": pv0118 = gc257E
      Case "RET": pv0118 = gc2584
      Case "IRE": pv0118 = gc2584: pv011A = pv011A Or gc259E
      Case "PUS": pv0118 = gc2586
      Case "POP": pv0118 = gc2588
      Case Else
        If Left$(p0114, 1) = "J" Then
          pv0118 = gc257E
        End If
      End Select
    End If
  End If
  If pv0116 = 0 Then
    l011E.M641D = gc256A
  Else
    l011E.M641D = pv0116
  End If
  l011E.M6426 = pv0118
  Select Case pv0118
  Case gc2580, gc2584
    pv011A = pv011A Or gc2598
  Case gc2582
    pv011A = pv011A Or gc259A
  End Select
  l0122 = p0114 & ","
  l0124 = InStr(l0122, ",")
  l0126 = Left$(l0122, l0124 - 1)
  l0120 = InStr(l0126, "\")
  If l0120 Then
    l011E.M6452 = fn2367(Mid$(l0126, l0120 + 1))
    l0126 = Left$(l0126, l0120 - 1)
    l011E.M6448 = fn2367(l0126)
  Else
    l011E.M6448 = fn2367(l0126)
    l011E.M6452 = l011E.M6448
  End If
  Select Case Right$(l0126, 2)
  Case ".b": l011E.M6438 = 1
  Case ".w", ".s": l011E.M6438 = 2
  Case ".?": l011E.M6438 = 2: pv011A = pv011A Or gc2596
  Case ".l", ".f": l011E.M6438 = 4
  Case ".p": l011E.M6438 = 4: pv011A = pv011A Or gc2596
  Case ".6": l011E.M6438 = 6
  Case ".d", ".q": l011E.M6438 = 8
  Case ".t", ".n": l011E.M6438 = 10
  Case ".N", ".F"
  Case Else
    If InStr(l0126, ".") Then
      MsgBox l0126, , "unknown op size"
    End If
  End Select
  l011E.M642E = pv011A
  l0122 = Trim$(Mid$(l0122, l0124 + 1))
  sub21F3 l011E, l0122
  gCmdByteArr(pv0112) = l011E
End Sub

Sub DisAsmCmd()
Dim l0134
Dim CmdByte2 As Integer
Dim CmdByte As Integer
Dim l013C As Integer
' Const mc0140 = 32 ' &H20%
' Const mc0142 = 37 ' &H25%
' Const mc0144 = 36 ' &H24%
' Const mc0146 = 21 ' &H15%
 
 'Last End Get's new Start
  gVACmdStart = gVACmdEnd
  gAsmNew = gAsmOld
  
  If m0014 < 0 Then gDisOutputTxt = ""

L515DC:
  CmdByte2 = 0
L515EA:
  Get g_hFileExe, gVAToRaw + gVACmdEnd, gSingleChar: gVACmdEnd = gVACmdEnd + 1
  CmdByte = Asc(gSingleChar)
  CmdByte2 = CmdByte2 + CmdByte
L51634:
  gCmdByteStruc = gCmdByteArr(CmdByte2)
  On gCmdByteStruc.M641D GoTo L517FA, L516A8, L517EE, L516BE, L516E8, L5170C, L5175C, L5178E, L517AE, L517CE

L5166E:
  sub2232
  If m0014 < 0 Then gDisOutputTxt = "<illegal>"
  Exit Sub

L516A8:
  CmdByte2 = &H100
  GoTo L515EA
L516BE:
  GoSub L518FE
  CmdByte2 = gCmdByteStruc.M6426 + gAsmNew.M661F
  GoTo L51634
L516E8:
  CmdByte2 = gCmdByteStruc.M6426 + gAsmNew.M6627
  GoTo L51634
L5170C:
  GoSub L518FE
  CmdByte2 = gCmdByteStruc.M6426 + gAsmNew.M661F
  If gAsmNew.M6616 = 3 Then
    CmdByte2 = CmdByte2 + 8
  End If
  GoTo L51634

L5175C:
  If m0014 < 0 Then gDisOutputTxt = gDisOutputTxt & fn2375()
  GoTo L515DC
L5178E:
  gAsmNew.M65F7 = Not gAsmNew.M65F7
  GoTo L515DC
L517AE:
  gAsmNew.M65EC = Not gAsmNew.M65EC
  GoTo L515DC
L517CE:
  gAsmNew.M6603 = gCmdByteStruc.M6426
  GoTo L515DC

L517EE:
  GoSub L518FE
L517FA:
  If gCmdByteStruc.M642E And gv24A4 Then GoTo L5166E
  If m0014 < 0 Then gDisOutputTxt = gDisOutputTxt & fn2375()
  gv272C = gCmdByteStruc.M6414
  
  For l0134 = 1 To gv272C
    If m0014 Then
      If m0014 < 0 Then
        If l0134 > 1 Then gDisOutputTxt = gDisOutputTxt & ","
      End If
      sub2218 gCmdByteStruc.M6441(l0134), gAsmNew.M6441(l0134), m0014
    Else
      sub2252 gCmdByteStruc.M6441(l0134)
    End If
  Next
Exit Sub

L518FE:
  If l013C Then Return
  l013C = True
  Get g_hFileExe, gVAToRaw + gVACmdEnd, gSingleChar
  gAsmNew.M660E = Asc(gSingleChar): gVACmdEnd = gVACmdEnd + 1
  gAsmNew.M6616 = gAsmNew.M660E \ &O100
  gAsmNew.M661F = (gAsmNew.M660E And &O70) \ &O10
  gAsmNew.M6627 = gAsmNew.M660E And &O7
  
  If gAsmNew.M6616 = 3 Then Return
  If gAsmNew.M65F7 Then
    Select Case gAsmNew.M6627
    Case 4
      Get g_hFileExe, gVAToRaw + gVACmdEnd, gSingleChar
      gAsmNew.M6632 = Asc(gSingleChar): gVACmdEnd = gVACmdEnd + 1
      gAsmNew.M663A = 2 ^ (gAsmNew.M6632 \ &O100)
      gAsmNew.M6643 = (gAsmNew.M6632 And &O70) \ &O10
      If gAsmNew.M6643 = 4 Then
        gAsmNew.M6643 = 0
      Else
        gAsmNew.M6643 = gAsmNew.M6643 + mc0140
      End If
      gAsmNew.M664D = gAsmNew.M6632 And &O7
      If gAsmNew.M664D = 5 And gAsmNew.M6616 = 0 Then
        gAsmNew.M664D = 0
        gAsmNew.M6656 = fn22D8(4, gAsmNew.subRef_M3C6A)
      Else
        gAsmNew.M664D = gAsmNew.M664D + mc0140
      End If
    Case 5
      If gAsmNew.M6616 Then
        gAsmNew.M664D = gAsmNew.M6627 + mc0140
      Else
        gAsmNew.M664D = 0
        gAsmNew.M6656 = fn22D8(4, gAsmNew.subRef_M3C6A)
      End If
    Case Else
      gAsmNew.M664D = gAsmNew.M6627 + mc0140
    End Select
    If gAsmNew.M6603 <= 0 Then
      Select Case gAsmNew.M664D
      Case 0
        If gAsmNew.M6643 = mc0142 Then
          gAsmNew.M6603 = mc00F8
        Else
          gAsmNew.M6603 = mc012C
        End If
      Case mc0142, mc0144
        gAsmNew.M6603 = mc00F8
      Case Else
        gAsmNew.M6603 = mc012C
      End Select
    End If
    Select Case gAsmNew.M6616
    Case 1
      gAsmNew.M6656 = fn22D8(1, gAsmNew.subRef_M3C6A)
    Case 2
      gAsmNew.M6656 = fn22D8(4, gAsmNew.subRef_M3C6A)
    End Select
  Else
    gAsmNew.M664D = m0060(gAsmNew.M6627)
    gAsmNew.M6643 = m0076(gAsmNew.M6627)
    Select Case gAsmNew.M6616
    Case 0
      If gAsmNew.M6627 = 6 Then
        gAsmNew.M6656 = fn22D8(2, gAsmNew.subRef_M3C6A)
        gAsmNew.M664D = 0
      End If
    Case 1
      gAsmNew.M6656 = fn22D8(1, gAsmNew.subRef_M3C6A)
    Case 2
      gAsmNew.M6656 = fn22D8(2, gAsmNew.subRef_M3C6A)
    End Select
    If gAsmNew.M6603 <= 0 Then
      If gAsmNew.M664D = mc0146 Then
        gAsmNew.M6603 = mc00F8
      Else
        gAsmNew.M6603 = mc012C
      End If
    End If
  End If
Return
End Sub

Sub sub1E89(p0148 As Long, p014A As Long, p014C As Integer)
Dim l014E As Integer
Static m0150 As Integer
' Const mc0158 = 30 ' &H1E%
  gv34F4 = m0150
  gVACmdEnd = p0148
  gVAToRaw2 = fn15A9(gVACmdEnd) - gVACmdEnd
  m0014 = 1
  Do
    gVACmdStart = gVACmdEnd
    DisAsmCmd
    If gCmdByteStruc.M642E And mc0050 Then
L51EA0:
      If gv34F4 Then
        ErrBeep2
      Else
        MsgBox "unknown opcode @" & Hex$(gVACmdStart)
        gv34F4 = mc0158
        gv0462 = gVACmdStart
        If p014C Then Exit Sub
      End If
      l014E = SomeAdrConv(gVACmdEnd)
      While fn202F(l014E) <= gVACmdEnd
        l014E = fn2052(l014E)
      Wend
      gVACmdEnd = fn202F(l014E)
    ElseIf gCmdByteStruc.M642E And gc2598 Then
      l014E = CreateSym(gVACmdEnd)
    End If
  Loop Until gVACmdEnd >= p014A
End Sub

Function fn2334(p0162 As String, p0164 As Integer) As String
  If Len(p0162) < p0164 Then
   fn2334 = String$(p0164 - Len(p0162), "0") + p0162
  Else
   fn2334 = p0162
  End If
End Function
