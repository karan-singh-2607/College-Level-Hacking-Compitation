Attribute VB_Name = "MODULE27"
' Module27
Option Explicit
Dim mFrm As frmBlkFile
Dim mhFile As Integer
Const gb2 = "gb2"
Const va2 = "va2"
'Const &H107 = 263 ' &H107%
'Const &H404 = 1028 ' &H404%
Const mc003C = -1 ' &HFFFF%
Const mc00DE = 4 ' &H4%
Const mc0164 = 4 ' &H4%
Const mc0166 = 32 ' &H20%

Function SaveScope(Frm As Form, p0130 As Integer, p0132 As Long) As Long
  gDataSize = p0130
  gv16C2 = p0132
  Frm.cmd_alloc.value = True
  SaveScope = gv16C2
End Function

Function openVa2(Frm As Form, Filename As String, p0140%, p0142%) As Integer
On Error Resume Next
  Frm.cmd_close.value = True
On Error GoTo 0
  
  If Len(Dir$(Filename)) Then
    Kill Filename
  End If
  
  openVa2 = openVa(Frm, Filename, p0140, p0142)
End Function

Function fn19B7(p0148 As Form, p014A As Integer) As Long
  gDataSize = p014A
  p0148.cmd_locate.value = True
  fn19B7 = gv16C2
End Function

Function openVa(frmBlk As Form, Filename As String, DataVersion As Integer, p0156 As Integer) As Integer
  Do
    gFilename = Filename
    gDataSize = p0156
    gDataVersion = DataVersion
    
    frmBlk.cmd_open.value = True
    If gVAFileName <= 0 Then Exit Function
    If gDataVersion = DataVersion Then Exit Do
    DbgMileStone "mismatch DataVersion"
    
    UnloadFrm frmBlk
    
    Kill Filename
  
  Loop
  
  openVa = gVAFileName

End Function

Function fn1A35(p0102%) As Integer
  If mhFile <= 0 Then
    fn1A35 = -1
  ElseIf fn19B7(mFrm, p0102) < 2 Then
    fn1A35 = -1
  Else
    fn1A35 = fn1A54(g_gbl(p0102).M4950)
  End If
End Function

Private Function fn1A43(p0106 As Integer) As Integer
    If gEXEdefEntriesCount2 > 0 Then
      p0106 = gEXEdefEntriesCount2: gEXEdefEntriesCount2 = gEXEdefEntries(p0106).Count
    Else
      gEXEdefEntriesCount1 = gEXEdefEntriesCount1 + 1: p0106 = gEXEdefEntriesCount1
      If p0106 > UBound(gEXEdefEntries) Then ReDim Preserve gEXEdefEntries(p0106 + 100)
    End If
    Get mhFile, , gEXEdefEntries(p0106)
    fn1A43 = p0106
End Function

Private Function fn1A54(p010A As Integer) As Integer
Dim l010C As Integer
Dim i%
Dim tmInt%
  If p010A = -1 Then p010A = 0
  Get mhFile, , tmInt
  If tmInt <= 0 Then
    p010A = 0
  Else
    l010C = fn1A43(p010A)
    For i = 2 To tmInt
      l010C = fn1A43(gEXEdefEntries(l010C).Count)
    Next
  End If
  fn1A54 = tmInt
End Function

Function SaveSubToVa(p0114 As Integer, p0116 As Integer, p0118 As Integer) As Integer
Dim l011A As Integer
Dim strucSize As Integer
Dim l011E As Long
Dim l0120 As Integer
  l0120 = g_gbl(p0114).M4950
  If l0120 > 0 Then
    Do
      strucSize = strucSize + 1
      l011A = gEXEdefEntries(l0120).Count
      If l011A = 0 Then Exit Do
      l0120 = l011A
    Loop
    l011E = Len(gEXEdefEntries(0)) * strucSize
  End If
  
  If p0118 Then
    l011E = l011E + fn1D79()
  End If
  
  If l011E = 0 Then Exit Function
  
  If SaveScope(mFrm, p0114, l011E + 2) < l011E Then
    MsgBox "error saving scope " & p0114
    SaveSubToVa = -1
    Exit Function
  End If
  
  Put mhFile, , strucSize
  
  If strucSize Then
    If strucSize <> fn1A85(g_gbl(p0114).M4950) Then DbgMileStone "error saving scope"
    SaveSubToVa = strucSize
    
    If p0116 Then
      gEXEdefEntries(l0120).Count = gEXEdefEntriesCount2
      gEXEdefEntriesCount2 = g_gbl(p0114).M4950
      g_gbl(p0114).M4950 = mc003C
    End If
  
  ElseIf p0116 Then
    g_gbl(p0114).M4950 = mc003C
  End If
  
  If p0118 Then
    sub1A74 mhFile
  End If
End Function

Private Function fn1A85(ByVal pv0126 As Integer) As Integer
Dim l0128 As Integer
  While pv0126 > 0
    l0128 = l0128 + 1
    Put mhFile, , gEXEdefEntries(pv0126)
    pv0126 = gEXEdefEntries(pv0126).Count
  Wend
  fn1A85 = l0128
End Function

Sub sub0299(p00C8 As Integer, p00CA%)
Dim l00CC As Integer
' Const mc00DE = 4 ' &H4%
  If p00C8 > gEXEdef_Arr5Count Then ErrBeep2: Exit Sub
  If gv16BC = p00C8 Then
    If g_gbl(gL3Max + p00C8).M4950 <> -1 Then Exit Sub
  End If
  
  If gIs3158 Then SaveSub True
  Assume gEXEdef_Subs(p00C8).ModuleIndex, "mod known"
  If gEXEdef_Subs(p00C8).ModuleIndex = 0 Then
    gv16BC = p00C8
    Exit Sub
  End If
  
  sub19F0 gEXEdef_Subs(p00C8).ModuleIndex, p00CA
  gL3Min = gL3Max + p00C8
  gL2Min = gL2Max + p00C8
  gv315A = -g_gbl(gL3Min).Size
  gv315C = g_gbl(gL2Min).Size
  sub0EB3 p00C8
  
  If fn1A35(gL3Min) < 0 Then
    sub17C9 p00C8, mc00DE
    SaveSub False
  Else
    sub1A12 mhFile
    gv16BC = p00C8
    sub101B
  End If
  
  If p00CA Then
    l00CC = gEXEdef_Subs(p00C8).M4289
    If l00CC Then
      frm3.cmdSync.value = 1
      sub0A5E l00CC
    End If
  End If

End Sub

Sub SaveAll(p003E%)
Dim i As Integer
Dim ii As Integer
Dim hFile As Integer
If g_IsDemoVersion Then
Else
  sub1943 p003E
  
  If gIs3124 Then
    
    gFileName_dat = CreateFullFileName("", gb2)
    
    hFile = FreeFile
    
    Open gFileName_dat For Binary Access Write As hFile
    
   'Write Ver
    ii = &H107: Put hFile, , ii
    
    Put hFile, , g_decl_num
    Put hFile, , gEXEdef1
    Put hFile, , gEXEdef2
    Put hFile, , gEXEdef_Arr5Count
    Put hFile, , gEXEdef4
    Put hFile, , gEXEdef5
    
    Put hFile, , gEXEdef6_Num_scopes
    
    Put hFile, , gEXEdefEntriesCount1
    
    For i = 1 To gEXEdef6_Num_scopes
      Put hFile, , g_gbl(i)
    Next
    
    Put hFile, , g_userTypesCount
    For i = 1 To g_userTypesCount
      Put hFile, , gEXEdefEntries2(i)
    Next
    
    Put hFile, , gEXEdef_Arr1Count
    For i = 1 To gEXEdef_Arr1Count
      Put hFile, , gEXEdef_Arr1(i)
    Next
    
    Put hFile, , gEXEdef_Arr2Count
    For i = 1 To gEXEdef_Arr2Count
      Put hFile, , gEXEdef_Arr2(i)
    Next
    
    Put hFile, , gReloc_ArrCount
    
    Put hFile, , gEXEdef_Arr4_importsCount
    
    i = 0: Put hFile, , i
    
    For i = 1 To g_decl_num
      Put hFile, , g_decl_objs(i)
    Next
    
    For i = 1 To gEXEdef_Arr5Count
      Put hFile, , gEXEdef_Subs(i)
    Next
    
    For i = 1 To gReloc_ArrCount
      Put hFile, , gReloc_Arr(i)
    Next
    
    For i = 1 To gEXEdef_Arr4_importsCount
      Put hFile, , gEXEdef_Arr4_imports(i)
    Next
    
    Close hFile
    
    If mhFile = 0 Then
      mhFile = openVa2(mFrm, CreateFullFileName("", va2), &H404, gEXEdef6_Num_scopes)
    End If
    
    For i = 1 To g_decl_num
      
      DebugPrintVar g_gbl(gLWhatEver + i).Size = g_decl_objs(i).Size, "gbl size"
      '''
      ii = SaveSubToVa(gLWhatEver + i, p003E, False)
      If g_decl_objs(i).statSize Then
        ii = SaveSubToVa(gL1Max + i, p003E, False)
      End If
    Next
    
    For i = 1 To gEXEdef2
      ii = SaveSubToVa(gL0Max + i, p003E, False)
    Next
    
    For i = 1 To gEXEdef_Arr5Count
      ii = SaveSubToVa(gL2Max + i, p003E, False)
    Next
    
    For i = gL4Max + 1 To gEXEdef6_Num_scopes
      ii = SaveSubToVa(i, p003E, False)
    Next
  
  End If
End If
  If p003E Then
    If mhFile Then Unload mFrm: mhFile = 0
  End If
  gIs3124 = False
End Sub

Sub sub0DC9()
Dim hFile As Integer
Dim j%
Dim TwoByteTmp%
Dim i%
Dim Ver%
' Const mc0164 = 4 ' &H4%
' Const mc0166 = 32 ' &H20%

  SaveAll True
  InitSth2
  gIs3124 = True
  gv16BA = 0
  gv16BC = 0: gIs3158 = False
  gFileName_dat = CreateFullFileName("", gb2)

If g_IsDemoVersion Then
Else
  If gIsVB5 < 0 Then
    GoTo SkipLoad_gbFile
  End If
  
 'Open Gb2
  If Len(Dir$(gFileName_dat)) Then
    hFile = FreeFile
    Open gFileName_dat For Binary Access Read As hFile
    Get hFile, , Ver
    Select Case Ver
    
   '&H107 = Expect Version
    Case &H107

'!!! Attention: Line added for better Debugging
'Please don't forget to remove later
GoTo SkipLoad_gbFile


      If MsgBox("use existing EXE definitions", mc0164 Or mc0166) = &H7 Then GoTo SkipLoad_gbFile
      Get hFile, , TwoByteTmp: Assume TwoByteTmp = g_decl_num, "prj match"
      Get hFile, , TwoByteTmp: Assume TwoByteTmp = gEXEdef1, "prj match"
      Get hFile, , TwoByteTmp: Assume TwoByteTmp = gEXEdef2, "prj match"
      Get hFile, , TwoByteTmp: Assume TwoByteTmp = gEXEdef_Arr5Count, "prj match"
      Get hFile, , TwoByteTmp: Assume TwoByteTmp = gEXEdef4, "prj match"
      Get hFile, , TwoByteTmp: Assume TwoByteTmp = gEXEdef5, "prj match"
      Get hFile, , gEXEdef6_Num_scopes: ReDim g_gbl(gEXEdef6_Num_scopes)
      Get hFile, , TwoByteTmp
      ReDim gEXEdefEntries(TwoByteTmp): gEXEdefEntriesCount1 = 0: gEXEdefEntriesCount2 = 0
      
      For i = 1 To gEXEdef6_Num_scopes
        Get hFile, , g_gbl(i)
        
        If g_gbl(i).M4950 Then
          g_gbl(i).M4950 = mc003C
        End If
      
      Next
      
      Get hFile, , g_userTypesCount
      
      ReDim gEXEdefEntries2(g_userTypesCount)
      
      For j = 1 To g_userTypesCount
        Get hFile, , gEXEdefEntries2(j)
      Next
      
      Get hFile, , gEXEdef_Arr1Count
      
      ReDim gEXEdef_Arr1(gEXEdef_Arr1Count)
      For j = 1 To gEXEdef_Arr1Count
        Get hFile, , gEXEdef_Arr1(j)
      Next
      
      Get hFile, , gEXEdef_Arr2Count
      ReDim gEXEdef_Arr2(gEXEdef_Arr2Count)
      For j = 1 To gEXEdef_Arr2Count
        Get hFile, , gEXEdef_Arr2(j)
      Next
      
      Get hFile, , gReloc_ArrCount: ReDim gReloc_Arr(gReloc_ArrCount)
      Get hFile, , gEXEdef_Arr4_importsCount: ReDim gEXEdef_Arr4_imports(gEXEdef_Arr4_importsCount)
      
      Get hFile, , TwoByteTmp
      If TwoByteTmp Then DbgMileStone "more gbls?"
      
      For j = 1 To g_decl_num
        Get hFile, , g_decl_objs(j)
      Next
      For j = 1 To gEXEdef_Arr5Count
        Get hFile, , gEXEdef_Subs(j)
      Next
      For j = 1 To gReloc_ArrCount
        Get hFile, , gReloc_Arr(j)
      Next
      For j = 1 To gEXEdef_Arr4_importsCount
        Get hFile, , gEXEdef_Arr4_imports(j)
      Next
      
     'open VariablesFile
      gFileName_dat = CreateFullFileName("", va2)
      
      If mhFile Then Unload mFrm
      Set mFrm = New frmBlkFile
      mhFile = openVa(mFrm, gFileName_dat, &H404, 0)
      
      If gDataSize < gEXEdef6_Num_scopes Then
      ' it's a New VariablesFile
        If gDataSize Then DbgMileStone "expect 0 für neu!"
        GoTo SkipLoad_gbFile
      End If
    ' Load Existing
      For i = 1 To g_decl_num
        
        Assume g_gbl(gLWhatEver + i).Size = g_decl_objs(i).Size, "gbl size"
        j = fn1A35(gLWhatEver + i)
        
        Assume g_gbl(gL1Max + i).Size = g_decl_objs(i).statSize, "stat size"
        If g_gbl(gL1Max + i).Size Then
          If fn1A35(gL1Max + i) < 0 Then
            ErrBeep
            GoTo SkipLoad_gbFile
          End If
        Else
          g_gbl(gL1Max + i).M4950 = 0
        End If
      Next
      
      For i = 1 To gEXEdef2
        j = fn1A35(gL0Max + i)
      Next
      
      For i = 1 To gEXEdef_Arr5Count
        j = fn1A35(gL2Max + i)
      Next
      
      For i = gL4Max + 1 To gEXEdef6_Num_scopes
        j = fn1A35(i)
      Next
      
      gIs3124 = False
      
    Case Else
    'Other Version
      Close hFile
      Kill gFileName_dat
    End Select
SkipLoad_gbFile:
    Close hFile
  End If
End If

  If gIs3124 Then
  
   'open varfile
    gFileName_dat = CreateFullFileName("", va2)
    
    If mhFile Then Unload mFrm
    Set mFrm = New frmBlkFile
    mhFile = openVa2(mFrm, gFileName_dat, &H404, gEXEdef6_Num_scopes)
    
    Assume mhFile, "varfile"
    sub1925
  Else
    sub0FE8
  End If
  sub19DD
End Sub

Sub SaveSub(p00E4%)
  If gL3Min Then
  
    If gIs3158 Then
      If SaveSubToVa(gL3Min, p00E4, True) < 0 Then
        MsgBox "error save sub"
      End If
    End If
    
    If p00E4 Then
      sub1A23 gL3Min: gL2Min = 0
      ReDim gBuggy(10): gvArrSize = 0
    End If
    
  End If
  
  gIs3158 = False
  
End Sub

Sub sub1943(p00C6%)
  SaveSub p00C6
End Sub

Sub UnloadFrm(Form As Form)
On Error Resume Next
  Unload Form
End Sub

Sub sub19F0(mod_nr As Integer, p00BE As Integer)
  If mod_nr = gv16BA Then Exit Sub
  DbgVar mod_nr, "mod nr"
  If mod_nr = 0 Then Exit Sub
  gv16BA = mod_nr
  SaveSub True
  gv16BC = 0
  gL1Min = gL1Max + mod_nr
If gIsVB5 Then
  sub17F6 gv16BA
End If
  If p00BE Then
    If gv097C Then
      sub1732 mod_nr
      sub19FF
    End If
  End If
End Sub

Private Sub sub1A23(p00F0 As Integer)
Dim l00F2 As Integer
Dim l00F4 As Integer
  If p00F0 > gEXEdef6_Num_scopes Then Exit Sub
  l00F4 = g_gbl(p00F0).M4950
  If l00F4 <= 0 Then Exit Sub
  Do
    l00F2 = gEXEdefEntries(l00F4).Count
    If l00F2 = 0 Then Exit Do
    l00F4 = l00F2
  Loop
  gEXEdefEntries(l00F4).Count = gEXEdefEntriesCount2
  gEXEdefEntriesCount2 = g_gbl(p00F0).M4950
  g_gbl(p00F0).M4950 = -1
End Sub

