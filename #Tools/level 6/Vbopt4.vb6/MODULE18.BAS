Attribute VB_Name = "MODULE18"
' Module18
Option Explicit
Dim SegSize As Long
Dim m0026 As Integer
Const mc0028 = 1 ' &H1%
Const mc002A = 2 ' &H2%
Const mc002C = 4 ' &H4%
Const mc004C = 32 ' &H20%
Const mc0050 = 39 ' &H27%
Const mc0054 = 16 ' &H10%
Const mc0056 = 31 ' &H1F%
Const NewHdr_Type = 2 ' &H2%
Const mc009C = 1 ' &H1%
Const mc00C6 = 256 ' &H100%
Const mc00DA = 48 ' &H30%
Const mc0152 = 14 ' &HE%
Const mc015E = 10 ' &HA%
Const mc0160 = 11 ' &HB%
Const mc0174 = 4 ' &H4%
Const mc0176 = 16 ' &H10%
Const mc0180 = 3 ' &H3%
Const mc0182 = 18 ' &H12%
Const SegmentTable_Type = 6 ' &H6%
Const SegTypeWriteableFlag = 1 ' &H1%
Const SegTypeData = 9 ' &H9%
Const SegTypeCode = 8 ' &H8%

Function fn14AE(p0126%, p0128 As Long, p012A As Integer, p012C As Integer) As Integer
Dim l012E As Long
  l012E = fn202F(p0126) + p0128
  If p012A = 0 Then
    p012A = fn1624(l012E)
  ElseIf gv161C(p012A).VOffset > l012E Then
    p012A = fn1624(l012E)
    If p012A = 0 Then Exit Function
  End If
  Do While gv161C(p012A).VOffset < l012E
    p012A = gv161C(p012A).Count
    If p012A = 0 Then GoTo L2597A
  Loop
If gIsVB_Ver Then
Else
  If gv161C(p012A).VOffset = l012E + 2 Then
    If gv161C(p012A).M36B3_count = gc1164 Then
      sub00B9 gv161C(p012A)
    Else
      GoTo L2597A
    End If
  End If
End If
  If gv161C(p012A).VOffset = l012E Then
    fn14AE = gv161C(p012A).subRef_M3C6A
    Exit Function
  End If

L2597A:
  If p012C Then
    Get g_hFileExe, fn1483(p0126) + p0128, l012E
    If l012E Then MsgBox fn2023(p0126) & " + " & Hex$(p0128), , "no valid pointer at"
  End If
End Function

Sub sub00B9(p0110 As T54AE)
Dim l0112%
Dim l0114 As Integer
Dim l0116 As Integer
  Select Case p0110.M36B3_count
  Case gc1164
    Get g_hFileExe, VAtoRaw(p0110.VOffset - 2), l0112
    l0114 = fn208C(fn202F(p0110.subRef_M3C6A))
    l0116 = fn06CC(l0114, ToUnSigned(l0112))
    If l0116 Then
      p0110.subRef_M3C6A = l0116
      p0110.VOffset = p0110.VOffset - 2
      p0110.M36B3_count = gc1168
    Else
    End If
  Case gc1166, gc1168
  Case Else
    ErrBeep
  End Select
End Sub

Sub sub06F3()
Dim i%
Dim Offset As Long
' Const NewHdr_Type = 2 ' &H2%
' Const mc009C = 1 ' &H1%
  InitHeadersVars Offset
  AddToLogTree NewHdr_Type, "NewHdr", Offset, 0, Len(gNE_Header), 1
  
  SegSize = 2 ^ gNE_Header.MiscFlags
  
  sub0672
  
 'Create Table for Imports & fill in valuesfor this module
  ReDim gImportDlls(gImportDLLCount)
  gImportDlls(0).ImageBase = &H10000
  gImportDlls(0).VOffset = (gSegCount + 1) * &H10000
  
  
  gImageBase1 = 0
  
  LoadSegmentTable
  
  PrepareMemPages 0, 0, 0
  
  For i = 1 To gImportDLLCount
    Offset = (i + gSegCount) * &H10000
    sub140D i, Offset, Offset + &H10000
  Next
  
  sub141D
  sub1433
  sub1444
  sub1455
  gEntryPoint_Page = fn13C6(gNE_Header.Initial_CS, gNE_Header.Initial_IP, "WinEntry", mc009C)
  sub1467
End Sub

Sub sub0D9B(p011A%, frmName$)
Dim l011E As Integer
  l011E = MakeSubRef(p011A, frmName, 0)
End Sub

Function MakeSubRef(OffsetSym As Integer, Name As String, TypeCode As Integer) As Integer
Dim subRef As Integer
Dim lOffsetSym As Integer
' Const mc004C = 32 ' &H20%
' Const mc0050 = 39 ' &H27%
' Const mc0054 = 16 ' &H10%
' Const mc0056 = 31 ' &H1F%
  If Len(Name) = 0 Then ErrBeep: Exit Function
  
  lOffsetSym = OffsetSym
  subRef = gReloc_Arr(OffsetSym).subRef_M3C6A
  
  Do While subRef < 0
    lOffsetSym = -subRef
    subRef = gReloc_Arr(lOffsetSym).subRef_M3C6A
  Loop
  
  If subRef = 0 Then
    
   'Inc and Redim
    gEXEdef_Arr4_importsCount = gEXEdef_Arr4_importsCount + 1
    If gEXEdef_Arr4_importsCount > UBound(gEXEdef_Arr4_imports) Then ReDim Preserve gEXEdef_Arr4_imports(gEXEdef_Arr4_importsCount + 10)
    
    subRef = gEXEdef_Arr4_importsCount
    gReloc_Arr(lOffsetSym).subRef_M3C6A = subRef
    gEXEdef_Arr4_imports(subRef).lOffsetSym_M3C6A = lOffsetSym
  Else
    Assume subRef > 0, "sub ref"
  End If
  
  gEXEdef_Arr4_imports(subRef).Name = Name
  
  If TypeCode Then
    Select Case gReloc_Arr(OffsetSym).M36B3_count
    Case mc004C
      Assume TypeCode < 0, "ref"
      Assume -TypeCode <= gEXEdef_Arr5Count, "sub ref"
    Case mc0050
      Assume TypeCode < 0, "ref"
      Assume -TypeCode <= g_userTypesCount, "udt ref"
    Case mc0054 To mc0056
      Assume TypeCode >= 0, "var type?"
    Case Else
      Assume TypeCode > 0, "ref"
    End Select
    
    gEXEdef_Arr4_imports(subRef).TypeCode_M5F96 = TypeCode
    
  End If
  
  MakeSubRef = subRef
  
End Function

Function fn13C6(p0064 As Integer, ByVal p0066, p0068$, p006A%) As Integer
Dim l006C As Integer
Dim l006E As Integer
  l006C = fn06CC(p0064, p0066)
  If l006C Then
    fn13C6 = l006C
    If gReloc_Arr(l006C).subRef_M3C6A = 0 Then
      l006E = MakeSubRef(l006C, p0068, p006A)
    End If
  End If
End Function

Function fn13D1(ByVal pv0074%, p0076&) As Integer
  fn13D1 = SomeAdrConv(fn0727(pv0074, p0076))
End Function

Function fn13E0(p007C%, ByVal p007E&) As Integer
  p007E = fn071D(p007C, p007E)
  fn13E0 = fn1624(p007E)
End Function

Function fn1483(p00F8%) As Long
  fn1483 = VAtoRaw(fn202F(p00F8))
End Function

Function fn1491(ByVal pv0102 As Long) As Integer
Dim l0104 As Integer
  l0104 = fn1624(pv0102)
  Do While l0104
    If gv161C(l0104).VOffset >= pv0102 Then
      If gv161C(l0104).VOffset = pv0102 Then
        Select Case gv161C(l0104).M36B3_count
        Case gc1166, gc1168
          l0104 = gv161C(l0104).subRef_M3C6A
          If l0104 <= 0 Then Exit Function
          If fn202F(l0104) <= 0 Then Exit Function
          fn1491 = l0104
        End Select
      End If
      Exit Function
    End If
    l0104 = gv161C(l0104).Count
  Loop
End Function

Function fn149E(ByVal pv0122 As Integer) As Integer
  pv0122 = gv161C(pv0122).Count
  If pv0122 = 0 Then Exit Function
  Select Case gv161C(pv0122).M36B3_count
  Case gc1166, gc1168
  Case gc1164
    sub00B9 gv161C(pv0122)
  End Select
  fn149E = pv0122
End Function

Sub sub13AA(p0058%, p005A&, p005C%)
  sub13B9 p0058, fn0727(gv106A, p005A), p005C
End Sub

Private Sub LoadSegmentTable()
Dim i
Dim SegmTbl_Entry As SegmTblEntrStruc
Dim SegName As String
Dim SegmentTableOffsetRaw As Long
Dim SegSize As Long
Dim SegOffsetVA As Long
Dim Sections() As Integer
 Const SegmentTable_Type = 6 ' &H6%

Dim SegType As Integer
Dim SegFlags As Integer
 
 Const SegTypeWriteableFlag = 1 ' &H1%
 Const SegTypeData = 9 ' &H9%
 Const SegTypeCode = 8 ' &H8%
 Const mc01A4 = 2 ' &H2%
  
  ReDim gSectionsForVB(gSegCount), Sections(gSegCount)
  
 'Seek to SegmentTable
  SegmentTableOffsetRaw = gMZ_Header.OffsetToPE + gNE_Header.SegmentTableOffset + 1
  Seek g_hFileExe, SegmentTableOffsetRaw
  
 'Log SegEntries
  AddToLogTree SegmentTable_Type, "SecEntries", SegmentTableOffsetRaw, 0, gSegCount * Len(SegmTbl_Entry), 1
  
  For i = 1 To gSegCount
   'Get SegmentTableEntry
    Get g_hFileExe, , SegmTbl_Entry
    
   'Calc and store Raw Begin Offset
    SegmentTableOffsetRaw = ToUnSigned(SegmTbl_Entry.SegOffset) * SegSize
    gSectionsForVB(i).RAW_Offset = SegmentTableOffsetRaw
    
   'Calc and store Segment Size
    If SegmTbl_Entry.size = 0 And SegmTbl_Entry.SegOffset Then
      SegSize = &H10000
    Else
      SegSize = ToUnSigned(SegmTbl_Entry.size)
    End If
    gSectionsForVB(i).Raw_Size = SegSize
   '... Segment Flags
    gSectionsForVB(i).Flags = SegmTbl_Entry.Flags
    
   '... Virtual Offset
    SegOffsetVA = i * &H10000
    gSectionsForVB(i).VA_Offset = SegOffsetVA
    
    If SegSize Then
   
   '... Segment Flags
      SegFlags = SegmTbl_Entry.Flags
      If SegFlags And SegTypeWriteableFlag Then
        SegType = SegTypeData
      Else
        SegType = SegTypeCode
        SegFlags = SegFlags Or mc01A4
      End If
     
    ' ...SegName
      SegName = " Seg" & Format$(i)
      StoreSectData SegOffsetVA, SegSize, SegmentTableOffsetRaw, SegName, SegFlags
      
     'Log Segment
      SegmentTableOffsetRaw = SegmentTableOffsetRaw + 1
      AddToLogTree SegType, SegName, SegmentTableOffsetRaw, SegOffsetVA, SegSize, 2
    
    End If
  Next i

End Sub

Private Sub sub141D()
Dim l0134 As Long
Dim l0136 As Long
Dim l0138 As Integer
Dim l013A As Integer
Dim l013C As Integer
Dim l013E As Integer
Dim l0142 As T5057
Dim l0146 As T5042
' Const mc0152 = 14 ' &HE%
  l0134 = gMZ_Header.OffsetToPE + gNE_Header.ENTRYTABLE + 1
  gv06D4 = 0
  Get g_hFileExe, l0134, gSingleChar: l0138 = Asc(gSingleChar)
  While l0138
    gv06D4 = gv06D4 + l0138
    Get g_hFileExe, , gSingleChar: l013A = Asc(gSingleChar)
    Select Case l013A
    Case 0
    Case &HFF
      While l0138
        Get g_hFileExe, , l0142
        l0138 = l0138 - 1
      Wend
    Case Else
      While l0138
        Get g_hFileExe, , l0146
        l0138 = l0138 - 1
      Wend
    End Select
    Get g_hFileExe, , gSingleChar: l0138 = Asc(gSingleChar)
  Wend
  l0136 = Seek(g_hFileExe)
  AddToLogTree mc0152, "EntryTable", (l0134), 0, l0136 - l0134, 1
  ReDim gv06A2(gv06D4)
  Get g_hFileExe, l0134, gSingleChar: l0138 = Asc(gSingleChar)
  l013C = 1
  While l0138
    Get g_hFileExe, , gSingleChar: l013A = Asc(gSingleChar)
    Select Case l013A
    Case 0
      l013C = l013C + l0138
    Case &HFF
      While l0138
        Get g_hFileExe, , l0142
        gv06A2(l013C).Flags = Asc(l0142.Flags)
        gv06A2(l013C).M3C63 = Asc(l0142.M506D)
        gv06A2(l013C).hFile = l0142.M504F
        If gv06A2(l013C).M3C63 <= gSegCount Then
          l013E = fn06CC(gv06A2(l013C).M3C63, l0142.M504F)
        End If
        l0138 = l0138 - 1
        l013C = l013C + 1
      Wend
    Case Else
      While l0138
        Get g_hFileExe, , l0146
        gv06A2(l013C).Flags = Asc(l0146.Flags)
        gv06A2(l013C).M3C63 = l013A
        gv06A2(l013C).hFile = l0146.M504F
        If l013A <= gSegCount Then
          l013E = fn06CC(l013A, (l0146.M504F))
        End If
        l0138 = l0138 - 1
        l013C = l013C + 1
      Wend
    End Select
    Get g_hFileExe, , gSingleChar: l0138 = Asc(gSingleChar)
  Wend
End Sub

Private Sub sub1433()
Dim l0178 As Long
Dim l017A As String
Dim l017C As Integer
Dim l017E As Integer
' Const mc0180 = 3 ' &H3%
' Const mc0182 = 18 ' &H12%
  l0178 = gMZ_Header.OffsetToPE + gNE_Header.ResidentNameTable + 1
  Get g_hFileExe, l0178, gSingleChar
  While Asc(gSingleChar)
    l017A = Space$(Asc(gSingleChar))
    Get g_hFileExe, , l017A
    l017C = InStr(l017A, Chr$(0))
    If l017C Then l017A = Left$(l017A, l017C - 1)
    Get g_hFileExe, , l017C
    If l017C > 0 Then
      l017E = fn13C6(gv06A2(l017C).M3C63, gv06A2(l017C).hFile, l017A, mc0180)
      gv06A2(l017C).subRef_M3C6A = l017E
    End If
    Get g_hFileExe, , gSingleChar
  Wend
  AddToLogTree mc0182, "Resident", (l0178), 0, Seek(g_hFileExe) - l0178, 2
End Sub

Private Sub sub1444()
Dim l0154 As Long
Dim l0156 As Long
Dim l0158%
Dim l015A%
' Const mc015E = 10 ' &HA%
' Const mc0160 = 11 ' &HB%
  Assume gNE_Header.ModuleTableEntryCount = gImportDLLCount, "mod cnt"
  Assume gModEndVA >= (gSegCount + gImportDLLCount) * &H10000, "mod limit"
  l0156 = gMZ_Header.OffsetToPE + gNE_Header.ModuleReferenceTable + 1
  l0154 = gMZ_Header.OffsetToPE + gNE_Header.IMPORTTABLE + 1
  AddToLogTree mc015E, "ModTable", (l0156), 0, gNE_Header.ModuleTableEntryCount * 2, 1
  AddToLogTree mc0160, "ImpTable", (l0154), 0, 0, 2
  ReDim gv04D6(gNE_Header.ModuleTableEntryCount)
  For l015A = 1 To gNE_Header.ModuleTableEntryCount
    Get g_hFileExe, l0156, l0158: l0156 = l0156 + 2
    Get g_hFileExe, l0154 + l0158, gSingleChar
    gv04D6(l015A).Name = Space$(Asc(gSingleChar))
    Get g_hFileExe, , gv04D6(l015A).Name
    sub067F l015A, gv04D6(l015A).Name & ".ept", gv04D6(l015A)
    If gv04D6(l015A).Name = "VB40016" Then
      gVBversion = &H416
    End If
  Next l015A
End Sub

Private Sub sub1455()
Dim l016C As Long
Dim l016E As String
Dim l0170 As Integer
Dim l0172 As Integer
' Const mc0174 = 4 ' &H4%
' Const mc0176 = 16 ' &H10%
  l016C = gNE_Header.NonResidentNameTable + 1
  Get g_hFileExe, l016C, gSingleChar
  While Asc(gSingleChar)
    l016E = Space$(Asc(gSingleChar))
    Get g_hFileExe, , l016E
    l0170 = InStr(l016E, Chr$(0))
    If l0170 Then l016E = Left$(l016E, l0170 - 1)
    Get g_hFileExe, , l0170
    If l0170 > 0 Then
      l0172 = fn13C6(gv06A2(l0170).M3C63, (gv06A2(l0170).hFile), l016E, mc0174)
      gv06A2(l0170).subRef_M3C6A = l0172
    End If
    Get g_hFileExe, , gSingleChar
  Wend
  AddToLogTree mc0176, "NonResident", (l016C), 0, Seek(g_hFileExe) - l016C, 2
End Sub

Private Sub sub1467()
Dim l00A0%
Dim l00A4() As T50ED
Dim l00AA%
Dim l00AC%
Dim l00B0%
Dim l00B2 As Long
Dim pat_sym As Integer
Dim l00B8 As T50ED
' Const mc00C6 = 256 ' &H100%
' Const mc00DA = 48 ' &H30%
  For l00A0 = 1 To gSegCount
    If gSectionsForVB(l00A0).Flags And mc00C6 Then
      sub0D4A l00A0
      If gv105A Then
        sub1476 l00A0, l00A4(), l00AA
        
        For l00AC = 1 To l00AA
          l00B8 = l00A4(l00AC)
          Select Case l00B8.M50D9 And 3
          Case m0026
            If l00B8.M50E1 = &HFF Then
              pat_sym = l00B8.M50E7
              l00B8.M50E1 = gv06A2(pat_sym).M3C63
              l00B8.M50E7 = gv06A2(pat_sym).hFile
              pat_sym = gv06A2(pat_sym).subRef_M3C6A
            Else
              pat_sym = 0
            End If
            If pat_sym = 0 Then
              pat_sym = fn06CC(l00B8.M50E1, ToUnSigned(l00B8.M50E7))
            End If
          Case mc0028
            pat_sym = fn06CC(-l00B8.M50E1, ToUnSigned(l00B8.M50E7))
          Case mc002A
            ErrBeep
            pat_sym = 0
          End Select
          
          DbgVar pat_sym, "pat sym"
          Select Case l00B8.M50F7
          Case gc1166
            If l00B8.M50D9 And mc002C Then
              l00B2 = ToUnSigned(l00B8.hFile)
              Get g_hFileExe, gv1056 + l00B2, l00B0
              pat_sym = fn06CC(l00B8.M50E1, ToUnSigned(l00B8.M50E7 + l00B0))
              sub13AA pat_sym, l00B2, l00B8.M50F7
            Else
              l00B0 = l00B8.hFile
              Do
                l00B2 = ToUnSigned(l00B0)
                sub13AA pat_sym, l00B2, l00B8.M50F7
                Get g_hFileExe, gv1056 + l00B2, l00B0
              Loop Until l00B0 = &HFFFF Or l00B2 = 0
            End If
            
          Case gc1164
            If l00B8.M50D9 And mc002C Then
              l00B2 = ToUnSigned(l00B8.hFile)
              Get g_hFileExe, gv1056 + l00B2, l00B0
              If l00B0 Then
                MsgBox gc0046, mc00DA
              End If
              pat_sym = fn06CC(l00B8.M50E1, ToUnSigned(l00B8.M50E7 + l00B0))
              sub13AA pat_sym, l00B2, l00B8.M50F7
            Else
              l00B0 = l00B8.hFile
              Do
                l00B2 = ToUnSigned(l00B0)
                sub13AA pat_sym, l00B2, l00B8.M50F7
                Get g_hFileExe, gv1056 + l00B2, l00B0
              Loop Until l00B0 = &HFFFF Or l00B2 = 0
            End If
          Case Else
            MsgBox Foundunknownfixups
          End Select
        Next
        
      End If
    End If
  Next
End Sub

Sub sub1476(p00DE As Integer, p00E2() As T50ED, p00E8%)
Dim l00EA
Dim l00EC As Long
Dim l00F0 As T50CD
Dim l00F4 As T50ED
  If gSectionsForVB(p00DE).Flags And mc00C6 Then
    l00EC = gSectionsForVB(p00DE).RAW_Offset + 1 + gSectionsForVB(p00DE).Raw_Size
    
    Get g_hFileExe, l00EC, p00E8
    
    If p00E8 < 0 Then
      p00E8 = 0
    End If
    
    ReDim p00E2(p00E8)
    For l00EA = 1 To p00E8
      Get g_hFileExe, , l00F0
      l00F4.hFile = l00F0.hFile
      l00F4.M50F7 = Asc(l00F0.lOffsetSym_M3C6A)
      l00F4.M50D9 = Asc(l00F0.M50D9)
      l00F4.M50E1 = l00F0.M50E1
      l00F4.M50E7 = l00F0.M50E7
      p00E2(l00EA) = l00F4
    Next l00EA
  Else
    Erase p00E2: p00E8 = 0
  End If
End Sub

