Attribute VB_Name = "MODULE34"
' Module34
Option Explicit
Dim MemPagesCreated As Integer
Const m_0x0100 = 256& ' &H100&
Const mc002E = 4095& ' &HFFF&
Dim gSectionMemAsPages() As Long
Dim PageBlock() As Integer
Dim mVAdressMapped As Integer
Dim mVAdressAsPage As Long
Dim mMapAdr2Result As Integer
Dim m00B4 As Integer
Const mc00D6 = 1 ' &H1%
Const mc0156 = 32 ' &H20%
Const mc0158 = 9 ' &H9%
Const mc015A = 13 ' &HD%
Const mc0168 = 39 ' &H27%
Const mc01AE = 2 ' &H2%
Const mc01B0 = 4 ' &H4%
Const mc01B2 = 3 ' &H3%
Const mc01BA = 41 ' &H29%
Const mc01CA = 38 ' &H26%

Function CreateSym(ByVal AdrThatNeedsRebase As Long) As Integer
Dim l00CE As Integer
Dim l00D0 As Integer
Dim AdrConved As Integer
Dim l00D4 As Long
' Const mc00D6 = 1 ' &H1%
  
  AdrConved = SomeAdrConv(AdrThatNeedsRebase)
  
  If mMapAdr2Result = 0 Then ErrBeep: Exit Function
  
  l00D0 = gReloc_Arr(AdrConved).Count
  
  If gReloc_Arr(AdrConved).Offset = AdrThatNeedsRebase Then
    CreateSym = AdrConved
    Exit Function
  End If
  
   Do While gReloc_Arr(AdrConved).Offset >= AdrThatNeedsRebase
     l00D0 = AdrConved
     AdrConved = gReloc_Arr(l00D0).M54BB
     If AdrConved = 0 Then
       If gReloc_Arr(l00D0).Offset = AdrThatNeedsRebase Then
         CreateSym = l00D0
         Exit Function
       End If
       Exit Do
     End If
   Loop
  
  Do While l00D0
    l00D4 = gReloc_Arr(l00D0).Offset
    If l00D4 >= AdrThatNeedsRebase Then
      If l00D4 = AdrThatNeedsRebase Then
        CreateSym = l00D0
        Exit Function
      End If
      Exit Do
    End If
    AdrConved = l00D0
    l00D0 = gReloc_Arr(l00D0).Count
  Loop
  
  m00B4 = True
  gReloc_ArrCount = gReloc_ArrCount + 1
  l00CE = gReloc_ArrCount
  If l00CE > UBound(gReloc_Arr) Then ReDim Preserve gReloc_Arr(l00CE + 100)
  gReloc_Arr(l00CE).Offset = AdrThatNeedsRebase
  gReloc_Arr(l00CE).Count = l00D0
  gReloc_Arr(l00CE).M54BB = AdrConved
  If l00D0 Then
    gReloc_Arr(l00D0).M54BB = l00CE
  End If
  
  If AdrConved Then
    gReloc_Arr(AdrConved).Count = l00CE
  End If
  
  gReloc_Arr(l00CE).M36B3_count = mc00D6
  AdrConved = PageBlock(mMapAdr2Result)
  If AdrConved Then
    
    If gReloc_Arr(AdrConved).Offset > AdrThatNeedsRebase Then
      Do
        PageBlock(mMapAdr2Result) = l00CE
        mMapAdr2Result = mMapAdr2Result - 1
      Loop While PageBlock(mMapAdr2Result) = AdrConved
    End If
    
  Else
    PageBlock(mMapAdr2Result) = l00CE
  End If
  
  CreateSym = l00CE
End Function

Function RetOffsetSym(VA As Long, Name$, TypeCode%) As Integer
Dim Dummy As Integer
Dim OffsetSym As Integer
  
  Assume VA > 0, "va>0"
  
  OffsetSym = CreateSym(VA)
  Assume OffsetSym, "sym created"
  
  Dummy = MakeSubRef(OffsetSym, Name, TypeCode)
  
  RetOffsetSym = OffsetSym

End Function

Function SomeAdrConv(Adress As Long) As Integer
Dim mPageBlockResult As Integer
  mMapAdr2Result = MapAdr2(Adress)
  
  If mMapAdr2Result = 0 Then ErrBeep: Exit Function
  
  mPageBlockResult = PageBlock(mMapAdr2Result)
  
  Do While mPageBlockResult
    SomeAdrConv = mPageBlockResult
    mPageBlockResult = gReloc_Arr(mPageBlockResult).Count
    If mPageBlockResult = 0 Then Exit Function
    If gReloc_Arr(mPageBlockResult).Offset > Adress Then Exit Do
  Loop
  
  While gReloc_Arr(mPageBlockResult).Offset > Adress
    mPageBlockResult = gReloc_Arr(mPageBlockResult).M54BB
  Wend
  
  SomeAdrConv = mPageBlockResult

End Function

Function fn2017(p0144 As Integer) As Integer
Dim l0146 As Long
  l0146 = gReloc_Arr(p0144).Offset
  fn2017 = l0146 < gImageBase2 Or l0146 >= gVOffset1
End Function

Function fn2023(p014C As Integer) As String
Dim l014E As Integer
Dim l0150 As String
Dim l0152 As Integer
Dim l0154 As Integer
' Const mc0156 = 32 ' &H20%
' Const mc0158 = 9 ' &H9%
' Const mc015A = 13 ' &HD%
' Const mc0168 = 39 ' &H27%
  If p014C = 0 Then
    MsgBox "loc 0?"
    fn2023 = "<nil>"
    Exit Function
  ElseIf p014C > 0 Then
    l0154 = gReloc_Arr(p014C).subRef_M3C6A
    While l0154 < 0
      l0154 = gReloc_Arr(-l0154).subRef_M3C6A
    Wend
  Else
    l0154 = -p014C
  End If
  If l0154 > 0 Then
    If gEXEdef_Arr4_imports(l0154).TypeCode_M5F96 < 0 Then
      Select Case gReloc_Arr(gEXEdef_Arr4_imports(l0154).lOffsetSym_M3C6A).M36B3_count
      Case mc0156, mc0158, mc015A
        If gv0B08 Then
          l0152 = -gEXEdef_Arr4_imports(l0154).TypeCode_M5F96
          l0152 = gEXEdef_Subs(l0152).M4289
          If l0152 Then
            fn2023 = fn0972(l0152)
            Exit Function
          End If
        End If
      Case mc0168
      Case Else
        DbgMileStone "more table types"
      End Select
    End If
    fn2023 = gEXEdef_Arr4_imports(l0154).Name
  Else
If g_IsDemoVersion Then
    fn2023 = fn0226(p014C)
Else
    l0152 = fn087D(gReloc_Arr(p014C).Offset)
    l014E = fn088F(gReloc_Arr(p014C).Offset)
If gIsVB_Ver Then
    If l0152 >= 0 Then
      l0150 = Hex$(l0152) & ":" & toHex16(l014E)
    Else
      MsgBox "unknown impname"
      l0150 = gv04D6(-l0152).Name & "." & Format$(l014E)
    End If
Else
    If l0152 > gSegCount Then
      ErrBeep
      MsgBox "unknown impname"
      l0150 = gv04D6(l0152 - gSegCount).Name & "." & Format$(l014E)
    ElseIf l0152 >= 0 Then
      l0150 = Hex$(l0152) & ":" & toHex16(l014E)
    Else
      MsgBox "unknown impname"
      l0150 = gv04D6(-l0152).Name & "." & Format$(l014E)
    End If
End If
    fn2023 = l0150
End If
  End If
End Function

Function fn202F(p0180 As Integer) As Long
  fn202F = gReloc_Arr(p0180).Offset
End Function

Function fn203B(p0184 As Integer) As Integer
Dim l0186 As Integer
  l0186 = gReloc_Arr(p0184).Count
  Assume l0186, "next loc"
  If l0186 Then
    fn203B = gReloc_Arr(l0186).Offset - gReloc_Arr(p0184).Offset
  End If
End Function

Function fn2047(p018A As Integer) As Integer
Dim l018C As Integer
  l018C = gReloc_Arr(p018A).subRef_M3C6A
  While l018C < 0
    l018C = gReloc_Arr(-l018C).subRef_M3C6A
  Wend
  fn2047 = l018C
End Function

Function fn2052(p0190 As Integer) As Integer
Dim l0192 As Integer
  l0192 = p0190
  Do
    l0192 = gReloc_Arr(l0192).Count
  Loop Until True
  fn2052 = l0192
End Function

Function MapModule(VA As Long) As Integer
  MapModule = VA \ gModulPageSize
End Function

Function fn206B(p019A As Integer, p019C As Integer) As Integer
Dim l019E As Long
Dim OutputString As String
Dim l01A2 As Integer
Dim IntBuff%
Dim l01A6 As Long
Dim l01A8 As Integer
Dim l01AA As Integer
Dim symbolRef As Integer
' Const mc01AE = 2 ' &H2%
' Const mc01B0 = 4 ' &H4%
' Const mc01B2 = 3 ' &H3%
' Const mc01BA = 41 ' &H29%
' Const mc01CA = 38 ' &H26%
  If p019C = 0 Then Exit Function
  
  Select Case p019C
  Case mc01AE
If gIsVB_Ver Then
    p019C = mc01B0
Else
    p019C = mc01B2
End If
  End Select
  
  If gReloc_Arr(p019A).M36B3_count = p019C Then Exit Function
  
  l019E = gReloc_Arr(p019A).Offset
  l01A6 = Seek(g_hFileExe)

L43B92:
  Select Case p019C
  Case mc01AE
If gIsVB_Ver Then
    p019C = mc01B0
Else
    p019C = mc01B2
End If
    GoTo L43B92
  Case mc01B0
    OutputString = """" & fn1561(l019E) & """"
  Case mc01B2
    OutputString = """" & ReadZeroString(g_hFileExe, l019E) & """"
  Case mc01BA
    Get g_hFileExe, fn1483(p019A), IntBuff
    OutputString = OutputString & " [" & toHex16(IntBuff) & ", "
    Get g_hFileExe, , IntBuff: OutputString = OutputString & toHex16(IntBuff) & ", "
    l01AA = fn1624(l019E + 4)
    If l01AA And gv161C(l01AA).VOffset = l019E + 4 Then
      symbolRef = gv161C(l01AA).subRef_M3C6A
      Assume symbolRef, "symbol reference"
      OutputString = OutputString & fn2023(symbolRef)
    Else
      Get g_hFileExe, , gv0120
      If gv0120 Then
        OutputString = OutputString & Hex$(gv0120) & "?"
      Else
        OutputString = OutputString & "NULL"
      End If
    End If
    OutputString = OutputString & "]"
    
  Case mc01CA
    DbgMileStone "name loc?"
    Get g_hFileExe, fn1483(p019A), gv0B84
    l01A8 = fn124C(gv0B84, False)
    If l01A8 = 0 Then
      l01A8 = Asc(Left$(gv0B84.M3CE1, 1))
      If l01A8 Then
        Mid$(gv0B84.M3CE1, 1, 1) = Chr$(l01A8 - 1)
        l01A8 = fn124C(gv0B84, True)
        Assume l01A8 > 0, "class defined"
      Else
        If gv0B84.M3CE1 <> String$(16, 0) Then
          DbgMileStone fn1285(gv0B84) & " class?"
        End If
      End If
    End If
    If l01A8 Then
      OutputString = fn120C(l01A8)
      symbolRef = l01A8
    End If
  End Select
  gReloc_Arr(p019A).M36B3_count = p019C
  
  l01A2 = fn2047(p019A)
  If Len(OutputString) Then
    If l01A2 Then
      If gEXEdef_Arr4_imports(l01A2).Name = OutputString Then Exit Function
    End If
    
    l01A2 = MakeSubRef(p019A, OutputString, symbolRef)
  
  End If
  Seek g_hFileExe, l01A6
  fn206B = l01A2
End Function

Function MapAdr2(VAdress As Long) As Integer

On Error GoTo Err_Map
  mVAdressMapped = VAdress \ gv20FA_Problem
  mVAdressAsPage = gSectionMemAsPages(mVAdressMapped)
  
  Assume mVAdressAsPage <> 0, "va mapped"
  MapAdr2 = (VAdress + mVAdressAsPage) \ m_0x0100

Exit Function

Err_Map:
  MsgBox Hex$(VAdress), , "Ill. addr"
  mVAdressAsPage = 0
Exit Function

End Function

Function MapAdr(p01DC As Long) As Integer
  MapAdr = p01DC \ gv20FA_Problem
End Function

Function fn208C(p01E0 As Long) As Integer
Dim l01E2
If gIsVB_Ver Then
  For l01E2 = gSegCount To 1 Step -1
    If p01E0 >= gSectionsForVB(l01E2).VA_Offset Then Exit For
  Next
  fn208C = l01E2
Else
  fn208C = p01E0 \ &H10000
End If
End Function

Sub PrepareMemPages(ImportDllIndex As Integer, VA_Start As Long, VA_End As Long)
Dim i
Dim AdrAsPage As Integer
Dim StartPage As Integer
Dim EndPage As Integer
Dim PageBlock_uBound As Integer
  
  'FirstCall
  If (ImportDllIndex = 0) And (VA_End = 0) Then
    gReloc_ArrCount = 0: ReDim gReloc_Arr(0)
    
   'Set Base
    gImageBase1 = gImportDlls(0).ImageBase
    gImageBase2 = gImageBase1
    gModStartVA = gImageBase1
    
   'Set VOffset
    gVOffset1 = gImportDlls(0).VOffset
    gModEndVA = gVOffset1

    '
    '!!!!!!
    gv20FA_Problem = &H10000
    StartPage = MapAdr(gModStartVA)
    EndPage = MapAdr(gModEndVA - 1)
    ReDim gSectionMemAsPages(StartPage To EndPage)
    
    ReDim PageBlock(0)
    gEXEdef_Arr4_importsCount = 0: ReDim gEXEdef_Arr4_imports(0)
    
   '!!!!
    gModulPageSize = &H10000
   If gIsVB_Ver Then
       
       MemPrepareSth
       For i = 1 To gPESectionCount
         
         'VA_End gets Size
         VA_End = gPESections(i).VSize
         VA_Start = gPESections(i).RVA_Offset + gImageBase1
         
         'if Size(=OffsetVA) bigger 0 then...
         If VA_End Then
         
           PrepareMemPages 0, VA_Start, VA_Start + VA_End
         Else
          'Section has no size
           Assume VA_End, "real section"
         End If
         
       Next
       

   Else
       
       gImageBase1 = 0
       
       MemPrepareSth
       
       If gPECurSection Then
         
         For i = 1 To gPECurSection
           VA_End = gPESections3(i).Raw_Size
           If VA_End Then
             VA_Start = gPESections3(i).VOffset
             
             PrepareMemPages 0, VA_Start, VA_Start + VA_End
             
           End If
         Next
       
       End If
   End If
  
  
'Other
  Else
    MemPagesCreated = False
    
   'Extend Modul Memory if needed
    If VA_Start < gModStartVA Then gModStartVA = VA_Start
    If VA_End > gModEndVA Then gModEndVA = VA_End
    
   'Module
    StartPage = MapAdr(gModStartVA)
    EndPage = MapAdr(gModEndVA - 1)
    
    ReDim Preserve gSectionMemAsPages(StartPage To EndPage)
    
   'Section
    StartPage = MapAdr(VA_Start)
    EndPage = MapAdr(VA_End - 1)
    
    For i = StartPage To EndPage
    
      If gSectionMemAsPages(i) = 0 Then
        PageBlock_uBound = UBound(PageBlock)
        ReDim Preserve PageBlock(PageBlock_uBound + gv20FA_Problem \ m_0x0100)
        
        PageBlock_uBound = PageBlock_uBound + 1
        mVAdressAsPage = PageBlock_uBound * m_0x0100 - i * gv20FA_Problem
        gSectionMemAsPages(i) = mVAdressAsPage
      End If
      
    Next
    
    AdrAsPage = CreateSym(VA_End - 1)
    gReloc_Arr(AdrAsPage).Offset = VA_End
    
    If gReloc_Arr(AdrAsPage).Count Then
      gReloc_Arr(gReloc_Arr(AdrAsPage).Count).M54BB = 0
      gReloc_Arr(AdrAsPage).Count = 0
    End If
    
    gReloc_Arr(AdrAsPage).M36B3_count = -ImportDllIndex
    For i = MapAdr2(VA_Start) To MapAdr2(VA_End - 1)
      PageBlock(i) = AdrAsPage
    Next
    
    AdrAsPage = CreateSym(VA_Start + 1)
    gReloc_Arr(AdrAsPage).Offset = VA_Start
    
    If gReloc_Arr(AdrAsPage).M54BB Then
      gReloc_Arr(gReloc_Arr(AdrAsPage).M54BB).Count = 0
      gReloc_Arr(AdrAsPage).M54BB = 0
    End If
    
    MemPagesCreated = True
  
  End If
End Sub

Sub sub140D(ImportDllIndex As Integer, ImageBase As Long, VOffset As Long)
  
  gImportDlls(ImportDllIndex).ImageBase = ImageBase
  gImportDlls(ImportDllIndex).VOffset = VOffset
  
  PrepareMemPages ImportDllIndex, ImageBase, VOffset
  
  gv21E2 = ImportDllIndex
  gv21E4 = 0
  
  ReDim gv21E6(100)
End Sub

Sub sub1FDE(p00E6 As Long, p00E8 As Long, p00EC As Control)
Dim l00EE As Integer
Dim l00F0 As Long
  p00EC.Clear
  p00EC.ListIndex = -1
  If p00E8 = 0 Then p00E8 = p00E6 + gv2102
  l00EE = SomeAdrConv(p00E6)
  If gReloc_Arr(l00EE).Offset < p00E6 Then l00EE = gReloc_Arr(l00EE).Count
  Do While l00EE
    l00F0 = gReloc_Arr(l00EE).Offset
    If l00F0 >= p00E8 Then Exit Do
    If gReloc_Arr(l00EE).M36B3_count < 0 Then Exit Do
    p00EC.AddItem fn086D(l00F0) & " " & fn0DD8(gReloc_Arr(l00EE).M36B3_count) & " " & fn2023(l00EE)
    p00EC.ItemData(p00EC.NewIndex) = l00EE
    l00EE = gReloc_Arr(l00EE).Count
  Loop
End Sub

Sub sub1FEB(p00F8 As Control)
Dim l00FA
Dim l00FC As Long
Dim l00FE As String
  p00F8.Clear
  p00F8.ListIndex = -1
  For l00FA = 1 To gPECurSection
    If gPESections3(l00FA).Raw_Size Then
      l00FE = gPESections3(l00FA).Name
      l00FC = gPESections3(l00FA).VOffset
      For l00FC = l00FC To l00FC + gPESections3(l00FA).Raw_Size - 1 Step gv2102
        p00F8.AddItem fn086D(l00FC) & l00FE
        p00F8.ItemData(p00F8.NewIndex) = l00FC
        l00FC = l00FC And Not mc002E
      Next
    Else
      MsgBox "lva=0 - remove!"
    End If
  Next
  For l00FA = 1 To gImportDLLCount
    p00F8.AddItem fn086D(gImportDlls(l00FA).ImageBase) & " " & gImportDlls(l00FA).DllName
    p00F8.ItemData(p00F8.NewIndex) = -l00FA
  Next
End Sub

Sub sub1FFB(p0114 As Integer, p0118 As Control)
  sub1FDE gImportDlls(p0114).ImageBase, gImportDlls(p0114).VOffset, p0118
End Sub

