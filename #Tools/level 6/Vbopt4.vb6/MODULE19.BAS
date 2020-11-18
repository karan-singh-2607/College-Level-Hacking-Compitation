Attribute VB_Name = "MODULE19"
' Module19
Option Explicit
Const FlagContainExecutabeCode = 32 ' &H20%
Const FlagContainsInitData = 64 ' &H40%
Dim mFirstThunkVA As Long
Dim Type_OriginalFirstThunk As Long
Dim Exports As ExportTableStruct
Dim mExport_Funcs() As T53F5
Dim Export_NumOfFuncs As Integer
Dim mExport_StringBuff As String
Dim Export_Base As Integer
Dim Export_FuncCount As Integer
Const mc00AE = 32 ' &H20%
Const Type_ExportsAddrs = 15 ' &HF%
Const Type_ExportsNames = 16 ' &H10%
Const Type_ExportsOrdinals = 17 ' &H11%
Const Type_ExportsExpNames = 18 ' &H12%
'Const 2 = 2 ' &H2%
Const mc0116 = 4 ' &H4%
Const mc0118 = 5 ' &H5%
Const mc011A = 3 ' &H3%
Const mc011C = 6 ' &H6%
Const mc0122 = 8 ' &H8%
Const mc0124 = 2 ' &H2%
Const mc0126 = 256 ' &H100%
Const mc0128 = 9 ' &H9%
Const mc012A = 1 ' &H1%
Const mc012C = 7 ' &H7%
Const Type_Exports = 14 ' &HE%
Const Type_Imports = 10 ' &HA%
Const Type_Res = 20 ' &H14%
Const Type_Reloc = 24 ' &H18%
Const Type_Debug = 28 ' &H1C%
Const mc0172 = 1 ' &H1%
Const Type_RelocChunk = 25 ' &H19%
Const Type_ImportDLLName = 12 ' &HC%
Const Type_FirstThunk = 11 ' &HB%
Const mc01DE = 3 ' &H3%

Function fn15A9(p020C As Long) As Long
  fn15A9 = p020C + gPESect_Text_VToRaw
End Function

Function fn1561(p00DA As Long) As String
Dim l00DE As String * 64
Dim l00E0 As Integer
Dim l00E2 As Long
Dim l00E4 As String
Dim l00E6 As Integer
  l00E2 = Seek(g_hFileExe)
  Get g_hFileExe, gVAToRaw2 + p00DA, l00DE
  Seek g_hFileExe, l00E2
  If Asc(Mid$(l00DE, 2, 1)) Then
    l00E0 = InStr(l00DE, Chr$(0))
    l00E4 = Left$(l00DE, l00E0 - 1)
  Else
    For l00E0 = 1 To Len(l00DE) Step 2
      l00E6 = Asc(Mid$(l00DE, l00E0, 1))
      If l00E6 = 0 Then Exit For
      l00E4 = l00E4 & Chr$(l00E6)
    Next
  End If
  fn1561 = l00E4
End Function

Sub LoadPE()
Dim i
Dim RAW_Offset As Long
Dim OptionalHeaderSize As Long
Dim imgdir_type As Integer
Dim SectionName As String
Dim VA_Offset As Long
' Const 2 = 2 ' &H2%
' Const mc0116 = 4 ' &H4%
' Const mc0118 = 5 ' &H5%
' Const mc011A = 3 ' &H3%
' Const mc011C = 6 ' &H6%
Dim Sect_Flag As Integer
' Const mc0122 = 8 ' &H8%
' Const mc0124 = 2 ' &H2%
' Const mc0126 = 256 ' &H100%
' Const mc0128 = 9 ' &H9%
' Const mc012A = 1 ' &H1%
' Const mc012C = 7 ' &H7%
' Const Type_Exports = 14 ' &HE%
' Const Type_Imports = 10 ' &HA%
' Const Type_Res = 20 ' &H14%
' Const Type_Reloc = 24 ' &H18%
' Const Type_Debug = 28 ' &H1C%
  gSegCount = 0: ReDim gSectionsForVB(10)
  InitHeadersVars RAW_Offset
  
 'LoadPE PE header Coff part
  Get g_hFileExe, RAW_Offset, gPE_HeaderCoff
  
  AddToLogTree 2, "PE Hdr", RAW_Offset, 0, Len(gPE_HeaderCoff), 1
  
  OptionalHeaderSize = gPE_HeaderCoff.OptionalHeaderSize
  
 'LoadPE PE header optional
  If OptionalHeaderSize = &HE0 Then
    
    Get g_hFileExe, , gPE_OptHdr
    
    gImageBase1 = gPE_OptHdr.ImageBase
    
   'Fill PEDirs
    For i = 1 To 16
      Get g_hFileExe, , gPE_Dir(i)
    Next
    
    AddToLogTree mc0116, "NT Hdr", RAW_Offset, 0, Len(gPE_OptHdr), 1
    AddToLogTree mc0118, "Img Dir", RAW_Offset, 0, 8 * 16, 1
    
  Else
    AddToLogTree mc011A, "Opt Hdr", RAW_Offset, 0, OptionalHeaderSize, 1
  End If
  
  AddToLogTree mc011C, "SecHdrs", RAW_Offset, 0, gPESectionCount * &H28, 1
  
  gPESectionCount = gPE_HeaderCoff.NumSections:  ReDim gPESections(gPESectionCount)
  
  
'Read Sections
  For i = 1 To gPESectionCount
    
    Get g_hFileExe, , gPESections(i)
    
    SectionName = StrCutZero(gPESections(i).Name)
    
    LSet gPESections(i).Name = SectionName
    
    RAW_Offset = gPESections(i).RAW_Offset
    
    If RAW_Offset Then
      If gPESections(i).Flags And FlagContainExecutabeCode Then
        imgdir_type = 8 'mc0122
        Sect_Flag = 2 Or 256 'mc0124 Or mc0126
      ElseIf gPESections(i).Flags And FlagContainsInitData Then
        imgdir_type = 9 'mc0128
        Sect_Flag = 1 'mc012A
      Else
        imgdir_type = 7 'mc012C
        Sect_Flag = 0
      End If
      
      VA_Offset = gPESections(i).RVA_Offset + gImageBase1
      
     'Store Section Values
      StoreSectData VA_Offset, gPESections(i).Raw_Size, RAW_Offset, SectionName, Sect_Flag
      
     'inc because of VB Base 1
      RAW_Offset = RAW_Offset + 1
      
      If SectionName = ".text" Then
        gPESect_Text_VToRaw = RAW_Offset - VA_Offset
        gPESect_Text_VStart = VA_Offset
        gPESect_Text_VEnd = VA_Offset + gPESections(i).VSize
      End If
      
      AddToLogTree imgdir_type, SectionName, RAW_Offset, VA_Offset, gPESections(i).Raw_Size, 1
    End If
  Next
  
  'Is last RawSection bigger than the files?
  If RAW_Offset < LOF(g_hFileExe) Then AddToLogTree 0, "more", RAW_Offset, 0, LOF(g_hFileExe) - RAW_Offset, 0
  
  Assume gPESect_Text_VStart <> 0, "found .text"
  
 'Optional Section Entry AddToTree
  For i = 1 To gPE_HeaderCoff.NumSections
    
    If gPE_Dir(i).Size Then
      
      imgdir_type = 0

      Select Case i
      Case 1: SectionName = "Exports":    imgdir_type = Type_Exports
      Case 2: SectionName = "Imports":    imgdir_type = Type_Imports
      Case 3: SectionName = "Resource":   imgdir_type = Type_Res
      Case 4: SectionName = "Exception"
      Case 5: SectionName = "Security"
      Case 6: SectionName = "Relocation": imgdir_type = Type_Reloc
      Case 7: SectionName = "Debug":      imgdir_type = Type_Debug
      Case 8: SectionName = "Copyright"
      Case 9: SectionName = "GlobalPtr"
      Case 10: SectionName = "TLS"
      Case 11: SectionName = "LoadConfig"
      Case Else: SectionName = "Img" & Format$(i)
      End Select
      
      DebugPrintVar imgdir_type, "imgdir type"
      
      RVAtoRaw_And_VA gPE_Dir(i).RVA_Offset, RAW_Offset, VA_Offset
      
      
    ' Log
      OptionalHeaderSize = gPE_Dir(i).Size
      AddToLogTree imgdir_type, SectionName, RAW_Offset, VA_Offset, OptionalHeaderSize, 2
      
      If imgdir_type = Type_Debug Then
        AddToLogTree Type_Debug, "stripped", RAW_Offset, 0, 0, 2
      End If
      
    End If
  Next
  
  ProcessRelocs
  ProcessImports
  ProcessExports
  sub1510
  sub1502

End Sub

Private Sub ProcessImportDirTable(ImportDirectoryTable As ImportDirectoryTableStruc)
Dim DllName As String
Dim OffsetRaw As Long
Dim ThunkValue As Long
Dim OffsetVA     As Long
' Const Type_ImportDLLName = 12 ' &HC%
' Const Type_FirstThunk = 11 ' &HB%
Dim SymOffset As Integer
Dim l01D2 As Integer
Dim l01D4 As Integer
' Const mc01DE = 3 ' &H3%
  
  'Get DllName
   OffsetVA = ImportDirectoryTable.DllName + gImageBase1
   OffsetRaw = RVAtoRaw(ImportDirectoryTable.DllName)
   DllName = ReadZeroString(g_hFileExe, OffsetRaw)
  AddToLogTree Type_ImportDLLName, DllName & ".Hints", OffsetRaw, OffsetVA, 0, 2
  
  'detect VBVersion via Imported Dll's
  Select Case UCase$(Left$(DllName, 8))
   Case "VB40032.":   gVBversion = &H432:  gIsVB_Ver = 4
   Case "MSVBVM50":   gVBversion = 5:      gIsVB_Ver = 5
   Case Else
  End Select
  
 'Store DllName
  gImportDLLCount = gImportDLLCount + 1
  gImportDlls(gImportDLLCount).DllName = DllName
  
  
 'Load *.exp
  LoadExpFile gImportDLLCount
  
  'Get FirstThunk Offset
   OffsetVA = ImportDirectoryTable.FirstThunk + gImageBase1
   mFirstThunkVA = OffsetVA
   OffsetRaw = VAtoRaw(mFirstThunkVA)
  AddToLogTree Type_FirstThunk, DllName & ".addrs", OffsetRaw, OffsetVA, 0, 2
  
  'Get OriginalFirstThunk Offset
   OffsetVA = ImportDirectoryTable.OriginalFirstThunk + gImageBase1
   OffsetRaw = VAtoRaw(OffsetVA)
   Type_OriginalFirstThunk = OffsetVA
  AddToLogTree Type_FirstThunk, DllName & ".lookup", OffsetRaw, OffsetVA, 0, 2
  
  Seek g_hFileExe, OffsetRaw
  Do
    
    Get g_hFileExe, , ThunkValue
    
    'OriginalFirstThunk...
    l01D2 = CreateSym(OffsetVA)
    
    If ThunkValue = 0 Then Exit Do
   
    If ThunkValue < 0 Then
    ' Hint is Present
      SymOffset = gv21E6(ThunkValue And &H7FFF)
      DebugPrintVar SymOffset <> 0, "valid import"
      
      If SymOffset Then gReloc_Arr(l01D2).subRef_M3C6A = -SymOffset
    Else
     'Hint is Missing (0000)
      DllName = ReadZeroString(g_hFileExe, RVAtoRaw(ThunkValue + 2))
      
      SymOffset = importsRead(ThunkValue + gImageBase1, DllName)
      
      Assume SymOffset, "imports read"
      gReloc_Arr(SymOffset).M36B3_count = mc01DE
      gReloc_Arr(l01D2).subRef_M3C6A = -SymOffset
    End If
    
    l01D4 = CreateSym(mFirstThunkVA + (OffsetVA - Type_OriginalFirstThunk))
    
    Assume l01D4, "addr rva"
    
    gReloc_Arr(l01D4).subRef_M3C6A = -SymOffset
    OffsetVA = OffsetVA + 4
  Loop
  
  Assume gImportDlls(gImportDLLCount).ImageBase <> 0, "imgbase"
End Sub

Private Sub LoadExpFile(ImportDllIndex As Integer)
Dim i As Integer
Dim ExpFileName As String
Dim hFile As Integer
Dim First%
Dim Last%
Dim ImageBase As Long
Dim l0204 As T5402
Dim lPE_OptHdr As PE_OptHdr
  ExpFileName = gImportDlls(ImportDllIndex).DllName
  i = InStr(ExpFileName, ".")
  
 'Test if Exp file was found
  ExpFileName = g_AppPath & Left$(ExpFileName, i) & "exp"
  If Len(Dir$(ExpFileName)) = 0 Then
    MsgBox ExpFileName, , "missing export file"
    Exit Sub
  End If
  
  
 'open *.Exp
  hFile = FreeFile
  Open ExpFileName For Binary As hFile
  Get hFile, , i
  If i = mc00AE Then
    Get hFile, , lPE_OptHdr
    
    Get hFile, , First
    Get hFile, , Last
    
    ImageBase = lPE_OptHdr.ImageBase
    
    sub140D ImportDllIndex, ImageBase, ImageBase + lPE_OptHdr.SizeOfImage
    For i = First To Last
      Get hFile, , l0204
      If l0204.RVA_Offset = 0 Then Exit For
      sub068C ImportDllIndex, l0204
    Next
  
  End If
  
  Close hFile
End Sub

Private Sub sub1502()
Dim l01E0 As Long
Dim l01E2 As Long
Dim l01E4 As Long
Dim l01E8 As Long
Dim l01EC As T541B
  l01E2 = gPE_Dir(gcDebug).RVA_Offset
  l01E4 = gPE_Dir(gcDebug).Size
  If l01E4 = 0 Then Exit Sub
  l01E0 = RVAtoRaw(l01E2)
  l01E8 = l01E0 + l01E4
  While l01E0 < l01E8
    Get g_hFileExe, l01E0, l01EC: l01E0 = l01E0 + Len(l01EC)
    AddToLogTree mc012C, "debug", l01EC.RAW_Offset + 1, l01E2 + gImageBase1, l01EC.Size, 1
  Wend
End Sub

Private Sub sub1510()
Dim l0194 As Long
Dim l0196 As Long
  InitResObj
  l0194 = gPE_Dir(gcResource).RVA_Offset
  If l0194 = 0 Then Exit Sub
  l0196 = RVAtoRaw(l0194)
  gv1598 = l0196
  gv159C = l0194 + gImageBase1
  sub151D "", 1, l0196, 0
End Sub

Sub RVAtoRaw_And_VA(RVA_Offset As Long, RAW_Offset As Long, VA_Offset As Long)
  RAW_Offset = RVAtoRaw(RVA_Offset)
  VA_Offset = RVA_Offset + gImageBase1
End Sub

Private Sub ExportLoadNames()
Dim i
Dim AddrsRaw As Long
Dim Addrs As Long
Dim AddressOfNames As Long
Dim Addrs2 As Long
Dim ExportsName() As Integer
  Export_NumOfFuncs = Exports.NumberOfFunctions
  
 'Dim Array for Exports
  Export_Base = Exports.Base
  Export_FuncCount = Export_Base + Export_NumOfFuncs - 1
  ReDim mExport_Funcs(Export_Base To Export_FuncCount)
  
  Addrs = Exports.Addrs
  
  AddrsRaw = RVAtoRaw(Addrs)
  Seek g_hFileExe, AddrsRaw
  
 'Read StringPointer
  For i = Export_Base To Export_FuncCount
    Get g_hFileExe, , mExport_Funcs(i).RVA_Offset
  Next
  
 'GetHints
  ReDim ExportsName(Exports.NumberOfNames) As Integer
  Seek g_hFileExe, RVAtoRaw(Exports.Ordinals)
  For i = 1 To Exports.NumberOfNames
    Get g_hFileExe, , ExportsName(i)
  Next
  
  AddressOfNames = Exports.DllName - 1
  
 'Get StringsOffsets
  Seek g_hFileExe, RVAtoRaw(Exports.Names)
  For i = 1 To Exports.NumberOfNames
    Get g_hFileExe, , Addrs
    If Addrs > Addrs2 Then Addrs2 = Addrs
    mExport_Funcs(ExportsName(i) + Export_Base).StringBuffOffs = Addrs - AddressOfNames
  Next
  
  Erase ExportsName
  
 'GetStrings
  mExport_StringBuff = Space$(Addrs2 - AddressOfNames + 30)
  Get g_hFileExe, RVAtoRaw(Exports.DllName), mExport_StringBuff
  
End Sub

Sub sub1541()
Dim l00B4 As Integer
Dim l00B6 As String
Dim l00BA
Dim l00BC As Integer
Dim l00BE As Integer
  l00B6 = g_AppPath & gExeFileNameOnly2 & ".exd"
  l00B4 = FreeFile
  Open l00B6 For Output As l00B4
  Print #l00B4, "Exports "; G_FullExeFileName
  For l00BA = Export_Base To Export_FuncCount
    l00BC = mExport_Funcs(l00BA).StringBuffOffs
    If l00BC Then
      l00BE = InStr(l00BC, mExport_StringBuff, Chr$(0))
      Print #l00B4, l00BA; Tab(7); fn086D(mExport_Funcs(l00BA).RVA_Offset); " "; Mid$(mExport_StringBuff, l00BC, l00BE - l00BC)
    End If
  Next
  Close l00B4
End Sub

Sub sub1551(p00CA As ListBox)
Dim l00CC
Dim l00CE As Integer
Dim l00D0 As Integer
Dim l00D2 As String
  l00D2 = Format$(Export_FuncCount)
  SetDecompilerListBox p00CA
  For l00CC = Export_Base To Export_FuncCount
    l00D0 = mExport_Funcs(l00CC).StringBuffOffs
    If l00D0 Or mExport_Funcs(l00CC).RVA_Offset Then
      RSet l00D2 = Format$(l00CC): DisOutput l00D2
      DisOutput fn0853(mExport_Funcs(l00CC).RVA_Offset + gImageBase1)
      If l00D0 Then
        l00CE = InStr(l00D0, mExport_StringBuff, Chr$(0))
        DisOutput " " & Mid$(mExport_StringBuff, l00D0, l00CE - l00D0)
      End If
      WriteOutputTxt
    End If
  Next
  UnlinkDecompilerListBox
End Sub

Private Sub ProcessExports()
Dim ExportsRaw As Long
Dim ExportsRVA As Long
Dim ExportsVA&
' Const Type_ExportsAddrs = 15 ' &HF%
' Const Type_ExportsNames = 16 ' &H10%
' Const Type_ExportsOrdinals = 17 ' &H11%
' Const Type_ExportsExpNames = 18 ' &H12%
  Export_Base = 1: Export_FuncCount = 0: Export_NumOfFuncs = 0
  
 'Exit if ExportDir in PE-Header points to 0000000
  ExportsRVA = gPE_Dir(gcExports).RVA_Offset
  If ExportsRVA = 0 Then Exit Sub
  
 'Read Export Struct
  ExportsRaw = RVAtoRaw(ExportsRVA)
  Get g_hFileExe, ExportsRaw, Exports
  
  RVAtoRaw_And_VA Exports.Addrs, ExportsRaw, ExportsVA
  AddToLogTree Type_ExportsAddrs, "Addrs", ExportsRaw, ExportsVA, 0, 3
  
  RVAtoRaw_And_VA Exports.Names, ExportsRaw, ExportsVA
  AddToLogTree Type_ExportsNames, "Names", ExportsRaw, ExportsVA, 0, 3
  
  RVAtoRaw_And_VA Exports.Ordinals, ExportsRaw, ExportsVA
  AddToLogTree Type_ExportsOrdinals, "Ordinals", ExportsRaw, ExportsVA, 0, 3
  
  RVAtoRaw_And_VA Exports.DllName, ExportsRaw, ExportsVA
  AddToLogTree Type_ExportsExpNames, "ExpNames", ExportsRaw, ExportsVA, 0, 3
  
  ExportLoadNames
  
End Sub

Private Sub ProcessRelocs()
Dim Reloc_Start As Long
Dim SomeSize As Long
Dim Reloc_Start_Raw As Long
Dim Reloc_End_Raw As Long
Dim RelocChunkEntries As Integer
Dim l0154&
Dim Reloc_Start_VA As Long
Dim l015A As Integer
Dim lImageBaseVA As Long
Dim lImageBase As Long
Dim NumEntrys As Integer
Dim l0164 As ImportDirectoryTableStruc
' Const mc0172 = 1 ' &H1%
' Const Type_RelocChunk = 25 ' &H19%
Dim l017A As Integer
  gImportDLLCount = 0
  SomeSize = gPE_Dir(gcImports).Size
  
  NumEntrys = (SomeSize \ Len(l0164))
  
 'Add also own module data to import list
  ReDim gImportDlls(NumEntrys)
  gImportDlls(0).DllName = gv_ExeFileName
  gImportDlls(0).ImageBase = gImageBase1
  gImportDlls(0).VOffset = gImageBase1 + gPE_OptHdr.SizeOfImage
  
  PrepareMemPages 0, 0, 0
  
  lImageBase = gImageBase1
  lImageBaseVA = lImageBase + gPE_OptHdr.AddressOfEntryPoint
  gEntryPoint_Page = RetOffsetSym(lImageBaseVA, "WinEntry", mc0172)
  
  Reloc_Start = gPE_Dir(gcRelocation).RVA_Offset
  If Reloc_Start = 0 Then Exit Sub
  
  Reloc_Start_VA = Reloc_Start + gImageBase1
  
  Reloc_Start_Raw = RVAtoRaw(Reloc_Start)
  Reloc_End_Raw = Reloc_Start_Raw + gPE_Dir(gcRelocation).Size
  
  Do While Reloc_Start_Raw < Reloc_End_Raw
   'RelocChunk_Base
    Get g_hFileExe, Reloc_Start_Raw, Reloc_Start
   'RelocChunk_Size
    Get g_hFileExe, , SomeSize
    RelocChunkEntries = (SomeSize \ 2) - 4
    If RelocChunkEntries > 0 Then
      
      ApplyRelocChunkEntries Reloc_Start + lImageBase, RelocChunkEntries
      
      AddToLogTree Type_RelocChunk, Hex$(Reloc_Start), Reloc_Start_Raw, Reloc_Start_VA, SomeSize, 2
      
      Reloc_Start_VA = Reloc_Start_VA + SomeSize
    Else
      Reloc_Start_Raw = Reloc_Start_Raw + SomeSize
    End If
  Loop
  
  l017A = fn1624(lImageBaseVA)
  l017A = fn165B(l017A)
  Reloc_Start_VA = gv161C(l017A).VOffset + 4
  Reloc_Start_Raw = VAtoRaw(Reloc_Start_VA)
  
  If Reloc_Start_Raw Then
    Do
      Reloc_Start_Raw = Reloc_Start_Raw - 6
      Reloc_Start_VA = Reloc_Start_VA - 6
      Get g_hFileExe, Reloc_Start_Raw, l0154
      If l0154 <> &H25FF Then Exit Do
      Get g_hFileExe, , SomeSize
      l015A = CreateSym(Reloc_Start_VA)
      gReloc_Arr(l015A).subRef_M3C6A = -CreateSym(SomeSize)
    Loop
  End If
End Sub

Private Sub ProcessImports()
Dim tmp As Long
Dim ImportTableAdrRaw As Long
Dim ImportTableAdr As Long
Dim ImportDirectoryTableStruc As ImportDirectoryTableStruc
  
 'Get Importtable Offset from PE-Header
  ImportTableAdr = gPE_Dir(gcImports).RVA_Offset
  If ImportTableAdr = 0 Then Exit Sub
  
 ' ProcessImportDirTable Entries (Dlls)
  tmp = RVAtoRaw(ImportTableAdr)
  Do
    
    ImportTableAdrRaw = tmp
    Get g_hFileExe, ImportTableAdrRaw, ImportDirectoryTableStruc
    tmp = Seek(g_hFileExe)
    
  ' Finished if Reach zero Entry
    If ImportDirectoryTableStruc.OriginalFirstThunk = 0 Then Exit Do
    
    ProcessImportDirTable ImportDirectoryTableStruc
    
  Loop
  
End Sub

