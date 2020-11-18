Attribute VB_Name = "MODULE17"
' Module17
Option Explicit

Function RVAtoRaw(RVA_Offset As Long) As Long
  RVAtoRaw = VAtoRaw(RVA_Offset + gImageBase1)
End Function

Function VAtoRaw(VA_Offset As Long) As Long
Dim i

'For PE-Files
If gIsVB_Ver Then
  
 'Find Section into which VA_Offset points to
  For i = gPECurSection To 1 Step -1
    If gPESections3(i).VOffset <= VA_Offset Then
    
      VAtoRaw = (VA_Offset - gPESections3(i).VOffset) + gPESections3(i).RAW_Offset
      
      Exit Function
    End If
  Next
  
 'Error no Section found
  ErrBeep

Else
  i = VA_Offset \ &H10000
  If i = 0 Then
    ErrBeep
    VAtoRaw = VA_Offset
  Else
    VAtoRaw = gSectionsForVB(i).RAW_Offset + (VA_Offset And &HFFFF&) + 1
  End If

End If

End Function

Sub sub0115(p00EA&)
  sub0D4A fn208C(p00EA)
End Sub

Sub sub0D4A(p00E8 As Integer)
  If p00E8 = gv106A Then Exit Sub
  gv106A = p00E8
  If gv106A Then
    gv1056 = 1 + gSectionsForVB(p00E8).RAW_Offset
    gv105A = gSectionsForVB(p00E8).Raw_Size
    Seek g_hFileExe, gv1056
    gv1062 = gSectionsForVB(p00E8).VA_Offset
    gv1066 = gv1062 + gv105A
    gVAToRaw2 = gv1056 - gv1062
  Else
    gv1056 = 0
    gv105A = 0
    gv1062 = 0
    gv1066 = 0
    gVAToRaw2 = 0
  End If
End Sub


Sub StoreSectDataForVB(VA_Offset As Long, Raw_Size As Long, RAW_Offset As Long, Sect_Flag As Integer)
  gSegCount = gSegCount + 1
  
  If gSegCount > UBound(gSectionsForVB) Then ReDim Preserve gSectionsForVB(gSegCount + 10)
  
  gSectionsForVB(gSegCount).RAW_Offset = RAW_Offset
  gSectionsForVB(gSegCount).Raw_Size = Raw_Size
  gSectionsForVB(gSegCount).VA_Offset = VA_Offset
  gSectionsForVB(gSegCount).Flags = Sect_Flag
End Sub

Sub AddToLogTree(Index As Integer, Nametxt As String, _
      Offset As Long, VOffset As Long, Size As Long, Level As Integer)
Dim l00AE As Integer
Dim Counter As Integer
  
 'Exit if No File Number...
  If Offset = 0 Then ErrBeep: Exit Sub
  
  gEntryCounter = gEntryCounter + 1
  
  'Enlarge array if needed
  If gEntryCounter > UBound(gEntrys) Then ReDim Preserve gEntrys(gEntryCounter + 10)
  
  gEntrys(gEntryCounter).Name = Nametxt
  gEntrys(gEntryCounter).Offset = Offset
  gEntrys(gEntryCounter).VOffset = VOffset
  gEntrys(gEntryCounter).Size = Size
  gEntrys(gEntryCounter).M36B3_count = Index
  gEntrys(gEntryCounter).Level_M4E13 = Level
  
  l00AE = 0
  Counter = gEntrys(l00AE).Count
  
  Do While Counter
    If gEntrys(Counter).Offset > Offset + 1 Then Exit Do
    l00AE = Counter
    Counter = gEntrys(l00AE).Count
  Loop
  
  gEntrys(gEntryCounter).Count = Counter
  gEntrys(l00AE).Count = gEntryCounter
  Offset = Offset + Size
End Sub

'StoreSectData VA_Offset, gPESections(i).Raw_Size, RAW_Offset, SectionName, Sect_Flag
Sub StoreSectData(VA_Offset&, Raw_Size&, ByVal RAW_Offset As Long, SectionName As String, Sect_Flag%)

If gIsVB_Ver Then
  StoreSectDataForVB VA_Offset, Raw_Size, RAW_Offset, Sect_Flag
End If
  
  If Raw_Size = 0 Then Exit Sub
  
 'inc because Vb starts fileoffset with 1
  RAW_Offset = RAW_Offset + 1
  
  gPECurSection = gPECurSection + 1
  
  If gPECurSection > UBound(gPESections3) Then ReDim Preserve gPESections3(gPECurSection + 10)
  gPESections3(gPECurSection).VOffset = VA_Offset
  gPESections3(gPECurSection).Raw_Size = Raw_Size
  gPESections3(gPECurSection).RAW_Offset = RAW_Offset
  gPESections3(gPECurSection).Name = SectionName
End Sub

Sub List_gEntrys(SomeListbox As Control)
Dim Listbox_Count As Integer
  SomeListbox.Clear
  SomeListbox.ListIndex = -1
  Listbox_Count = gEntrys(0).Count
   
  Dim tmp
  
  tmp = Array(BlockAlign_r("Name", 20), "Offset", "VOffset", "Size", "M36B3_count", "Level_M4E13", "Count")
  SomeListbox.AddItem Join(tmp, String(1, vbTab))

  
  Do
    
    With gEntrys(Listbox_Count)

    tmp = Array(BlockAlign_r(.Name, 20), Hex(.Offset - 1), Hex(.VOffset), Hex(.Size), .M36B3_count, .Level_M4E13, .Count)
    SomeListbox.AddItem Join(tmp, String(1, vbTab))

    End With
    SomeListbox.ItemData(SomeListbox.NewIndex) = Listbox_Count
    Listbox_Count = gEntrys(Listbox_Count).Count
  Loop While Listbox_Count
End Sub

Sub sub1329(p00C6 As Control)
Dim l00C8
  p00C6.Clear
  For l00C8 = 1 To gPECurSection
    p00C6.AddItem fn086D(gPESections3(l00C8).RAW_Offset) & fn0853(gPESections3(l00C8).VOffset) & "[" & Hex$(gPESections3(l00C8).Raw_Size) & "]"
  Next
End Sub

Sub InitHeadersVars(Offset As Long)
  gPESect_Text_VToRaw = 0
  
  gPESect_Text_VStart = 0
  gPESect_Text_VEnd = 0
  
  gEntryCounter = 0: ReDim gEntrys(10)
  gEntrys(0).Name = gv_ExeFileName
  gPECurSection = 0: ReDim gPESections3(10)
  Offset = 1
  
  AddToLogTree 0, "Headers", Offset, 0, 0, 0
  
  AddToLogTree gcStub_Type, "Stub", Offset, 0, gMZ_Header.OffsetToPE, 1
End Sub

Sub LoadExeStrucs()
Dim slashPos As Integer
Dim l00DC As Form
  
  If Len(g_ExeFilePath) = 0 Then File_Name G_FullExeFileName
  
  slashPos = FindLastOccurenceOf(G_FullExeFileName, "\")
  
  gv_ExeFilePath = Left$(G_FullExeFileName, slashPos)
  gv_ExeFileName = Mid$(G_FullExeFileName, slashPos + 1)
  
  slashPos = InStr(gv_ExeFileName, ".")
  gExeFileNameOnly2 = Left$(gv_ExeFileName, slashPos - 1)
  
  g_IsNE = False
  gIsVB_Ver = False
  
  gVBversion = 0
  gEntryPoint_Page = 0
  
  gImageBase1 = 0
  
 'Close File when opened
  If g_hFileExe Then Close g_hFileExe Else g_hFileExe = FreeFile
  
  'Get Dos-Header
  Open G_FullExeFileName For Binary Access Read As g_hFileExe
   Get g_hFileExe, 1, gTwoByteBuff
   If gTwoByteBuff <> "MZ" Then Exit Sub
   
   Get g_hFileExe, 1, gMZ_Header
   If gMZ_Header.reloc < &H40 Then Exit Sub
   
  'Get PE or NE Header
   Get g_hFileExe, gMZ_Header.OffsetToPE + 1, gTwoByteBuff
   Select Case gTwoByteBuff
   Case "NE"
     g_IsNE = True
     LoadNE
     
   Case "PE"
     gIsVB_Ver = True
     LoadPE
     
   Case Else
     DbgMileStone "no Windows exe?"
     gTwoByteBuff = "??"
     
   End Select
   
   Assume gVBversion, "unknown VB version"
   gIsVB5 = gIsVB_Ver = 5
   
   If l00DC Is Nothing Then Exit Sub
   
   l00DC.Show
   l00DC.control4.value = True
   
End Sub

