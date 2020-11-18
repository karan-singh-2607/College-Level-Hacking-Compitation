Attribute VB_Name = "MODULE8"
' Module8
Option Explicit

Function fn06CC(ByVal pv0050 As Integer, ByVal p0052 As Long) As Integer
  If pv0050 < 0 Then pv0050 = -pv0050 + gSegCount
  fn06CC = CreateSym(fn071D(pv0050, p0052))
End Function

Function fn0710(p0040 As Integer) As String
  fn0710 = "#" & Format$(p0040)
End Function

Function fn071D(p0044 As Integer, p0046 As Long) As Long
  fn071D = CLng(p0044) * &H10000 Or (p0046 And &HFFFF&)
End Function

Function fn0727(p004A As Integer, p004C As Long) As Long
  fn0727 = p004A * &H10000 Or p004C
End Function

Function fn073D(p0032 As Integer, p0034 As Long) As Long
  If p0032 > gSegCount Then ErrBeep: Exit Function
  fn073D = gSectionsForVB(p0032).RAW_Offset + 1 + (p0034 And &HFFFF&)
End Function

Sub sub06D9()
Dim l0056
  WriteOutputTxt
  For l0056 = 1 To gSegCount
    DisOutput Str$(l0056) & " " & Hex$(gSectionsForVB(l0056).RAW_Offset) & " [" & Hex$(gSectionsForVB(l0056).Raw_Size) & "]"
    WriteOutputTxt
  Next l0056
End Sub

Sub LoadNE()
 'LoadNEHeader
  Get g_hFileExe, gMZ_Header.OffsetToPE + 1, gNE_Header
  
  gSegCount = gNE_Header.SegmentTableEntryCount
  gImportDLLCount = gNE_Header.ModuleTableEntryCount
  
  sub06F3
  
  NE_GetResource
  
End Sub

Sub sub072F(p0020 As Control)
Dim l0022
Dim l0024 As String
  p0020.Clear
  p0020.ListIndex = -1
  For l0022 = 1 To gSegCount
    l0024 = Str$(l0022) & " [" & Hex$(gSectionsForVB(l0022).Raw_Size) & "] at " & Hex$(gSectionsForVB(l0022).RAW_Offset) & " " & Hex$(gSectionsForVB(l0022).Flags)
    p0020.AddItem l0024
    p0020.ItemData(p0020.NewIndex) = l0022
  Next l0022
End Sub

