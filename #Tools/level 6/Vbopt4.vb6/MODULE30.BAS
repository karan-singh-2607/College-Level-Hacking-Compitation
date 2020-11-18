Attribute VB_Name = "MODULE30"
' Module30
Option Explicit
Const mc002C = 2 ' &H2%
Const mc0044 = 20 ' &H14%
Const mc0084 = -832 ' &HFCC0%
Const mc0086 = 12288 ' &H3000%
Const mc008E = 4096 ' &H1000%
Const mc0090 = 8192 ' &H2000%

Private Sub sub1CE0(p005A As Long, ByVal p005C As Long)
Dim l005E As Integer
Dim l0060 As Integer
  If gvArrSize Then
    l0060 = gBuggy(l005E).Next_M5721
    Do While l0060
      If gBuggy(l0060).VOffset >= p005A Then
        If gBuggy(l0060).VOffset > p005A Then Exit Do
        If gBuggy(l0060).M36B3_count And gc1880 Then
          If gBuggy(l0060).M4C81 = p005C Then Exit Sub
        End If
      End If
      l005E = l0060
      l0060 = gBuggy(l0060).Next_M5721
    Loop
  End If
  l0060 = fn1D50(p005A, l005E, l0060)
  gBuggy(l0060).M4C81 = p005C
  gBuggy(l0060).M36B3_count = gc187E Or gc1880
End Sub

Sub sub186C(p0062 As Long, ByVal pv0064 As Integer)
Dim l0066 As Integer
Dim l0068 As Integer
Dim l006A As Integer
  If p0062 <= gMetaCodeLines(gMetaCodeLinesCount).VOffset Then Exit Sub
  
  Select Case pv0064
  Case gc185C, gc185E, gc1860, gc1870
    If gVACmdStart > p0062 Then pv0064 = pv0064 Or gc1864
  End Select
  
  If gvArrSize Then
    l0068 = gBuggy(l0066).Next_M5721
    Do While l0068
      If gBuggy(l0068).VOffset >= p0062 Then
        If gBuggy(l0068).VOffset > p0062 Then Exit Do
        l006A = gBuggy(l0068).M36B3_count
        If l006A = 0 Then GoTo L3CB82
        If l006A = pv0064 Then Exit Sub
      End If
      l0066 = l0068
      l0068 = gBuggy(l0068).Next_M5721
    Loop
  End If
  l0068 = fn1D50(p0062, l0066, l0068)

L3CB82:
  gBuggy(l0068).M36B3_count = pv0064
  gIs3158 = True
  Select Case pv0064
  Case 0
    Exit Sub
  Case gc1870, mc0044
  Case gc186C
    Exit Sub
  Case gc185C
  Case mc002C
  Case Else
    DbgMileStone "more label types?"
    Exit Sub
  End Select
End Sub

Sub sub18B7(ByVal pv0070 As Integer, p0072 As String, p0074 As Integer)
Dim l0076 As Integer
' Const mc0084 =-832 ' &HFCC0%
' Const mc0086 = 12288 ' &H3000%
' Const mc008E = 4096 ' &H1000%
' Const mc0090 = 8192 ' &H2000%
  If gVACmdEnd <= gMetaCodeLines(gMetaCodeLinesCount).VOffset Then Exit Sub
  If gVACmdEnd = gv34EC Then Exit Sub
  l0076 = pv0070 And &HFF
  If Left$(p0072, 1) = "'" Then
    pv0070 = pv0070 Or gc1882
  End If
  Select Case l0076
  Case gc187C
    DebugPrintVar p0074, "valid arg"
    Select Case gEXEdefEntries(p0074).M6894.M36B3_count And mc0084
    Case mc0086
      gv356E(gEXEdefEntries(p0074).M689A) = gMetaCodeLinesCount + 1
      If (gEXEdefEntries(p0074).M6894.M36B3_count And mc008E) = 0 Then
        DbgMileStone "assume unused"
      End If
    Case mc0090
      gv356E(gEXEdefEntries(p0074).M689A) = gMetaCodeLinesCount + 1
      pv0070 = pv0070 Or gc1882
    End Select
    l0076 = 0
  Case gc185C, gc185E, gc1860
    If gBuggy(p0074).VOffset < gVACmdEnd Then pv0070 = pv0070 Or gc1864
  Case Else
    l0076 = 0
  End Select
  gMetaCodeLinesCount = gMetaCodeLinesCount + 1
  If gMetaCodeLinesCount > UBound(gMetaCodeLines) Then ReDim Preserve gMetaCodeLines(gMetaCodeLinesCount + 10)
  gMetaCodeLines(gMetaCodeLinesCount).VOffset = gv34EC: gv34EC = gVACmdEnd
  gMetaCodeLines(gMetaCodeLinesCount).M36B3_count = pv0070
  gMetaCodeLines(gMetaCodeLinesCount).M574B = p0072
  gMetaCodeLines(gMetaCodeLinesCount).M5742 = p0074
  If l0076 = 0 Then Exit Sub
  gMetaCodeLines(gMetaCodeLinesCount).M5755 = gBuggy(p0074).M572A
  gBuggy(p0074).M572A = gMetaCodeLinesCount
End Sub

Sub sub1CF2(p0094 As Control)
Dim l0096 As Integer
Dim l0098 As String
  p0094.Clear
  l0096 = gBuggy(0).Next_M5721
  While l0096
    l0098 = Hex$(gBuggy(l0096).VOffset) & fn085C(gBuggy(l0096).M36B3_count)
    l0098 = l0098 & " " & fn1D20(l0096)
    p0094.AddItem l0098
    p0094.ItemData(p0094.NewIndex) = l0096
    l0096 = gBuggy(l0096).Next_M5721
  Wend
End Sub

Sub sub1D03(p009E As Control)
Dim l00A0 As Integer
  p009E.Clear
  Do
    p009E.AddItem CStr(l00A0) & fn0853(gBuggy(l00A0).VOffset) & fn085C(gBuggy(l00A0).M36B3_count) & " " & fn1D20(l00A0)
    l00A0 = gBuggy(l00A0).Next_M5721
  Loop While l00A0
End Sub

Function fn1D12(p00A6 As Long) As Integer
Dim l00A8 As Integer
  If p00A6 < gBuggy(0).VOffset Then Exit Function
  l00A8 = gBuggy(0).M5718
  While gBuggy(l00A8).VOffset > p00A6
    l00A8 = gBuggy(l00A8).M5718
  Wend
  If gBuggy(l00A8).VOffset < p00A6 Then
    l00A8 = fn1D50(p00A6, l00A8, gBuggy(l00A8).Next_M5721)
    gBuggy(l00A8).M36B3_count = gc187E
  ElseIf False Then
    Do Until gBuggy(l00A8).M36B3_count And gc187E
      l00A8 = gBuggy(l00A8).M5718
    Loop
  End If
  fn1D12 = l00A8
End Function

Function fn1D20(p00AC As Integer) As String
Dim l00AE As Integer
Dim l00B0 As String
  l00AE = gBuggy(p00AC).M36B3_count
  If l00AE And gc1882 Then Exit Function
  If l00AE And gc1880 Then
    l00B0 = Format$(ToUnSigned(gBuggy(p00AC).M4C81))
    fn1D20 = l00B0
  Else
    Select Case l00AE And &HFF
    Case Is < gc1866, gc186C
      l00B0 = "L" & CStr(p00AC)
    Case gc1868, gc186A
      l00B0 = "Do"
    Case gc1874
      l00B0 = "Case Else"
    Case gc1876
      l00B0 = "End Select"
    Case gc1878
      l00B0 = "Else"
    Case gc187A
      l00B0 = "End If"
    Case Else
    End Select
    fn1D20 = l00B0
  End If
End Function

Function fn1D2E() As Integer
  fn1D2E = fn1D5C(0)
End Function

Sub sub1D3F(p00B6 As Integer)
Dim l00B8 As Integer
Dim l00BA As Integer
Dim l00BE As T56DF
  If p00B6 Then
    ReDim gMetaCodeLines(10): gMetaCodeLinesCount = 0
    ReDim gBuggy(10): gvArrSize = 0
    gBuggy(0).VOffset = gModuleVAStart
    gBuggy(0).M36B3_count = gc187E
    l00BA = fn1D50(gModuleVAEnd, 0, 0)
    gBuggy(l00BA).M36B3_count = gc187E Or gc1882
  Else
    l00B8 = gv2F50.M5340
    If l00B8 Then
      Get g_hFileExe, gVAToRaw2 + gModuleVAEnd + l00B8, l00B8
      For l00BA = 1 To l00B8
        Get g_hFileExe, , l00BE
        sub1CE0 gModuleVAStart + ToUnSigned(l00BE.hFile), l00BE.M4C81
      Next
    End If
  End If
End Sub

Sub sub1A12(p00CC As Integer)
Dim i%
Dim ArrSize%
  sub1D3F True
  
  If p00CC Then Get p00CC, , ArrSize
  
  If ArrSize > 0 Then
    gvArrSize = ArrSize
    ReDim gBuggy(gvArrSize)
    For i = 0 To gvArrSize
      Get p00CC, , gBuggy(i)
    Next
  End If
  
  sub1D3F False
  
End Sub

Private Function fn1D50(p00D8 As Long, ByVal pv00DA As Integer, ByVal pv00DC As Integer) As Integer
  gvArrSize = gvArrSize + 1
  If gvArrSize > UBound(gBuggy) Then
    ReDim Preserve gBuggy(gvArrSize + 10)
  End If
  gBuggy(gvArrSize).Next_M5721 = pv00DC: gBuggy(pv00DC).M5718 = gvArrSize
  gBuggy(gvArrSize).M5718 = pv00DA: gBuggy(pv00DA).Next_M5721 = gvArrSize
  gBuggy(gvArrSize).VOffset = p00D8
  gIs3158 = True
  fn1D50 = gvArrSize
End Function

Function fn1D5C(p00E0 As Integer) As Integer
Dim l00E2 As Integer
  Do
    l00E2 = gBuggy(p00E0).Next_M5721
    If l00E2 = 0 Then Exit Do
    p00E0 = l00E2
  Loop While gBuggy(p00E0).M36B3_count And gc1882
  fn1D5C = p00E0
End Function

Function fn1D6C(p00E6 As Integer) As Integer
  If p00E6 >= gMetaCodeLinesCount Then Exit Function
  fn1D6C = p00E6 + 1
End Function

Sub sub1A74(p00E8 As Integer)
Dim i%
  If p00E8 Then
  If gvArrSize Then
    Put p00E8, , gvArrSize
    For i = 0 To gvArrSize
      Put p00E8, , gBuggy(i)
    Next
  Else
    i = -1
    Put p00E8, , i
  End If
  End If
End Sub

Function fn1D79() As Integer
  If gvArrSize Then
    fn1D79 = 2 + (gvArrSize + 1) * Len(gBuggy(0))
  Else
    fn1D79 = 2
  End If
End Function

Function Translate_MetaCodeLines_VAToLineNr(VAOffset As Long) As Integer
Dim i&
  For i = gMetaCodeLinesCount To 1 Step -1
    If gMetaCodeLines(i).VOffset <= VAOffset Then
      Translate_MetaCodeLines_VAToLineNr = i
      Exit Function
    End If
  Next
End Function
