Attribute VB_Name = "MODULE4"
' Module4
Option Explicit
Dim ResObjNameStrArray As String
Dim m004C As String

Function fn0381(ByVal pv00B4 As Integer, p00B6) As Integer
  If pv00B4 < 0 Then
    fn0381 = pv00B4 And &H1F
    Exit Function
  End If
  fn0381 = AppendToResStrArray(p00B6)
End Function

Function AppendToResStrArray(ByVal NameToAppend As String) As Integer
Dim Pos As Integer
Dim ResTypeLibPos
 'Already included
  Pos = InStr(ResObjNameStrArray, NameToAppend & "§")
  If Pos Then
   'Yes...
    For ResTypeLibPos = 1 To g_ResTypeLibPos
      If g_ResObjs(ResTypeLibPos).ResObjName = Pos Then
        AppendToResStrArray = ResTypeLibPos
        Exit Function
      End If
    Next
    ErrBeep
  End If
  
  QuoteWithParag NameToAppend, ResObjNameStrArray, Pos
  ResTypeLibPos = g_ResTypeLibPos + 1
  g_ResTypeLibPos = ResTypeLibPos
  
  If ResTypeLibPos > UBound(g_ResObjs) Then ReDim Preserve g_ResObjs(ResTypeLibPos + 10)
  
  g_ResObjs(ResTypeLibPos).ResObjName = Pos
  
  AppendToResStrArray = ResTypeLibPos
End Function

Function fn040F(p00A4 As Integer) As String
Dim l00A6 As Integer
Dim l00A8 As String
Dim l00AA As Integer
  l00A6 = gv026C(p00A4).M36BA
  l00AA = gv026C(p00A4).StringBuffOffs
  If l00A6 < 0 Then
    l00A8 = "#" & Hex$(l00A6 And &H7FFF)
    If l00AA Then
      l00A8 = ExtractKeyWord(m004C, l00AA) & "." & l00A8
    End If
  ElseIf l00AA Then
    l00A8 = ExtractKeyWord(m004C, l00AA)
  Else
    ErrBeep
  End If
  fn040F = l00A8
End Function

Function GetResTypeName(p00B0 As Integer) As String
  GetResTypeName = ExtractKeyWord(ResObjNameStrArray, g_ResObjs(p00B0).ResObjName)
End Function

Function fn03EC(p008A As Integer, p008C As Integer, p008E$) As Integer
Dim l0090 As Integer
Dim l0092 As Integer
  l0090 = g_ResObjs(p008A).M3595
  If p008C < 0 Then
    l0092 = p008C
  Else
    l0092 = fn0835(p008E, m004C)
  End If
  sub03F8 l0092, p008A
  fn03EC = gv02B4
End Function

Sub sub025C(p009C As Integer)
Dim l00A0 As Long
  gv02A2 = gv026C(p009C)
  gv02B4 = p009C
  g_FileExeOffset = 1 + gv02A2.RAW_Offset
  l00A0 = gv02A2.Size
  gv02C0 = g_FileExeOffset + l00A0
  Seek g_hFileExe, g_FileExeOffset
End Sub

Sub sub03AE(p006A As Integer)
Dim l006C As Integer
Dim l006E As Long
Dim l0070 As String
  DisOutput "file offset " & Hex$(g_FileExeOffset - 1): WriteOutputTxt
  sub025C p006A
  l006C = gv02A2.M36B3_count
  Select Case l006C
  Case ResID_STRING
    MsgBox "implement dmp_res RT_STRING"
  Case gc01EE
    Get g_hFileExe, g_FileExeOffset, l006E
    DebugPrintVar l006E = &H10, "GUID size"
    Get g_hFileExe, , gv0B84
    DisOutput fn1285(gv0B84): WriteOutputTxt
    Exit Sub
  Case gc01EC
    MsgBox "implement dmp_res RT_TYPELIB"
  Case 1 To gc01F2
  Case Else
    l0070 = GetResTypeName(l006C)
    ErrBeep
  End Select
  sub03B9 g_hFileExe, g_FileExeOffset, 0, gv02A2.Size
End Sub

Sub sub03C9(p007A As Integer, p007E As Control)
Dim l0080 As Integer
  p007E.Clear
  p007E.ListIndex = -1
  l0080 = g_ResObjs(p007A).M3595
  While l0080
    p007E.AddItem fn040F(l0080)
    p007E.ItemData(p007E.NewIndex) = l0080
    l0080 = gv026C(l0080).Count
  Wend
End Sub

Sub sub03DB(p0084 As Control)
Dim l0086 As Integer
  p0084.Clear
  p0084.ListIndex = -1
  For l0086 = 1 To g_ResTypeLibPos
    If g_ResObjs(l0086).M3595 Then
      p0084.AddItem GetResTypeName(l0086)
      p0084.ItemData(p0084.NewIndex) = l0086
    End If
  Next
End Sub

Sub sub03F8(p0096 As Integer, p0098 As Integer)
Dim l009A As Integer
  gv02B4 = 0
  gv02B6 = p0098
  l009A = g_ResObjs(p0098).M3595
  Do While l009A
    If gv026C(l009A).M36BA = p0096 Then
      sub025C l009A
      Exit Sub
    End If
    l009A = gv026C(l009A).Count
  Loop
End Sub

Sub InitResObj()
  ReDim g_ResObjs(gc01F2): g_ResTypeLibPos = gc01F0
  ReDim gv026C(10): gv029E = 0
  ResObjNameStrArray = "": m004C = ""
  QuoteWithParag "CURSOR", ResObjNameStrArray, g_ResObjs(ResID_CURSOR).ResObjName
  QuoteWithParag "BITMAP", ResObjNameStrArray, g_ResObjs(ResID_BITMAP).ResObjName
  QuoteWithParag "ICON", ResObjNameStrArray, g_ResObjs(ResID_ICON).ResObjName
  QuoteWithParag "MENU", ResObjNameStrArray, g_ResObjs(ResID_MENU).ResObjName
  QuoteWithParag "DIALOG", ResObjNameStrArray, g_ResObjs(ResID_DIALOG).ResObjName
  QuoteWithParag "STRING", ResObjNameStrArray, g_ResObjs(ResID_STRING).ResObjName
  QuoteWithParag "FONTDIR", ResObjNameStrArray, g_ResObjs(ResID_FONTDIR).ResObjName
  QuoteWithParag "FONT", ResObjNameStrArray, g_ResObjs(ResID_FONT).ResObjName
  QuoteWithParag "ACCELERATOR", ResObjNameStrArray, g_ResObjs(ResID_ACCELERATOR).ResObjName
  QuoteWithParag "RCDATA", ResObjNameStrArray, g_ResObjs(ResID_RCDATA).ResObjName
  QuoteWithParag "GROUP_CURSOR", ResObjNameStrArray, g_ResObjs(ResID_GROUP_CURSOR).ResObjName
  QuoteWithParag "GROUP_ICON", ResObjNameStrArray, g_ResObjs(ResID_GROUP_ICON).ResObjName
  QuoteWithParag "NAMETABLE", ResObjNameStrArray, g_ResObjs(ResID_NAMETABLE).ResObjName
  QuoteWithParag "VERSIONINFO", ResObjNameStrArray, g_ResObjs(ResID_VERSIONINFO).ResObjName
  
  Assume AppendToResStrArray("TYPELIB") = gc01EC, "preset rsrc types"
  Assume AppendToResStrArray("GUID") = gc01EE, "preset rsrc types"
End Sub

Sub sub0426(p005E As Integer, p0060 As Integer, p0062 As String, p0064 As Long, p0066 As Long, ByVal p6%)
  ' called (p1%,              p2%,              ByVal p3$,       p4&,           p5&, ByVal p6%)
  If p0066 = 0 Then Exit Sub
  If p0060 >= 0 Then
    If Len(p0062) = 0 Then ErrBeep: Exit Sub
  End If
  
  gv029E = gv029E + 1
  
  If gv029E > UBound(gv026C) Then ReDim Preserve gv026C(gv029E + 10)
  gv026C(gv029E).M36B3_count = p005E
  gv026C(gv029E).M36BA = p0060
  
  If Len(p0062) Then QuoteWithParag p0062, m004C, gv026C(gv029E).StringBuffOffs
  gv026C(gv029E).Size = p0066
  gv026C(gv029E).RAW_Offset = p0064
  gv026C(gv029E).Count = g_ResObjs(p005E).M3595
  g_ResObjs(p005E).M3595 = gv029E
End Sub

