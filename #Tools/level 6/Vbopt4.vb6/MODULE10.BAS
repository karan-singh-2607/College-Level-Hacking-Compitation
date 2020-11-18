Attribute VB_Name = "MODULE10"
' Module10
Option Explicit

Function ExtractKeyWord(Str As String, StartPos As Integer) As String
Dim endPos As Integer
  If StartPos = 0 Then Exit Function
  endPos = InStr(StartPos, Str, "§")
  ExtractKeyWord = Mid$(Str, StartPos, endPos - StartPos)
End Function

Function fn0822(p00B6 As String) As Long
  gv0786.M3CE1 = p00B6
  LSet gv078C = gv0786
  fn0822 = gv078C.M3D0D
End Function

Function fn0835(p00C8 As String, p00CA As String) As Integer
  fn0835 = InStr(p00CA, p00C8 & "§")
End Function

Function fn0842(p00CE As Integer) As String
  fn0842 = Right$(Hex$(p00CE Or &H100), 2)
 End Function

Function ConCatHex8(p00D2 As Integer) As String
  ConCatHex8 = " " & Right$(Hex$(p00D2 Or &H100), 2)
End Function

Function fn0853(p00D6 As Long) As String
Dim l00DA As String * 9
  RSet l00DA = Hex$(p00D6)
  fn0853 = l00DA
End Function

Function fn085C(p00DE As Integer) As String
  fn085C = " " & Right$(Hex$(p00DE Or &H10000), 4)
End Function

Function fn0865(p00E2 As Currency) As String
Dim l00E4 As Long
Dim l00E8 As String * 8
  l00E4 = p00E2
  RSet l00E8 = Hex$(l00E4)
  fn0865 = l00E8
End Function

Function fn086D(p00EC As Long) As String
Dim l00F0 As String * 8
  RSet l00F0 = Hex$(p00EC)
  fn086D = l00F0
End Function

Function fn087D(ByVal p00F8 As Long) As Integer
  fn087D = (p00F8 And &HFFFF0000) \ &H10000
End Function

Function fn0887(p00FC As Integer, p00FE As Integer) As Long
  fn0887 = (CLng(p00FC) And &HFFFF&) * p00FE
End Function

Function fn088F(ByVal p0102 As Long) As Integer
  If p0102 And &H8000& Then
    fn088F = CInt(p0102 Or &HFFFF0000)
  Else
    fn088F = p0102 And &HFFFF&
  End If
End Function

Function fn0899(p0106 As Integer) As String
  gv0738.NameSize = p0106
  LSet gv0734 = gv0738
  fn0899 = gv0734.M3CE1
End Function

Function fn08A0(p010A As Long) As String
  gv078C.M3D0D = p010A
  LSet gv0786 = gv078C
  fn08A0 = gv0786.M3CE1
End Function

Function fn08A7(p010E As String, p0110 As Integer) As Integer
  gv0734.M3CE1 = Mid$(p010E, p0110, 2)
  LSet gv0738 = gv0734
  fn08A7 = gv0738.NameSize
End Function

Function AppendPathSlash(ByVal p0114 As String) As String
  If Right$(p0114, 1) = "\" Then AppendPathSlash = p0114 Else AppendPathSlash = p0114 & "\"
End Function

Function FindLastOccurenceOf(Str As String, ByVal StrtoFind As String) As Integer
Dim tmpPos As Integer
  Do
    FindLastOccurenceOf = tmpPos
    tmpPos = InStr(tmpPos + 1, Str, StrtoFind)
  Loop While tmpPos
End Function

Function fn08C9(p0120 As String) As String
Dim l0122
  For l0122 = 1 To Len(p0120)
    Mid$(p0120, l0122, 1) = Chr$(Asc(Mid$(p0120, l0122, 1)) Xor &H1F Xor l0122)
  Next
  fn08C9 = p0120
End Function

Function fn08E5(p012A%, p012C As Integer) As Integer
Dim l012E As Long
  l012E = ToUnSigned(p012A) + p012C
  If l012E >= &H8000& Then l012E = l012E Or &HFFFF0000
  fn08E5 = l012E
End Function

Function fn08ED(p0132%, p0134 As Integer) As Integer
  fn08ED = ToUnSigned(p0132) \ p0134
End Function

Function ToUnSigned(p0138 As Integer) As Long
  ToUnSigned = CLng(p0138) And &HFFFF&
End Function

Function fn08FD(p013C%, p013E As Integer) As Integer
Dim l0140 As Long
  l0140 = ToUnSigned(p013C) * p013E
  If l0140 >= &H8000& Then l0140 = l0140 Or &HFFFF0000
  fn08FD = l0140
End Function

Function fn090F(p0148 As String) As Long
  fn090F = CLng("&h" & p0148)
End Function

Function fn0917(p014C As String, p014E As Integer) As String
Dim l0150 As Integer
  If p014E < 0 Then
    fn0917 = "$" & toHex16(CLng(p014E)) & "?" 'CLng(
    Stop
  Else
    l0150 = InStr(p014E, p014C, Chr$(0)) - p014E
    fn0917 = Mid$(p014C, p014E, l0150)
  End If
End Function

Function StrCutZero(p0154 As String) As String
Dim l0156 As Integer
  l0156 = InStr(p0154, Chr$(0))
  If l0156 Then
    StrCutZero = Left$(p0154, l0156 - 1)
  Else
    StrCutZero = p0154
  End If
End Function

Sub Assume(ByVal Condition As Integer, ByVal ErrTitle As String)
  If Condition Then Exit Sub
  MsgBox ErrTitle, VbMsgBoxStyle_vbCritical, "assertion failed - aborting"

End Sub

'
Sub ErrBeep2()
  Beep
End Sub

Sub ErrBeep()
  Beep
End Sub

Sub DebugPrintVar(ByVal Var%, ByVal VarName$)
   Debug.Print , "'" & VarName & "': " & Var
'Stop
End Sub

Sub DbgVar(ByVal p1%, ByVal p2$)
   Debug.Print p2 & ":" & p1
End Sub

Sub QuoteWithParag(ByVal StrToAppend As String, Str As String, Pos As Integer)
 'Append '§' at the end if needed
  If Right$(StrToAppend, 1) <> "§" Then StrToAppend = StrToAppend & "§"
  
 'Is already in string?
  Pos = InStr(Str, StrToAppend)
  If Pos = 0 Then
    'No
    Pos = Len(Str) + 1
    
    'Is First in Line?
    If Pos = 1 Then
     'Yes Append parag also in Front
      Pos = 2
      Str = "§" & StrToAppend
    Else
      'No - so just concat
      Str = Str & StrToAppend
    End If
  
  End If

End Sub

Sub sub0815(p00AA As ComboBox, p00AC As Variant)
Dim l00B0
Dim l00B2 As Long
  If VarType(p00AC) = 8 Then
    For l00B0 = p00AA.ListCount - 1 To 0 Step -1
      If p00AA.List(l00B0) = p00AC Then Exit For
    Next
  Else
    l00B2 = p00AC
    For l00B0 = p00AA.ListCount - 1 To 0 Step -1
      If p00AA.ItemData(l00B0) = l00B2 Then Exit For
    Next
  End If
  p00AA.ListIndex = l00B0
End Sub

Sub sub0829(p00BC As Integer, p00BE As Integer)
Dim l00C4 As Long
  gv085A.M3D8A = p00BE
  gv085A.M3D92 = p00BC
  l00C4 = p00BE And &HFFFF&
  gv085A.M3DC7 = (l00C4 And &H1F) * 2
  gv085A.M3DC0 = (l00C4 \ &H20) And &H3F
  gv085A.M3DB9 = (l00C4 \ &H800) And &H1F
  l00C4 = p00BC And &HFFFF&
  gv085A.M3DB2 = l00C4 And &H1F
  gv085A.M3DAB = (l00C4 \ &H20) And &HF
  gv085A.M3DA3 = ((l00C4 \ &H200) And &H7F) + 1980
  gv085A.M3D85 = DateSerial(gv085A.M3DA3, gv085A.M3DAB, gv085A.M3DB2) + TimeSerial(gv085A.M3DB9, gv085A.M3DC0, gv085A.M3DC7)
  gv085A.M0917 = Format$(gv085A.M3D85, "hh:mm:ss")
  gv085A.M3D9B = Format$(gv085A.M3D85, "dd.mm.yyyy")
End Sub

Sub sub08D9(p0124 As Variant)
  gv085A.M3D85 = p0124
  gv085A.M3DA3 = Year(p0124)
  gv085A.M3DAB = Month(p0124)
  gv085A.M3DB2 = Day(p0124)
  gv085A.M3DB9 = Hour(p0124)
  gv085A.M3DC0 = Minute(p0124)
  gv085A.M3DC7 = Second(p0124)
  gv085A.M3D8A = (fn08FD(gv085A.M3DB9, &H800) Or gv085A.M3DC0 * &H20 Or gv085A.M3DC7 \ 2)
  gv085A.M3D92 = (gv085A.M3DA3 - 1980) * &H200 + gv085A.M3DAB * &H20 + gv085A.M3DB2
  gv085A.M0917 = Format$(gv085A.M3D85, "hh:mm:ss")
  gv085A.M3D9B = Format$(gv085A.M3D85, "dd.mm.yyyy")
End Sub

Sub sub0905(p0142 As Variant)
  gv085A.M3D85 = p0142
End Sub

Function toHex16(ByVal p00F4 As Long) As String
  If p00F4 < &H10000 Then
    toHex16 = Right$(Hex$(p00F4 Or &H10000), 4)
  Else
    toHex16 = Hex$(p00F4)
  End If
End Function

