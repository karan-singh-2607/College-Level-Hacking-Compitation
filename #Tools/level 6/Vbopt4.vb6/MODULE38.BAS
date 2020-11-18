Attribute VB_Name = "MODULE38"
' Module38
Option Explicit
Dim m000A(1) As String
Dim m0020 As Integer

Function fn2367(p002A As String) As Integer
Dim l002C As String
Dim l002E As Integer
Dim l0030 As Integer
Dim l0032
  l002C = " " & p002A & " "
  For l0032 = 0 To 1
    l002E = InStr(m000A(l0032), l002C)
    If l002E Then
      l0030 = Asc(Mid$(m000A(l0032), l002E - 1))
      If l0032 Then l0030 = l0030 + 256
      fn2367 = l0030
      Exit Function
    End If
  Next
  m0020 = m0020 + 1
  l0030 = m0020
  fn2367 = l0030
  gv2730(l0030) = p002A
  If l0030 > 255 Then
    l0032 = 1
    l0030 = l0030 - 256
  Else
    l0032 = 0
  End If
  m000A(l0032) = m000A(l0032) & Chr$(l0030) & l002C
End Function

Sub sub2314()
  ReDim gv2730(511): m0020 = 0
  Erase m000A
  gv2730(0) = "<undefined>"
End Sub

Function fn2375() As String
Dim l0036 As String
  If gAsmNew.M65EC Then
    l0036 = gv2730(gCmdByteStruc.M6452)
  Else
    l0036 = gv2730(gCmdByteStruc.M6448)
  End If
  If Right$(l0036, 1) = "?" Then
    If gAsmNew.M65EC Then
      Mid$(l0036, Len(l0036), 1) = "l"
    Else
      Mid$(l0036, Len(l0036), 1) = "w"
    End If
  End If
  fn2375 = l0036
End Function
