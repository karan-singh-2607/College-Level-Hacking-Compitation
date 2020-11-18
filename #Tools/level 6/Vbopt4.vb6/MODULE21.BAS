Attribute VB_Name = "MODULE21"
' Module21
Option Explicit
Const mc003E = 3 ' &H3%
Const mc0040 = 5 ' &H5%
Const mc0042 = 2 ' &H2%

Sub sub13B9(p001E As Integer, p0020 As Long, p0022 As Integer)
Dim l0024 As Integer
Dim l0026 As Integer
Dim l0028 As Integer
Dim l002A As Integer
  gv164E = gv164E + 1
  l0024 = gv164E
  If l0024 > UBound(gv161C) Then ReDim Preserve gv161C(l0024 + 50)
  gv161C(l0024).VOffset = p0020
  gv161C(l0024).M36B3_count = p0022
  gv161C(l0024).subRef_M3C6A = p001E
  l002A = MapModule(p0020)
  If l002A > UBound(gv15E8) Then ErrBeep: Exit Sub
  l0028 = gv15E8(l002A)
  If l0028 = 0 Then
    l0028 = gv161C(l0028).M54BB
    If l0028 = 0 Then
      gv161C(0).M54BB = l0024
      gv161C(0).Count = l0024
      gv15E8(l002A) = l0024
      Exit Sub
    End If
    While gv161C(l0028).VOffset > p0020
      l0028 = gv161C(l0028).M54BB
    Wend
    l0026 = gv161C(l0028).Count
    gv15E8(l002A) = l0024
  Else
    If gv161C(l0028).VOffset > p0020 Then
      gv15E8(l002A) = l0024
      l0026 = l0028
      l0028 = gv161C(l0026).M54BB
    Else
      l0026 = gv161C(l0028).Count
      Do While l0026
        If gv161C(l0026).VOffset > p0020 Then Exit Do
        l0026 = gv161C(l0026).Count
      Loop
      l0028 = gv161C(l0026).M54BB
    End If
  End If
  If gv161C(l0028).VOffset = p0020 Then
    gv164E = gv164E - 1
    Exit Sub
  End If
  gv161C(l0028).Count = l0024: gv161C(l0024).Count = l0026
  gv161C(l0026).M54BB = l0024: gv161C(l0024).M54BB = l0028
End Sub

Function fn1624(p0030&) As Integer
Dim l0032 As Integer
  l0032 = gv15E8(MapModule(p0030))
  
  While l0032
    If gv161C(l0032).VOffset >= p0030 Then
      fn1624 = l0032
      Exit Function
    End If
    l0032 = gv161C(l0032).Count
  Wend

End Function

Function fn1631(p0036&, p0038 As Integer) As Integer
Dim l003A As Integer
' Const mc003E = 3 ' &H3%
' Const mc0040 = 5 ' &H5%
' Const mc0042 = 2 ' &H2%
  l003A = fn1624(p0036)
  If p0038 Then
    If gv161C(l003A).VOffset <> p0036 Then Exit Function
  End If
If gIsVB_Ver Then
Else
  Select Case gv161C(l003A).M36B3_count
  Case mc003E, mc0040
  Case mc0042
    sub00B9 gv161C(l003A)
  End Select
End If
  fn1631 = l003A
End Function

Sub MemPrepareSth()
  gv164E = 0: ReDim gv161C(0)
  ReDim gv15E8(MapModule(gImageBase2) To MapModule(gVOffset1))
End Sub

Function fn164F(p004A As Integer) As Integer
  fn164F = gv161C(p004A).Count
End Function

Function fn165B(p004E As Integer) As Integer
  fn165B = gv161C(p004E).M54BB
End Function
