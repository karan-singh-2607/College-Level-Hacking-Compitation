Attribute VB_Name = "MODULE14"
' Module14
Option Explicit
Const mc0048 = -12288 ' &HD000%
Const mc004A = 15 ' &HF%
Const mc004C = 11 ' &HB%
Const mc007C = 63 ' &H3F%

Function fn0F3C(p0050 As Integer, p0052 As Integer) As String
Dim l0054 As Integer
Dim l0056
Dim l0058 As Integer
  If p0052 Then
    l0054 = gv38BA(p0052).M4950
    For l0056 = gv38BA(p0052).M4959 To 1 Step -1
      l0058 = l0054 + l0056
      If p0050 >= gv38F0(l0058).hFile Then
        If p0050 = gv38F0(l0058).hFile Then
          fn0F3C = fn0F4A(l0058)
        Else
If gIsVB_Ver Then
Else
          ErrBeep
End If
          fn0F3C = fn0F4A(l0058) & "+" & toHex16(p0050 - gv38F0(l0058).hFile)
        End If
        Exit Function
      End If
    Next
  End If
  ErrBeep
  fn0F3C = "+" & toHex16(CLng(p0050))
End Function

Function fn0F4A(p0060 As Integer) As String
  fn0F4A = "." & ExtractKeyWord(gv3606, gv38F0(p0060).StringBuffOffs)
End Function

Sub sub0AE5()
  gv38EC = 0: gv3922 = 0
  ReDim gv38BA(10), gv38F0(10)
  gv360A = 0
  gv16B0 = True
End Sub

Sub sub0B4F(p001E As Integer, p0020 As Integer, p0022 As String)
Dim l0024 As Integer
  l0024 = InStr(p0022, " ")
  p0022 = Trim$(Mid$(p0022, l0024 + 1))
  l0024 = InStr(p0022, "'")
  If l0024 Then p0022 = Trim$(Left$(p0022, l0024 - 1))
  gv38EC = gv38EC + 1
  If gv38EC > UBound(gv38BA) Then ReDim Preserve gv38BA(gv38EC + 10)
  gv38BA(gv38EC).Name = p0022
  gv38BA(gv38EC).M403B = p001E
  gv38BA(gv38EC).hFile = Seek(p0020)
End Sub

Sub sub0F08(p0026 As Integer, p002A As Control)
Dim l002C As Integer
Dim l002E As Integer
Dim l0030
  p002A.Clear
  l002E = gv38BA(p0026).M4950
  For l0030 = 1 To gv38BA(p0026).M4959
    l002C = l002E + l0030
    p002A.AddItem fn0F4A(l002C)
    p002A.ItemData(p002A.NewIndex) = l002C
  Next
End Sub

Sub FillCombo_userTypes2(p0038 As ComboBox)
Dim l003A
  p0038.Clear
  p0038.ListIndex = -1
  For l003A = 1 To gv38EC
    p0038.AddItem gv38BA(l003A).Name
    p0038.ItemData(p0038.NewIndex) = l003A
  Next
End Sub

Sub sub0F2A(p003C() As Integer, p0042 As Long, p0044 As Integer)
Dim l0046 As Integer
' Const mc0048 =-12288 ' &HD000%
' Const mc004A = 15 ' &HF%
' Const mc004C = 11 ' &HB%
  l0046 = gv38BA(p0044).M4950
  For l0046 = l0046 + 2 To l0046 + gv38BA(p0044).M4959
    p003C(p0042 + gv38F0(l0046).hFile) = l0046 Or mc0048
    If (gv38F0(l0046).M6894.M36B3_count And mc004A) = mc004C Then
      sub0F2A p003C(), p0042 + gv38F0(l0046).hFile, gv38F0(l0046).M6894.M42D1
    End If
  Next
End Sub

Sub sub0F5B()
Dim l0066
Dim l0068 As Integer
Dim l006A As Integer
Dim l006C
Dim l006E As Long
Dim l0070 As Long
' Const mc007C = 63 ' &H3F%
  gv3752 = 0
  For l0066 = 1 To gv38EC
    gv3752 = gv3752 + gv38BA(l0066).size
  Next
  ReDim gv3720(gv3752)
  For l0066 = 1 To gv38EC
    l0068 = gv38BA(l0066).M4950
    For l006C = 1 To gv38BA(l0066).M4959
      l006A = l0068 + l006C
      l006E = ToUnSigned(gv38F0(l006A).hFile)
      gv3720(l0070 + l006E) = l006A Or mc0048
      If (gv38F0(l006C).M6894.M36B3_count And mc007C) = mc004C Then
        sub0F2A gv3720(), l006E, gv38F0(l006C).M6894.M42D1
      End If
    Next
    l0070 = l0070 + ToUnSigned(gv38BA(l0066).size)
  Next
End Sub

