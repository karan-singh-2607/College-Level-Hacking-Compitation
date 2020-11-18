Attribute VB_Name = "MODULE11"
' Module11
Option Explicit

Sub sub0951(p0048 As ComboBox)
Dim l004A
Dim l004C As String
Dim l004E As Integer
  p0048.Clear
  For l004A = 1 To gv097C
    l004C = gv094A(l004A).M3F0B
    l004E = FindLastOccurenceOf(l004C, "\")
    If l004E Then
      l004C = Mid$(l004C, l004E + 1)
    End If
    p0048.AddItem l004C
    p0048.ItemData(p0048.NewIndex) = l004A
  Next
  p0048.ListIndex = -1
End Sub

Sub sub0963(p0054 As ComboBox, p0056 As Integer, p0058 As Integer)
Dim l005A As Integer
Dim l005C As Integer
  p0054.Clear
  l005C = gv094A(p0056).M3F1B
  If p0058 Then l005A = l005C Else l005A = l005C + 2
  For l005A = l005C To l005C + gv094A(p0056).M3F13
    If gv0AD6(l005A).M4045 > gv0AD6(l005A - 1).M4045 Then
      p0054.AddItem gv0AD6(l005A).Name
      p0054.ItemData(p0054.NewIndex) = l005A
    End If
  Next l005A
  p0054.ListIndex = -1
End Sub

Function fn0972(p0060 As Integer) As String
  fn0972 = gv0AD6(p0060).Name
End Function
