Attribute VB_Name = "MODULE23"
' Module23
Option Explicit
Const mc0024 = 20 ' &H14%
Const mc002C = 16 ' &H10%
Const mc0042 = 21 ' &H15%
Const mc0044 = 23 ' &H17%
Const mc0046 = 6 ' &H6%
Const mc0048 = 22 ' &H16%

Sub sub151D(ByVal p0006 As String, p0008 As Integer, ByVal p000A As Long, p000C As Integer)
Dim l000E
Dim ResTypeName As String
Dim i
Dim l0014%
Dim l0016%
Dim l0018 As Integer
Dim l001C As T5427
Dim l0020 As T5440
' Const mc0024 = 20 ' &H14%
' Const mc002C = 16 ' &H10%
  Get g_hFileExe, p000A, l001C
  l0018 = l001C.M5437 + l001C.M5410
  If Len(p0006) Then
    AddToLogTree mc0024, p0006, p000A, p000A - gv1598 + gv159C, Len(l001C) + l0018 * Len(l0020), p0008
  Else
  End If
  p000A = Seek(g_hFileExe)
  For l000E = 1 To l001C.M5437
    Get g_hFileExe, p000A, l0020
    p000A = Seek(g_hFileExe)
    If Len(p0006) Then
      ResTypeName = p0006 & "."
    Else
      ResTypeName = ""
    End If
    If l0020.M544E > 0 Then MsgBox "named rsrc > 0"
    Get g_hFileExe, gv1598 + (l0020.M544E And &H7FFFFFFF), l0014
    For i = 1 To l0014
      Get g_hFileExe, , l0016
      ResTypeName = ResTypeName & Chr$(l0016)
    Next
    
    If p0008 = 1 Then
      If Left$(ResTypeName, 5) = "_IID_" Then
        p000C = fn0381(p000C, "GUID")
      Else
        p000C = fn0381(p000C, ResTypeName)
      End If
    End If
    sub16B9 ResTypeName, p0008, l0020.M5459, p000C
  Next
  For l000E = 1 To l001C.M5410
    Get g_hFileExe, p000A, l0020
    p000A = Seek(g_hFileExe)
    l0014 = l0020.M544E
    If l0014 < 0 Then MsgBox "ord rsrc < 0"
    If Len(p0006) > 1 Then
      ResTypeName = p0006 & ".#" & Hex$(l0014)
    ElseIf l0014 <= mc002C Then
      
      ResTypeName = GetResTypeName(l0014)
      If p0008 = 1 Then
        p000C = l0014
      End If
    
    Else
      MsgBox p0006, , "add rsrc type?"
    End If
    If Len(ResTypeName) = 0 Then
      ResTypeName = "#" & Hex$(l0014)
    End If
    
    sub16B9 ResTypeName, p0008, l0020.M5459, p000C
  Next
End Sub

Private Sub sub16B9(p0030 As String, p0032 As Integer, ByVal p0034 As Long, p0036%)
Dim l003A As T5465
Dim l003C As Integer
Dim l003E As Integer
' Const mc0042 = 21 ' &H15%
' Const mc0044 = 23 ' &H17%
' Const mc0046 = 6 ' &H6%
' Const mc0048 = 22 ' &H16%
  If p0034 < 0 Then
    p0034 = p0034 And &H7FFFFFFF
    sub151D p0030, p0032 + 1, p0034 + gv1598, p0036
  Else
    p0034 = p0034 + gv1598
    Get g_hFileExe, p0034, l003A
    AddToLogTree mc0042, p0030, p0034, p0034 - gv1598 + gv159C, Len(l003A), p0032
    If l003A.RVA_Offset Then
      Select Case p0036
      Case 0
        ErrBeep
        If Left$(p0030, 5) = "_IID_" Then
          p0036 = fn0381(p0036, "GUID")
        Else
          p0036 = fn0381(p0036, Left$(p0030, InStr(p0030, ".") - 1))
        End If
        l003C = mc0044
      Case mc0046: l003C = mc0048
      Case Else
        l003C = mc0044
      End Select
      p0034 = RVAtoRaw(l003A.RVA_Offset)
      AddToLogTree l003C, p0030 & ".CP" & Format$(l003A.M5475), (p0034), p0034 - gv1598 + gv159C, l003A.Size, p0032 + 1
      l003E = InStr(p0030, ".") + 2
      l003E = Val("&h" & Mid$(p0030, l003E))
      If l003E = 0 Then MsgBox p0030, , "ID=0?"
      sub0426 p0036, l003E, p0030, p0034 - 1, l003A.Size, l003A.M5475
    End If
  End If
End Sub
