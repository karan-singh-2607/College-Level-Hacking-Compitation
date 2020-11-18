Attribute VB_Name = "MODULE9"
' Module9
Option Explicit

Sub DisOutput(DisOutputTxt As String)
  gDisOutputTxt = gDisOutputTxt & DisOutputTxt
End Sub

Sub WriteOutputTxt()
On Error Resume Next
  
  If hFile_Out Then
    Print #hFile_Out, gDisOutputTxt
  Else
    
    If Len(gDisOutputTxt) Then
      Listbox_File_Out.AddItem gDisOutputTxt
      Listbox_File_Out.ItemData(Listbox_File_Out.NewIndex) = gVACmdStart
    End If
    
  End If
  
  gDisOutputTxt = ""
  
End Sub

Sub sub01BA(p005A As Integer, p005C As Long, p005E As Long, p0060 As Long)
  p005E = p0060
  Seek p005A, p005C + p005E
End Sub

Sub sub0268(p0020 As Integer, p0024 As Long, p3&)
Seek p0020, p0024
End Sub

Sub sub0272(p0026 As Integer, p002A As Long, p3&)
' called (p1%, p2 As type, p3&)
  Seek p0026, p002A
End Sub

Sub sub03B9(p004A As Integer, p004C As Long, p004E As Long, p0050 As Long)
  p004E = p0050
  Seek p004A, p004C + p0050
End Sub

Sub sub0776(p0052%, p0054 As Long, p0056 As Long, p0058 As Long)
  If (p0056 Xor p0058) And 1 Then
    sub03B9 p0052, p0054, p0056, p0058
  Else
    sub01BA p0052, p0054, p0056, p0058
  End If
End Sub

Sub sub0784(p0062 As String, p0064 As Integer)
Dim l0066 As Integer
  l0066 = InStr(p0064, p0062, Chr$(0)) - p0064
  If l0066 < 0 Then l0066 = 0
  DisOutput " """ & Mid$(p0062, p0064, l0066) & """"
  WriteOutputTxt
End Sub

Sub UnlinkDecompilerListBox()
  If hFile_Out Then
    Close hFile_Out
    hFile_Out = 0
  End If
  
  Set Listbox_File_Out = Nothing
  Screen.MousePointer = 0
End Sub

Sub sub079A(ByVal p002E As String)
  If hFile_Out Then
    MsgBox "file already open"
    Close hFile_Out
  End If
  hFile_Out = FreeFile
  Screen.MousePointer = 11
  Open p002E For Output As hFile_Out
End Sub

Sub sub07A3(p0030 As Integer)
  gDisOutputTxt = Space$(p0030)
  gv06EA = p0030
End Sub

Sub SetDecompilerListBox(p0034 As ListBox)
  Set Listbox_File_Out = p0034
  p0034.Clear
End Sub

Sub sub07B7(ByVal p0036 As String)
  If Len(gDisOutputTxt) Then
    If Right$(gDisOutputTxt, 1) <> " " And Left$(p0036, 1) <> " " Then
      gDisOutputTxt = gDisOutputTxt & " " & p0036
    Else
      gDisOutputTxt = gDisOutputTxt & p0036
    End If
  Else
    gDisOutputTxt = p0036
  End If
End Sub

Sub sub07C0(p003A As Long, p003C() As String)
Dim l0042 As Long
Dim l0044
Dim l0048 As Integer
  l0042 = 1
  For l0044 = 0 To UBound(p003C)
    If p003A And l0042 Then
      GoSub L12F5E
      DisOutput p003C(l0044)
      p003A = p003A Xor l0042
    End If
    l0042 = l0042 * 2
  Next
  If p003A Then
    GoSub L12F5E
    DisOutput "0x" & Hex$(p003A)
  End If
Exit Sub

L12F5E:
  If l0048 Then
    DisOutput " | "
  Else
    DisOutput " ": l0048 = True
  End If
Return
End Sub

Sub sub07CF(Lic_Data As Integer, RegC As Long, p0018 As Long, p001A As Long)
  p0018 = p001A
  Seek Lic_Data, RegC + p001A
End Sub

