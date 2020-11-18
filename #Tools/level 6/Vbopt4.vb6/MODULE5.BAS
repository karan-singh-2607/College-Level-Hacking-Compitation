Attribute VB_Name = "MODULE5"
' Module5
Option Explicit
Const Ver_dat_VB4 = 3416 ' &HD58%
Const mc0056 = "\vbt"
Const mc005A = 20480 ' &H5000%
Const mc005C = 2304 ' &H900%
Const mc0072 = 4 ' &H4%
Const mc0076 = 1 ' &H1%
Const mc00BA = 16384 ' &H4000%
Const mc00BC = 16 ' &H10%

Function fn04E4(p0086 As Integer) As String
Dim l0088 As Integer
Dim l008A As Integer
  l0088 = gDatArr1(p0086).Name
  If l0088 Then
    l008A = InStr(l0088, gDatStrings2, "§") - l0088
    fn04E4 = Mid$(gDatStrings2, l0088, l008A)
  End If
End Function

Function fn0596() As Integer
End Function

Sub DbgMileStone(ByVal p1$)
  'no noparams?
Debug.Print "Milestone: " & p1
End Sub

Sub sub045E()
End Sub

Sub sub046B()
End Sub

Sub sub047B()
End Sub

Sub sub0487(p00E6 As Integer, p00E8%)
  sub0299 p00E6, p00E8
End Sub

Sub UnloadMe()
Dim l0074 As Integer
  l0074 = fn171A(0, 0, &H1)
  SaveAll True
End Sub

Sub sub04B1(p007A As Integer)
  frmEXE.cmdShowSub.HelpContextID = p007A
  frmEXE.cmdShowSub.Value = True
End Sub

Sub sub04C0()
End Sub

Sub sub04D2()
End Sub

Sub sub04F2()
End Sub

Sub sub0501()
End Sub

Sub sub0511()
End Sub

Sub sub0526()
End Sub

Sub sub053A()
End Sub

Sub sub054A()
End Sub

Sub sub055A()
End Sub

Sub LoadDat(ByVal FileName_part As String)
Dim hFile As Integer
Dim l00AE As String
Dim TwoByte%
Dim NextIdx As Integer
Dim StringSize As Integer
' Const mc00BA = 16384 ' &H4000%
' Const mc00BC = 16 ' &H10%
If gIsVB_Ver Then
  ReDim gDatArrLinks(mc005C)
Else
  ReDim gDatArrLinks(mc005A)
  ReDim gv03C4(mc005A)
End If
  hFile = FreeFile

L1111E:
  gFileName_dat = FileName_part & ".dat"
  If Len(Dir$(gFileName_dat)) = 0 Then
    ShowMsgBox mc00BA
    ReDim gDatArr1(0)
    gDatArrSize = 0
    Exit Sub
  End If
  
  Open gFileName_dat For Binary Access Read As hFile
  
  Get hFile, , TwoByte
  Select Case TwoByte
  Case Ver_dat_VB4
    
   'Get Size for Strings
    Get hFile, , TwoByte: gDatStrings1 = Space$(TwoByte)
    Get hFile, , TwoByte: gDatStrings2 = Space$(TwoByte)
    
   'Get Size for Array
    Get hFile, , gDatArrSize: ReDim gDatArr1(gDatArrSize), gDatArr2(gDatArrSize)
    
   'Get Strings
    Get hFile, , gDatStrings1
    Get hFile, , gDatStrings2
    
    For TwoByte = 1 To gDatArrSize
      Get hFile, , gDatArr1(TwoByte)
      
      NextIdx = gDatArr1(TwoByte).Next
      
      If NextIdx >= 0 Then
        gDatArrLinks(NextIdx) = TwoByte
        Get hFile, , StringSize
        
        gDatArr2(TwoByte) = Input$(StringSize, hFile)
      End If
      
    Next
    gDatArrSize = TwoByte - 1
  Case Else
    MsgBox l00AE, mc00BC, "wrong version"
  End Select

L1131C:
  Close hFile
End Sub

Sub sub0579()
  ShowMsgBox "Feature not available in this version"
End Sub

Sub StartupInit()
  GetAppPath
  ChDrive App.Path
  ChDir App.Path
  InitResObj
End Sub

Sub File_Name(ByVal FullPath As String)
  Dim tmpPos As Integer
  tmpPos = FindLastOccurenceOf(FullPath, "\")
  If tmpPos Then
    g_ExeFilePath = Left(FullPath, tmpPos)
    FullPath = Mid$(FullPath, tmpPos + 1)
  End If
  
  tmpPos = InStr(FullPath, ".")
  
  If tmpPos Then FullPath = Left$(FullPath, tmpPos - 1)
  gExeNameOnly = FullPath
End Sub

Sub LoadVBDat()
If gIsVB_Ver = 5 Then
  LoadDat App.Path & mc0056 & "5"
ElseIf gIsVB_Ver = 4 Then
  LoadDat App.Path & mc0056 & "432"
ElseIf g_IsNE Then
  LoadDat App.Path & mc0056 & "416"
Else
  ShowMsgBox mc0072
  Exit Sub
End If
End Sub

