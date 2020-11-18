Attribute VB_Name = "MODULE7"
' Module7
Option Explicit
Const mc001C = 1 ' &H1%
Const mc0044 = 2 ' &H2%
Const mc0046 = 5 ' &H5%
Const mc0048 = 11 ' &HB%
Const mc004A = 12 ' &HC%
Const mc004C = 9 ' &H9%
Const mc004E = 10 ' &HA%
Const mc0050 = 16 ' &H10%
Const mc0052 = 1 ' &H1%
Const mc0054 = 20 ' &H14%
Const mc0056 = 5 ' &H5%
Const mc0058 = 7 ' &H7%
Const mc005A = 28 ' &H1C%
Const mc005C = 10 ' &HA%
Const mc005E = 29 ' &H1D%
Const mc0060 = 13 ' &HD%
Const mc0062 = 15 ' &HF%

Function importsRead(Offset As Long, ImportName$) As Integer
Dim TypeCode As Integer
' Const mc0044 = 2 ' &H2%
' Const mc0046 = 5 ' &H5%
' Const mc0048 = 11 ' &HB%
' Const mc004A = 12 ' &HC%
' Const mc004C = 9 ' &H9%
' Const mc004E = 10 ' &HA%
' Const mc0050 = 16 ' &H10%
' Const mc0052 = 1 ' &H1%
' Const mc0054 = 20 ' &H14%
' Const mc0056 = 5 ' &H5%
' Const mc0058 = 7 ' &H7%
' Const mc005A = 28 ' &H1C%
' Const mc005C = 10 ' &HA%
' Const mc005E = 29 ' &H1D%
' Const mc0060 = 13 ' &HD%
' Const mc0062 = 15 ' &HF%
  TypeCode = 2 'mc0044
  
  Select Case ImportName
  Case "ThunRTMain"
    TypeCode = 5 'mc0046
  Case "ProcCallEngine"
    TypeCode = 11 ' mc0048
  Case "MethCallEngine"
    TypeCode = 12 ' mc004A
  Case "DllFunctionCall"
    TypeCode = 9 ' mc004C
  Case "AssignRecord", "DestructRecord", "CopyRecord"
    TypeCode = mc004E
  Case "GetMem1", "PutMem1"
    TypeCode = mc0050 + mc0052
  Case "GetMem2", "PutMem2"
    TypeCode = mc0054
  Case "GetMem4", "PutMem4"
    TypeCode = mc0050 + mc0056
  Case "GetMem8", "PutMem8"
    TypeCode = mc0050 + mc0058
  Case "GetMemStr", "PutMemStr"
    TypeCode = mc005A
  Case "GetMemVar", "SetMemVar", "PutMemVar"
    TypeCode = mc0050 + mc005C
  Case "GetMemObj", "SetMemObj", "PutMemObj"
    TypeCode = mc0050 + mc005E
  Case "GetMemNewObj", "PutMemNewObj", "SetMemNewObj"
    TypeCode = mc0050 + mc0060
  Case "__vbaErrorOverflow"
    TypeCode = mc0062
  End Select
  
  importsRead = RetOffsetSym(Offset, ImportName, TypeCode)

End Function

Function fn0666(ByVal p0068 As String) As String
Dim l006A As Integer
  l006A = FindLastOccurenceOf(p0068, "\")
  If l006A Then p0068 = Mid$(p0068, l006A + 1)
  l006A = InStr(p0068, ".")
  If l006A Then p0068 = Left$(p0068, l006A - 1)
  fn0666 = p0068
End Function

Sub sub0672()
  gv0508 = 0
  gv053C = 0
  Erase gv050A
End Sub

Sub sub067F(p006E As Integer, ByVal p0070 As String, p0074 As T3AD0)
Dim l0076
Dim l0078 As Integer
Dim l007A
Dim l007C As String
Dim l007E As Integer
Dim l0080
Dim l0082 As Integer
Dim l0084 As Integer
  p0074.Name = fn0666(p0070)
  p0074.M3ADA = gv053C: p0074.M3AE2 = 0
  l0078 = FreeFile
  If Right$(p0070, 4) <> ".ept" Then p0070 = p0070 & ".ept"
  Open g_AppPath & p0070 For Binary Access Read As l0078
  Get l0078, , l0076
  If l0076 = mc001C Then
    Get l0078, , p0074.M3AE2: gv053C = gv053C + p0074.M3AE2
    ReDim Preserve gv050A(gv053C)
    Get l0078, , l0080
    ReDim Preserve gEXEdef_Arr4_imports(gEXEdef_Arr4_importsCount + l0080)
    For l0076 = 1 To p0074.M3AE2
      Get l0078, , gv050A(l0076 + p0074.M3ADA)
    Next
    l0082 = p0074.M3ADA
    For l0076 = 1 To l0080
      Do
        l0082 = l0082 + 1
      Loop Until gv050A(l0082)
      Get l0078, , l007A
      l007C = Space$(l007A)
      Get l0078, , l007C
      Select Case l007C
      Case "_THUNRTMAIN": l007E = mc0046
      Case "_PROCCALLENGINE": l007E = mc0048
      Case "_METHCALLENGINE": l007E = mc004A
      Case "_DLLFUNCTIONCALL": l007E = mc004C
      Case "_ASSIGNRECORD": l007E = mc004E
      Case "_COPYRECORD": l007E = mc004E
      Case "_DESTRUCTRECORD": l007E = mc004E
      Case "__GETMEM1", "__PUTMEM1": l007E = mc0050 + mc0052
      Case "__GETMEM2", "__PUTMEM2": l007E = mc0054
      Case "__GETMEM4", "__PUTMEM4": l007E = mc0050 + mc0056
      Case "__GETMEM8", "__PUTMEM8": l007E = mc0050 + mc0058
      Case "__GETMEMSTR", "__PUTMEMSTR": l007E = mc005A
      Case "__GETMEMVAR", "__SETMEMVAR", "__PUTMEMVAR": l007E = mc0050 + mc005C
      Case "__GETMEMOBJ", "__SETMEMOBJ", "__PUTMEMOBJ": l007E = mc0050 + mc0060
      Case "__GETMEMNEWOBJ", "__SETMEMNEWOBJ", "__PUTMEMNEWOBJ": l007E = mc0050 + mc0060
      Case Else
        Select Case Left$(l007C, 5)
        Case "__GET", "__PUT", "__SET"
          MsgBox l007C, , "ept-typ?"
          l007E = mc0050
        Case Else
          l007E = mc0044
        End Select
      End Select
      l0084 = fn13C6(-p006E, l0082, l007C, l007E)
    Next
  Else
    p0070 = ""
  End If
  Close l0078
End Sub

Sub sub068C(p001E As Integer, p0022 As T5402)
Dim l0024 As Long
Dim l0026 As String
Dim l0028 As Integer
  l0024 = gImportDlls(p001E).ImageBase + p0022.RVA_Offset
  l0026 = p0022.Name
  If l0026 = "" Then
    ErrBeep
    l0026 = gImportDlls(p001E).DllName & "." & Format$(p0022.M5410)
  End If
  l0028 = importsRead(l0024, l0026)
  If p0022.M5410 Then
    Assume p001E = gv21E2, "current ordinals"
    If p0022.M5410 > gv21E4 Then
      gv21E4 = p0022.M5410
      If gv21E4 > UBound(gv21E6) Then ReDim Preserve gv21E6(gv21E4 + 20)
    End If
    gv21E6(p0022.M5410) = l0028
  End If
End Sub

