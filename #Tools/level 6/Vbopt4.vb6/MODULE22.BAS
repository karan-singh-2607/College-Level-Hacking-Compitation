Attribute VB_Name = "MODULE22"
' Module22
Option Explicit
Dim m0006 As Long
Const mc002E = 16 ' &H10%
Const mc0034 = 18 ' &H12%

Sub NE_GetResource()
Dim l000C As T54E0
Dim l0010 As T54F0
Dim l0012 As Long
Dim l0016
Dim l001C As String
Dim l001E As Integer
Dim l0020 As Long
Dim l0022 As String
' Const mc002E = 16 ' &H10%
' Const mc0034 = 18 ' &H12%
  InitResObj
  If gNE_Header.RESOURCETABLE = gNE_Header.ResidentNameTable Then Exit Sub
  l0012 = gMZ_Header.OffsetToPE + gNE_Header.RESOURCETABLE + 1
  Get g_hFileExe, l0012, l0016: m0006 = 2 ^ l0016
  l0020 = Seek(g_hFileExe)
  Do
    Get g_hFileExe, l0020, l000C: l0020 = l0020 + Len(l000C)
    If l000C.M36B3_count = 0 Then Exit Do
    l001E = l000C.M36B3_count And &H7FFF
    If l000C.M36B3_count < 0 And l001E <= mc002E Then
      l0022 = ""
    Else
      Get g_hFileExe, l0012 + l001E, gSingleChar
      l0022 = Space$(Asc(gSingleChar))
      Get g_hFileExe, , l0022
      If Left$(l0022, 5) = "_IID_" Then
        l001E = mc0034
      Else
        l001E = fn0381(l000C.M36B3_count, l0022)
      End If
    End If
    l0016 = l000C.M54E9
    While l0016 > 0
      Get g_hFileExe, l0020, l0010: l0020 = l0020 + Len(l0010)
      If l0010.ResObjName < 0 Then
        l001C = l0022
      Else
        Get g_hFileExe, l0012 + l0010.ResObjName, gSingleChar
        l001C = Space$(Asc(gSingleChar))
        Get g_hFileExe, , l001C
        If Len(l0022) Then
          l001C = l0022 & "." & l001C
        End If
      End If
      sub0426 l001E, l0010.ResObjName, l001C, ToUnSigned(l0010.hFile) * m0006, ToUnSigned(l0010.M54F9) * m0006, 0
      l0016 = l0016 - 1
    Wend
  Loop
End Sub
