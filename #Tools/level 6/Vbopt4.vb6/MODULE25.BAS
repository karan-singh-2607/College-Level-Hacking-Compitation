Attribute VB_Name = "MODULE25"
' Module25
Option Explicit
Const mc000A = 411 ' &H19B%
Const mc000C = 420 ' &H1A4%
Const Lic_FileExt = "se2"
Const mc0074 = 4 ' &H4%
Const mc0076 = 32 ' &H20%
Const mc0078 = 7 ' &H7%
Const mc00CA = 16 ' &H10%
Const mc00E6 = 3 ' &H3%
Const mc00F0 = -32766 ' &H8002%
Const mc00F2 = -32767 ' &H8001%

Function fn171A(p0052 As Integer, p0054 As Integer, p0056 As Integer) As Integer
Static m0058 As Integer
Static m005A As String
Static m005E As Integer
Static m0060 As Integer
Static m0062 As Long
Dim l0066 As Long
' Const mc0074 = 4 ' &H4%
' Const mc0076 = 32 ' &H20%
' Const mc0078 = 7 ' &H7%
  If p0056 = gc16A2 Then
    If m005E Then Close m005E: m005E = 0
    m0058 = 0
    m0062 = 1
    Exit Function
  End If
  
  If p0056 = gv16A0 Then
    If m0058 = gc16AA Then
      If m005E Then Close m005E: m005E = 0
    ElseIf m0058 = gc16AE Then
      m0062 = Seek(m005E)
    End If
    If m005E = 0 Then
      m0058 = 0
    End If
    fn171A = m005E
    Exit Function
  End If
  If p0056 <> m0058 Then
    If m005E Then Close m005E: m005E = 0
    m0058 = 0
    gFileName_dat = ""
  Else
    gFileName_dat = m005A
  End If
  If p0056 And gc16A4 Then
    m005A = CreateFullFileName("", "t0")
  ElseIf p0056 And gc16AC Then
    m005A = CreateFullFileName("", Lic_FileExt)
  ElseIf p0056 And gc16A8 Then
    m005A = CreateFullFileName("", "m" & Format$(p0052))
  End If
  If gFileName_dat <> m005A Then
    If m005E Then Close m005E: m005E = 0
    gFileName_dat = m005A
  End If
  Select Case p0056
  Case gc16A4
    If m005E = 0 Then
      m005E = FreeFile
      Open m005A For Input As m005E
    End If
    If p0054 = 0 Then
      p0054 = gv094A(p0052).M3F1B
    End If
    Seek m005E, gv0AD6(p0054 - 1).M4045
  Case gc16A6
    If p0052 <= 1 Then
      If m005E Then Close m005E: m005E = 0
      If Len(Dir$(m005A)) Then Kill m005A
      gv0AD6(0).M4045 = 1
      m005E = FreeFile
      Open m005A For Output As m005E
    Else
      m005E = FreeFile
      Open m005A For Append As m005E
    End If
  Case gc16A8
    If p0052 = 0 Then
      If Len(Dir$(m005A)) Then
If g_IsDemoVersion Then
        m0060 = False
Else
        If MsgBox("Use Existing Definitions?", mc0074 Or mc0076) = mc0078 Then
          m0060 = False
        Else
          m0060 = True
        End If
End If
        If Not m0060 Then
          m005A = CreateFullFileName("", Lic_FileExt)
          If Len(Dir$(m005A)) Then Kill m005A
          Exit Function
        End If
      End If
    End If
    If m005E = 0 Then
      m005E = FreeFile
      Open m005A For Binary Access Read As m005E
    End If
    Seek m005E, 1
  Case gc16AA
    If p0052 = 0 Then m0060 = True
    If m005E Then Close m005E Else m005E = FreeFile
    If Len(Dir$(m005A)) Then Kill m005A
    Open m005A For Binary Access Write As m005E
  Case gc16AC
    If Not m0060 Then Exit Function
    If m005E = 0 Then
      m005E = FreeFile
      Open m005A For Binary Access Read As m005E
    End If
    l0066 = gv0AD6(p0054).M4054
    If l0066 <= 0 Then
      fn171A = 0
      Exit Function
    End If
    Seek m005E, l0066
  Case gc16AE
    If m005E = 0 Then
      m005E = FreeFile
      Open m005A For Binary As m005E
      m0062 = LOF(m005E) + 1
    End If
    gv0AD6(p0054).M4054 = m0062
    Seek m005E, m0062
  End Select
  m0058 = p0056
  fn171A = m005E
End Function

Function CreateFullFileName(ByVal FileName As String, ByVal Extension As String) As String
Dim FileNameDotPos As Integer
  
 'Get ProjectDir from User if it is not set(= "")
  If gProjectDirPath = "" Then
    If Len(gv_ExeFilePath) = 0 Then gv_ExeFilePath = g_ExeFilePath
    
    'Get ProjectDir from User
    frmProjectDir.DriveListBox1.Drive = g_ExeFilePath
    frmProjectDir.DirListBox1.Path = g_ExeFilePath
    frmProjectDir.TextBox1 = gExeFileNameOnly2
  ''frmProjectDir.Show 1
    
    gProjectDirPath = g_ExeFilePath
  
  End If
  
 'Test if FileName is not empty
  If Len(FileName) Then
   'Split away extension
    FileNameDotPos = InStr(FileName, ".")
    If FileNameDotPos Then
      CreateFullFileName = gProjectDirPath & Left$(FileName, FileNameDotPos) & Extension
    Else
      CreateFullFileName = gProjectDirPath & FileName & "." & Extension
    End If
  
  Else
    '... well FileName is empty
    CreateFullFileName = gProjectDirPath & gExeNameOnly & "." & Extension
  
  End If
End Function

Sub sub0A5E(p00EE As Integer)
' Const mc00F0 =-32766 ' &H8002%
' Const mc00F2 =-32767 ' &H8001%
Dim l00F4 As Integer
Dim l00F6
' Const mc00F8 = 4 ' &H4%
  If gv16B6 = p00EE Then Exit Sub
  If gv16B8 Then sub0BD8
  sub1732 gv0AD6(p00EE).M403B
  sub0BB9 gv0AD6(p00EE).M404D
  Select Case gv0AD6(p00EE).Flags
  Case mc00F0
    gv16B6 = p00EE
    Exit Sub
  Case mc00F2
    Exit Sub
  End Select
  l00F4 = fn171A(gv16B2, p00EE, gc16AC)
  If l00F4 = 0 Then GoTo L2A69C
  Get l00F4, , l00F6
  If l00F6 <> mc000C Then
    MsgBox gFileName_dat, , "wrong module version"
    GoTo L2A69C
  End If
  Get l00F4, , gv371E: Get l00F4, , gv3718
  Get l00F4, , gv371C: Get l00F4, , gv371A
  Get l00F4, , gv36E4: ReDim gv36AE(gv36E4)
  For l00F6 = 1 To gv36E4
    Get l00F4, , gv36AE(l00F6)
  Next
  Get l00F4, , l00F6: gv36E0 = Space$(l00F6)
  Get l00F4, , gv36E0
  gv16B6 = p00EE

L2A69C:
  l00F4 = fn171A(gv16B2, p00EE, gv16A0)
  If p00EE <> gv16B6 Then
    sub0BAB p00EE, &H4
  Else
    sub1740
  End If
End Sub

Sub sub0AF5()
Dim l0090 As Integer
Dim l0092
Dim l0094
  gv16B0 = True
  l0090 = fn171A(0, 0, gc16A8)
  If l0090 = 0 Then Exit Sub
  Get l0090, 1, l0092
  Select Case l0092
  Case 0
  Case 2
    GoSub L29CDE
  Case 3
    GoSub L29CA4
  Case Else
    MsgBox gFileName_dat, , "wrong version"
  End Select
  l0090 = fn171A(0, 0, gv16A0)
Exit Sub

L29CA4:
  Get l0090, , gv097C: ReDim gv094A(gv097C)
  Get l0090, , gv0980
  GoTo L29D0A

L29CDE:
  Get l0090, , gv097C: ReDim gv094A(gv097C)
  gv0980 = 0
L29D0A:
  Get l0090, , gv0982
  Get l0090, , gv0B08: ReDim gv0AD6(gv0B08)
  Get l0090, , gv0B0A
  Get l0090, , gv0B0C
  Get l0090, , gv38EC: ReDim gv38BA(gv38EC)
  For l0092 = 1 To gv38EC
    Get l0090, , gv38BA(l0092)
  Next
  Get l0090, , gv3922: ReDim gv38F0(gv3922)
  For l0092 = 1 To gv3922
    Get l0090, , gv38F0(l0092)
  Next
  Get l0090, , gv360A: ReDim gv35D4(gv360A)
  For l0092 = 1 To gv360A
    Get l0090, , gv35D4(l0092)
  Next
  Get l0090, , l0092: gv3606 = Space$(l0092)
  Get l0090, , gv3606
  For l0094 = 1 To gv097C
    Get l0090, , gv094A(l0094)
  Next
  For l0092 = 1 To gv0B08
    Get l0090, , gv0AD6(l0092)
  Next
  gv0AD6(0).M4045 = 1
  gv16B0 = False
Return
End Sub

Sub sub0B23()
Dim l00E8 As Integer
Dim l00EA As Integer
Dim l00EC As Integer
  If Not gv16B4 Then Exit Sub
  l00E8 = gv16B2
  If l00E8 = 0 Then MsgBox "cannot save mod0", mc00CA: Exit Sub
  l00EA = fn171A(l00E8, 0, gc16AA)
  l00EC = mc000A: Put l00EA, , l00EC
  Put l00EA, , gv094A(l00E8).M3F13
  Put l00EA, , gv0984
  Put l00EA, , gv0988
  Put l00EA, , gv3676
  For l00EC = 1 To gv3676
    Put l00EA, , gv3640(l00EC)
  Next
  l00EC = Len(gv3672): Put l00EA, , l00EC
  Put l00EA, , gv3672
  l00EA = fn171A(l00E8, 0, gv16A0)
  gv16B4 = False
End Sub

Sub sub0B31()
Dim l00C2 As Integer
Dim l00C4 As Integer
Dim l00C6
  If Not gv16B0 Then Exit Sub
  l00C2 = fn171A(0, 0, gc16AA)
  l00C4 = 3: Put l00C2, 1, l00C4
  Put l00C2, , gv097C
  Put l00C2, , gv0980
  Put l00C2, , gv0982
  Put l00C2, , gv0B08
  Put l00C2, , gv0B0A
  Put l00C2, , gv0B0C
  Put l00C2, , gv38EC
  For l00C4 = 1 To gv38EC
    Put l00C2, , gv38BA(l00C4)
  Next
  Put l00C2, , gv3922
  For l00C4 = 1 To gv3922
    Put l00C2, , gv38F0(l00C4)
  Next
  Put l00C2, , gv360A
  For l00C4 = 1 To gv360A
    Put l00C2, , gv35D4(l00C4)
  Next
  l00C4 = Len(gv3606): Put l00C2, , l00C4
  Put l00C2, , gv3606
  For l00C6 = 1 To gv097C
    Put l00C2, , gv094A(l00C6)
  Next
  For l00C4 = 1 To gv0B08
    Put l00C2, , gv0AD6(l00C4)
  Next
  l00C2 = fn171A(0, 0, gv16A0)
  gv16B0 = False
End Sub

Sub sub0BD8()
Dim l002C As Integer
Dim l002E As Integer
Dim l0030 As Integer
Dim l0032 As Integer
  If Not gv16B8 Then Exit Sub
  l002C = gv16B6
  l0032 = gv0AD6(l002C).M403B
  If l0032 = 0 Then
    Exit Sub
  End If
  l002E = fn171A(l0032, l002C, gc16AE)
  l0030 = mc000C: Put l002E, , l0030
  Put l002E, , gv371E: Put l002E, , gv3718
  Put l002E, , gv371C: Put l002E, , gv371A
  Put l002E, , gv36E4
  For l0030 = 1 To gv36E4
    Put l002E, , gv36AE(l0030)
  Next
  l0030 = Len(gv36E0): Put l002E, , l0030
  Put l002E, , gv36E0
  l002E = fn171A(l0032, l002C, gv16A0)
  gv16B8 = False
End Sub

Sub sub1732(p00C8 As Integer)
' Const mc00CA = 16 ' &H10%
Dim l00CC As Integer
Dim l00CE
' Const mc00E6 = 3 ' &H3%
  If p00C8 = gv16B2 Then Exit Sub
  If p00C8 = 0 Then MsgBox "cannot load mod0", mc00CA: Exit Sub
  If gv16B4 Then sub0B23
  l00CC = fn171A(p00C8, 0, gc16A8)
  If l00CC Then
    Get l00CC, , l00CE
    If l00CE = 0 Then GoTo L2A30A
    If l00CE <> mc000A Then MsgBox gFileName_dat, , "wrong version": GoTo L2A30A
    Get l00CC, , l00CE
    If l00CE <> gv094A(p00C8).M3F13 Then MsgBox gFileName_dat, , "wrong modfile": GoTo L2A30A
    Get l00CC, , gv0984
    Get l00CC, , gv0988
    Get l00CC, , gv3676: ReDim gv3640(gv3676)
    For l00CE = 1 To gv3676
      Get l00CC, , gv3640(l00CE)
    Next
    Get l00CC, , l00CE
    gv3672 = Space$(l00CE): Get l00CC, , gv3672
    gv16B2 = p00C8
    gv16B4 = False
L2A30A:
    l00CC = fn171A(p00C8, 0, gv16A0)
  End If
  If gv16B2 <> p00C8 Then
    sub0B13 p00C8, mc00E6
  End If
End Sub

