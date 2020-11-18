Attribute VB_Name = "MODULE6"
' Module6
Option Explicit
Const mc005C = -32562 ' &H80CE%
Const mc005E = 16640 ' &H4100%
Const mc0060 = 24400 ' &H5F50%
Const vbCritical = 16 ' &H10%
Const mc006C = 48 ' &H30%
Const mc006E = 1 ' &H1%
Const mc0090 = 5 ' &H5%
Const mc0092 = 2 ' &H2%
Const mc00A8 = 4 ' &H4%
Const mc00AA = 2 ' &H2%
Const mc00AC = 3 ' &H3%
Const mc00AE = 7 ' &H7%

Function fn05EA() As Integer
  fn05EA = gv045C
End Function

Sub ShowMsgBox(MsgText As Variant)
' Const vbCritical = 16 ' &H10%
' Const mc006C = 48 ' &H30%
' Const mc006E = 1 ' &H1%
' Const vbString = 8
  If gv0460 = 0 Then
  End If
  
  If VarType(MsgText) = vbString Then
    MsgBox MsgText, vbCritical, "News from VB Discompiler"
    gv045A = -1
  Else
    If gv045C And MsgText Then Exit Sub
    sub05F9 Val(MsgText), mc006C, mc006E
  End If
End Sub

Function fn0604() As Integer
  fn0604 = gv045A
End Function

Sub sub0611()
  gv045A = 0: gv0460 = 0
End Sub

Sub sub05F9(p0072 As Integer, p0074 As Integer, p0076 As Integer)
Dim l0078 As Integer
Dim l007A As String
' Const mc0090 = 5 ' &H5%
' Const mc0092 = 2 ' &H2%
' Const mc00A8 = 4 ' &H4%
' Const mc00AA = 2 ' &H2%
' Const mc00AC = 3 ' &H3%
' Const mc00AE = 7 ' &H7%
  If gv045A = 0 Then gv045A = p0072
  If p0072 And mc005C Then l0078 = l0078 Or gc04A8
  If p0072 And mc005E Then l0078 = l0078 Or gc04A6
  If p0072 And mc0060 Then l0078 = l0078 Or gc04A4
  If l0078 = 0 Then l0078 = gc04A2
  gv045C = gv045C Or p0072
  Select Case p0072
  Case 0
  Case gc047A: l007A = gc002E
  Case gc047C: l007A = gc0032
  Case gc047E: l007A = "VB Version" & Hex$(gVBversion) & gc0036
  Case gc0480: l007A = gc003A
  Case gc0482: l007A = Foundunknownfixups
  Case gc0484: l007A = gc004A
  Case gc0488: l007A = "Feature not yet implemented"
  Case gc0486: l007A = gFileName_dat & gc0052
        p0074 = vbCritical
If g_IsDemoVersion Then
        p0076 = mc0090
Else
        p0076 = mc0092
End If
  Case gc048A: l007A = gc005A
  Case gc048C: l007A = gc005E
  Case gc048E: l007A = gc0062
  Case gc0490: l007A = gc0066
  Case gc0492: l007A = gc006A
  Case gc0494: l007A = gc006E
  Case gc0496: l007A = gc0072 & Chr$(10) & gFileName_dat
  Case gc0498: l007A = gc004E
  Case Else: l007A = gc002A & Format$(p0072)
  End Select
  
  If Len(l007A) Then
    If Right$(l007A, 1) = "!" Then l007A = l007A & Chr$(10) & gc0026
    gv0474 = MsgBox(l007A, p0074 Or p0076, "News from VB Discompiler")
    
    Select Case gv0474
    Case mc00A8
      Exit Sub
    Case mc00AA, mc00AC, mc00AE
      GoTo L119D0
    End Select
  End If
  l0078 = l0078 And Not gv045E
  gv045E = gv045E Or l0078
  Select Case l0078
  Case 0
  Case gc04A8
    MsgBox gc000E, vbCritical, gc000A
    gv0474 = mc00AC
  Case gc04A6
    gv0474 = MsgBox(gc0012, 48 + 1, gc000A)
  Case gc04A4
    gv0474 = MsgBox(gc0016, 32 + 1, gc000A)
  Case Else
    gv0474 = MsgBox(gc001A, 64 + 1, gc000A)
  End Select
  Select Case gv0474
  Case mc00AA, mc00AC, mc00AE
  Case Else
    Exit Sub
  End Select

L119D0:
  gv045E = gv045E Or gc04AA
End Sub

Sub sub0620(ByVal p00BA As String)
  frmMain.CmdOpenProject = p00BA
End Sub
