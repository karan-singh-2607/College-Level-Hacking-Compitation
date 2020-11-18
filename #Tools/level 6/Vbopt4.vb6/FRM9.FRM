VERSION 5.00
Begin VB.Form frmEXE 
   Caption         =   "*.EXE"
   ClientHeight    =   1455
   ClientLeft      =   2805
   ClientTop       =   1590
   ClientWidth     =   8520
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   1455
   ScaleWidth      =   8520
   Begin VB.Frame FrameDebug 
      Caption         =   "debug"
      Height          =   795
      Left            =   3180
      TabIndex        =   6
      Top             =   0
      Width           =   4995
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "refresh"
         Height          =   255
         Left            =   1500
         TabIndex        =   14
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton CmdDumpMod 
         Caption         =   "dump mod"
         Height          =   255
         Left            =   3660
         TabIndex        =   13
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox chk_showNames 
         Caption         =   "show names"
         Height          =   204
         Left            =   60
         TabIndex        =   12
         Top             =   240
         Value           =   1  'Aktiviert
         Width           =   1344
      End
      Begin VB.CommandButton cmdDump 
         Caption         =   "dump"
         Height          =   315
         Left            =   1500
         TabIndex        =   11
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdScanSub 
         Caption         =   "scan sub"
         Height          =   255
         Left            =   2460
         TabIndex        =   10
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "test"
         Height          =   255
         Left            =   2460
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdDecompile 
         Caption         =   "decompile"
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdDumpTok 
         Caption         =   "dump tok"
         Height          =   255
         Left            =   3660
         TabIndex        =   7
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdShowSub 
      Caption         =   "show sub"
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   1140
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdSync 
      Caption         =   "sync"
      Height          =   286
      Left            =   2280
      TabIndex        =   4
      Top             =   60
      Width           =   858
   End
   Begin VB.ListBox ListBox1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      MousePointer    =   1  'Pfeil
      TabIndex        =   3
      Top             =   780
      Width           =   2640
   End
   Begin VB.CommandButton cmdInit 
      Caption         =   "init"
      Height          =   253
      Left            =   121
      TabIndex        =   2
      Top             =   1100
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox ComboBox1 
      Height          =   288
      Left            =   120
      Style           =   2  'Dropdown-Liste
      TabIndex        =   1
      Top             =   420
      Width           =   3036
   End
   Begin VB.ComboBox ComboBox2 
      Height          =   288
      Left            =   120
      Style           =   2  'Dropdown-Liste
      TabIndex        =   0
      Top             =   60
      Width           =   2064
   End
End
Attribute VB_Name = "frmEXE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frm9.FRM
Option Explicit
Dim mVAToRaw As Long
Const mc003A = 8 ' &H8%
Const mc0084 = 2 ' &H2%
Const mc0086 = 1 ' &H1%
Const mc00F8 = 4096 ' &H1000%
Const mc0128 = 35 ' &H23%
Const mc0142 = 8 ' &H8%
Const mc0144 = 33 ' &H21%

Sub cmdDecompile_Click()
  frmDecompiler.Show
  frmDecompiler.control1(2).value = True
  frmStatements.cmd_init.value = True
End Sub

Sub cmdDumpTok_Click()
  ShowMsgBox "Feature not available in demo version"
End Sub

Sub CmdDumpMod_Click()
  ShowMsgBox "Feature not available in this version"
End Sub

Sub cmdDump_Click()
'  ShowMsgBox "Feature not available in this version"
   Form1.Show
End Sub

Sub cmdinit_Click()
  
  frmScan.control5 = "Loading"
  frmScan.control4 = gExeFileNameOnly2
  frmScan.control3 = "Scanning"
  frmScan.control2 = "Fixups"
  frmScan.SetFocus
  DoEvents
  
  gv045E = 0
  LoadExeStrucs
  
  Caption = gExeFileNameOnly2 & Right$(G_FullExeFileName, 4)
  sub0CA7
  sub0620 "selected source line will display here"
  
  If gv045E >= mc003A Then
    Unload Me
    Exit Sub
  End If
  
  sub0E62 ComboBox2
  ComboBox2.SetFocus
  ComboBox2.ListIndex = ComboBox2.ListCount = 0
End Sub


Sub CmdRefresh_Click()
  gv16BC = 0
  ComboBox1_Click
End Sub

Sub cmdShowSub_Click()
  sub2B41 cmdShowSub.HelpContextID
End Sub

Sub cmdSync_Click()
Dim l004C As Integer
Dim l004E As Integer
Dim l0050 As Integer
  If gv097C = 0 Then ErrBeep2: Exit Sub
  l0050 = frm3.control5.ListIndex
  If l0050 < 0 Then Exit Sub
  l004E = frm3.control5.ItemData(l0050)
  If l004E <= 0 Then Exit Sub
  l004C = gv0AD6(l004E).M404D
  If l004C = 0 Then Exit Sub
  If l004C > 0 Then
    sub2B41 l004C
  End If
End Sub

Sub ComboBox2_Click()
Dim l0060 As Integer
  ComboBox1.Clear
  ListBox1.Clear
  l0060 = ComboBox2.ListIndex
  If l0060 < 0 Then Exit Sub
  l0060 = ComboBox2.ItemData(l0060)
  sub19F0 l0060, True
  sub0E72 ComboBox1, l0060, 0
  If ActiveControl = ComboBox2 Then
    ComboBox1.SetFocus
    ComboBox1.ListIndex = ComboBox1.ListCount = 0
  End If
  sub110E
End Sub

Sub ComboBox1_Click()
Dim l0076 As Integer
  l0076 = ComboBox1.ListIndex
  If l0076 < 0 Then Exit Sub
  l0076 = ComboBox1.ItemData(l0076)
  sub028D l0076, ""
End Sub

Sub sub2B62(p0078%, p007A%)
Dim l007C
Dim l007E As Long
Static m0082 As Integer
' Const mc0084 = 2 ' &H2%
' Const mc0086 = 1 ' &H1%
  sub1860 p0078, p007A
  Select Case gv34F4
  Case m0082
  Case mc0084
    MsgBox "no more errors"
    ComboBox1.SetFocus
  Case mc0086
    MsgBox "section not scanned"
  Case Else
    DbgMileStone "da ist was passiert"
    DbgMileStone "assert sub shown"
    l007E = gv0462
    sub2B41 p0078
    For l007C = ListBox1.ListCount - 1 To 1 Step -1
      If ListBox1.ItemData(l007C) <= gv0462 Then
        ListBox1.ListIndex = l007C
        ListBox1.SetFocus
        Exit Sub
      End If
    Next
    DbgMileStone "line not found!?"
    Exit Sub
  End Select
  ComboBox1.SetFocus
End Sub

Sub sub2B41(ByVal pv008A As Integer)
Dim l008C As Integer
Dim l008E
  l008C = gEXEdef_Subs(pv008A).ModuleIndex
  ComboBox2.ListIndex = l008C - 1
  DoEvents
  For l008E = ComboBox1.ListCount - 1 To 0 Step -1
    If ComboBox1.ItemData(l008E) = pv008A Then
      ComboBox1.SetFocus
      ComboBox1.ListIndex = -1
      ComboBox1.ListIndex = l008E
      Exit Sub
    End If
  Next
  ErrBeep
End Sub

Sub Form_Load()
  Icon = frmMain.Icon
'  FrameDebug.Visible = False
'  If gv097C <= 0 Then
'    cmdSync.Visible = False
'  End If
End Sub

Sub Form_Paint()
Static mMarkerFontSet As Integer
  If mMarkerFontSet <> gMarkerFontSet Then
    ListBox1.FontSize = gFontSize
    ListBox1.FontBold = gFontBold
    mMarkerFontSet = gMarkerFontSet
  End If
End Sub

Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End Sub

Sub Form_Resize()
  If WindowState = 1 Then Exit Sub
  ListBox1.Width = ScaleWidth - ListBox1.Left
  ListBox1.Height = ScaleHeight - ListBox1.Top
End Sub

Sub Form_Unload(Cancel As Integer)
  UnloadMe
End Sub

Sub ListBox1_Click()
Dim l00CE As Integer
  l00CE = ListBox1.ListIndex
  If l00CE < 0 Then Exit Sub
  sub0620 (ListBox1.List(l00CE))
End Sub

Sub ListBox1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ListBoxItemIndex As Integer
Dim CurMetaCodeLineNr As Integer
' Const mc00F8 = 4096 ' &H1000%

'MouseButtonConstants
'Const vbRightButton& = 2
'Const vbLeftButton& = 1
  
  Select Case Button
  Case vbRightButton
    Set gv02D6 = ListBox1
    PopupMenu frmMain.Menu_popup
    
  Case vbLeftButton
    
    ListBoxItemIndex = ListBox1.ListIndex
    If ListBoxItemIndex < 0 Then Exit Sub
    
    CurMetaCodeLineNr = 0
    
    'frmDecompiler.Init(clear)
    frmDecompiler.control1(0).value = True
    
   'Get VA Offset from Line and validate it
    gModuleVACur = ListBox1.ItemData(ListBoxItemIndex)
    If (gModuleVACur < gModuleVAStart) Or (gModuleVACur >= gModuleVAEnd) Then
      Exit Sub
    End If
    
    gVACmdStart = 0
    
   'Test VitualAddress To Raw Value
    If gVAToRaw2 <> mVAToRaw Then
      ErrBeep
      gVAToRaw2 = mVAToRaw
    End If
    
    If gvArrSize = 0 Then Exit Sub
    
   'Let Line inside Decompiler
    CurMetaCodeLineNr = Translate_MetaCodeLines_VAToLineNr(gModuleVACur)
    If CurMetaCodeLineNr = 0 Then
      Exit Sub
    End If
    
    ListBoxItemIndex = CurMetaCodeLineNr
    Do
      ListBoxItemIndex = fn1D6C(ListBoxItemIndex)
    Loop Until (ListBoxItemIndex = 0) Or ((gMetaCodeLines(ListBoxItemIndex).M36B3_count And mc00F8) = 0)
    
    gVACmdEnd = gMetaCodeLines(ListBoxItemIndex).VOffset
    If gVACmdEnd > gModuleVACur Then gModuleVACur = gVACmdEnd
    
    gVACmdEnd = gMetaCodeLines(CurMetaCodeLineNr).VOffset
    
    sub00E9
    
    frmDecompiler.Show
    
   'Fill in MetaCode into frmDecompiler
    Do
        gVACmdStart = gVACmdEnd
        gDisOutputTxt = fn18AB(gVACmdEnd)
        CurMetaCodeLineNr = Len(gDisOutputTxt)
        
        frmDecompiler.control1(1).value = True
        
        If gv34F4 Then Exit Do
        If (CurMetaCodeLineNr > 0) And (Left$(gDisOutputTxt, 1) <> "'") Then Exit Do
    Loop
    
    If gv34F4 > mc0084 Then
      sub1892 frmDecompiler.Lst_OutPut
    End If
  
  End Select

End Sub

Sub sub2BA6(p0100%)
  sub2B62 gv0468 + 1, p0100
End Sub

Sub sub2BB3(p0104%)
Dim l0106 As Integer
Dim l0108 As Integer
  l0106 = ComboBox1.ListIndex
  If l0106 < 0 Then l0106 = 0
  l0108 = ComboBox1.ItemData(l0106)
  If l0108 < 0 Then l0108 = g_decl_objs(l0106).M3F1B + 1
  sub2B62 l0108, p0104
End Sub

Sub sub028D(p0112 As Integer, ByVal p0114 As String)
Dim l0116 As Integer
Dim l011E As Integer
' Const mc0128 = 35 ' &H23%
' Const mc0142 = 8 ' &H8%
' Const mc0144 = 33 ' &H21%
  If Len(p0114) Then
    sub079A p0114
  Else
    SetDecompilerListBox ListBox1
  End If
  If p0112 <= 0 Then
    gv16BC = p0112
    l0116 = -p0112
    Select Case gReloc_Arr(l0116).M36B3_count
    Case mc0128
If gIsVB_Ver Then
      sub0115 fn202F(l0116)
Else
      sub0D4A fn208C(fn202F(l0116))
End If
If gIsVB5 Then
      DecompFrmDesign gVAToRaw2, fn202F(l0116), g_decl_objs(gv16BA).Name_M411A
Else
      sub21A1 gVAToRaw2, fn202F(l0116), g_decl_objs(gv16BA).Name_M411A
End If
    Case Else
      DisOutput "<unknown format>": WriteOutputTxt
    End Select
  ElseIf p0112 > gEXEdef_Arr5Count Then
    ErrBeep2
    Exit Sub
  Else
    sub0487 p0112, True
    sub109D chk_showNames
    l011E = gEXEdef_Subs(p0112).lOffsetSym_M3C6A
If gIsVB5 < 0 Then
    Do
      DisOutput fn2023(l011E) & " '" & fn0226(l011E)
      WriteOutputTxt
      sub0126 l011E, fn202F(l011E), fn202F(fn2052(l011E)), mc0142
      l011E = fn2052(l011E)
    Loop Until gReloc_Arr(l011E).M36B3_count <> mc0142
Else
    sub0126 l011E, (gModuleVAStart), gModuleVAEnd, mc0144
End If
    If Len(p0114) = 0 Then sub110E
If gIsVB5 Then
Else
    cmdDecompile_Click
End If
    ComboBox1.SetFocus
  End If
  UnlinkDecompilerListBox
  mVAToRaw = gVAToRaw
End Sub

Sub chk_showNames_Click()
  ComboBox1_Click
End Sub
