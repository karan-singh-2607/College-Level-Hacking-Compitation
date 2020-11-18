VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "DoDi's VB4 Disassembler"
   ClientHeight    =   5640
   ClientLeft      =   405
   ClientTop       =   1860
   ClientWidth     =   7665
   Icon            =   "FRM1.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox PictureBox 
      Align           =   1  'Oben ausrichten
      Height          =   592
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   7605
      TabIndex        =   0
      Top             =   0
      Width           =   7664
      Begin VB.TextBox CmdOpenProject 
         Height          =   528
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "FRM1.frx":030A
         Top             =   0
         Width           =   4496
      End
   End
   Begin VB.Menu Menu_File 
      Caption         =   "&File"
      Begin VB.Menu mi_OpenMAK 
         Caption         =   "Open *.MAK"
      End
      Begin VB.Menu mi_OpenEXE 
         Caption         =   "Open *.&EXE"
      End
      Begin VB.Menu mi_Text 
         Caption         =   "&Text..."
      End
      Begin VB.Menu mi_seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mi_Prototype 
         Caption         =   "&Prototype"
      End
      Begin VB.Menu mi_MakeProject 
         Caption         =   "&Make Project"
      End
      Begin VB.Menu mi_Info 
         Caption         =   "&Info..."
      End
      Begin VB.Menu mi_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu Menu_Edit 
      Caption         =   "&Edit"
      Begin VB.Menu mi_Edit_Classes 
         Caption         =   "&Classes"
      End
      Begin VB.Menu mi_Edit_Functions 
         Caption         =   "&Functions"
      End
      Begin VB.Menu mi_Edit_Statements 
         Caption         =   "&Statements"
      End
      Begin VB.Menu mi_Edit_Variables 
         Caption         =   "&Variables"
      End
   End
   Begin VB.Menu Menu_Show 
      Caption         =   "&Show"
      Begin VB.Menu mi_Show_Classes 
         Caption         =   "&Classes"
      End
      Begin VB.Menu mi_Show_Controls 
         Caption         =   "Contro&ls"
      End
      Begin VB.Menu mi_Show_Labels 
         Caption         =   "&Labels"
      End
      Begin VB.Menu mi_Show_Variables 
         Caption         =   "&Variables"
      End
      Begin VB.Menu mi_Show_Resource 
         Caption         =   "&Resource"
      End
      Begin VB.Menu mi_Show_UserTypes 
         Caption         =   "&User Types"
      End
      Begin VB.Menu mi_Show_Subs 
         Caption         =   "Su&bs"
         Visible         =   0   'False
      End
      Begin VB.Menu mi_Show_seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mi_Show_FixedFont 
         Caption         =   "Fixed Font"
      End
   End
   Begin VB.Menu Menu_Debug 
      Caption         =   "&Debug"
      Begin VB.Menu mi_Debug_Segments 
         Caption         =   "&Segments"
      End
      Begin VB.Menu mi_Debug_Tables 
         Caption         =   "&Tables"
      End
      Begin VB.Menu mi_Debug_ModuleTables 
         Caption         =   "&Module Tables"
      End
      Begin VB.Menu mi_Debug_Tokens 
         Caption         =   "&Tokens"
      End
      Begin VB.Menu mi_Debug_References 
         Caption         =   "Re&ferences"
      End
      Begin VB.Menu mi_Debug_matchtokens 
         Caption         =   "match tokens"
      End
      Begin VB.Menu mi_Debug_scanSource 
         Caption         =   "scan source"
      End
      Begin VB.Menu mi_Debug_scanSubs 
         Caption         =   "scan subs"
      End
      Begin VB.Menu mi_Debug_scanTokens 
         Caption         =   "scan tokens"
      End
      Begin VB.Menu mi_Debug_ExtSubs 
         Caption         =   "ExtSubs"
      End
      Begin VB.Menu mi_Debug_Statements 
         Caption         =   "Statements"
      End
      Begin VB.Menu mi_Debug_seperator3 
         Caption         =   "-"
      End
      Begin VB.Menu mi_Debug_DebugMode 
         Caption         =   "DebugMode"
      End
      Begin VB.Menu mi_Debug_Demo 
         Caption         =   "Demo"
      End
      Begin VB.Menu mi_Debug_Lite 
         Caption         =   "Lite"
      End
      Begin VB.Menu mi_Debug_Pro 
         Caption         =   "Pro"
      End
      Begin VB.Menu mi_Debug_Voll 
         Caption         =   "Voll"
      End
   End
   Begin VB.Menu Menu_Windows 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
      Begin VB.Menu mi_Wide 
         Caption         =   "&Wide"
      End
   End
   Begin VB.Menu Menu_popup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mi_ToClipboard 
         Caption         =   "To &Clipboard"
      End
      Begin VB.Menu mi_ToCollection 
         Caption         =   "To Co&llection"
      End
      Begin VB.Menu mi_ToWin 
         Caption         =   "To &Win"
      End
      Begin VB.Menu mi_ToFile 
         Caption         =   "To &File"
      End
      Begin VB.Menu mi_ToPrinter 
         Caption         =   "To &Printer"
      End
      Begin VB.Menu mi_Font 
         Caption         =   "&Font"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frm1.FRM
Option Explicit
Dim m001A As Integer
Const mc00D4 = 8 ' &H8%

Sub mi_Info_Click()
  frmAbout.Show 1
End Sub

Sub mi_Wide_Click()
Dim l0020 As Integer
  l0020 = Not mi_Wide.Checked
  mi_Wide.Checked = l0020
  m001A = 2 + l0020
  Arrange m001A
End Sub

Sub mi_Show_Controls_Click()
  ShowMsgBox "Feature not available in this version"
End Sub

Sub mi_Edit_Classes_Click()
If g_IsDemoVersion Then
  ShowMsgBox "Feature not available in demo version"
Else
  ShowMsgBox "Feature not yet implemented"
End If
End Sub

Sub mi_Edit_Statements_Click()
If g_IsDemoVersion Then
  ShowMsgBox "Feature not available in demo version"
Else
  ShowMsgBox "Feature not yet implemented"
End If
End Sub

Sub mi_Edit_Functions_Click()
If g_IsDemoVersion Then
  ShowMsgBox "Feature not available in demo version"
Else
  ShowMsgBox "Feature not yet implemented"
End If
End Sub

Sub mi_Edit_Variables_Click()
If g_IsDemoVersion Then
  ShowMsgBox "Feature not available in demo version"
Else
  ShowMsgBox "Feature not yet implemented"
End If
End Sub

Sub mi_OpenEXE_Click()
  UnloadMe
  
  If CommandlineMode Then
     G_FullExeFileName = Command$
  Else
     frmOpenFileText.Tag = "*.EXE"
     frmOpenFileText.Show 1
     G_FullExeFileName = frmOpenFileText.Tag
  End If
  
  
  If Len(G_FullExeFileName) Then
    
    gProjectDirPath = ""
    gv_ExeFilePath = ""
    
    Caption = App.Title
    frmScan.Show
    frmScan.SetFocus
    
    Set g_WorkFrm = frmEXE
    frmEXE.cmdInit.value = 1
    frmEXE.Show

    
    Unload frmScan
    
    Arrange m001A
  
  End If
  
  CommandlineMode = False
  
End Sub

Sub mi_Exit_Click()
Dim l0050
  Unload Me
  For l0050 = Forms.Count - 1 To 0 Step -1
    Unload Forms(l0050)
  Next
  Close
  End
End Sub

Sub mi_Show_FixedFont_Click()
  frmFixedFont.Show 1
End Sub

Sub mi_Font_Click()
  ShowMsgBox "Feature not yet implemented"
End Sub

Sub mi_Show_Variables_Click()
  frmScopes.Show
  frmScopes.TextBox = Format$(gLWhatEver + 1)
End Sub

Sub mi_OpenMAK_Click()
  OpenMAKFile
End Sub

Sub mi_Debug_matchtokens_Click()
  MsgBox "Feature not available in this version"
End Sub

Sub mi_MakeProject_Click()
If g_IsDemoVersion Then
  ShowMsgBox "Feature not available in demo version"
Else
  MakeProject True
End If
End Sub

Sub mi_Show_Resource_Click()
If g_IsDemoVersion Then
  ShowMsgBox "Feature not available in demo version"
Else
  frmResources.Show
  frmResources.cmd_Refresh.value = 1
End If
End Sub

Sub mi_Debug_scanSubs_Click()
  MsgBox "Feature not available in this version"
End Sub

Sub mi_Debug_scanTokens_Click()
  MsgBox "Feature not available in this version"
End Sub

Sub mi_Show_Classes_Click()
  ShowMsgBox "Feature not available in this version"
End Sub

Sub mi_Debug_ExtSubs_Click()
  MsgBox "Feature not available in this version"
End Sub

Sub mi_Show_Labels_Click()
If g_IsDemoVersion Then
  ShowMsgBox "Feature not available in this version"
Else
  frm17.Show
End If
End Sub

Sub mi_Debug_ModuleTables_Click()
  MsgBox "Feature not available in this version"
End Sub

Sub mi_Debug_References_Click()
  ShowMsgBox "Feature not available in this version"
End Sub

Sub mi_Debug_Segments_Click()
  ShowMsgBox "Feature not available in this version"
End Sub

Sub mi_Debug_Tables_Click()
  ShowMsgBox "Feature not available in this version"
End Sub

Sub mi_Debug_Tokens_Click()
  ShowMsgBox "Feature not available in this version"
End Sub

Sub mi_Show_UserTypes_Click()
  frmUserDefinedTypes.Show
End Sub

Sub mi_Debug_Statements_Click()
  frmStatements.Show
  frmStatements.cmd_init.value = True
End Sub

Sub mi_ToClipboard_Click()
Dim l0070
Dim l0072 As String
  Clipboard.Clear
  For l0070 = 0 To gv02D6.ListCount - 1
    l0072 = l0072 & gv02D6.List(l0070) & Chr$(13) & Chr$(10)
  Next l0070
  Clipboard.SetText l0072
End Sub

Sub mi_ToCollection_Click()
Dim l007C
Dim l0080 As Form
Dim l0084 As ListBox
If g_IsDemoVersion Then
  ShowMsgBox "Feature not available in demo version"
Else
  If gv02D6 Is Nothing Then Exit Sub
  Set l0080 = gv02D6.Parent
  Set l0084 = frm19.control1
  If l0084 Is gv02D6 Then ErrBeep2: Exit Sub
  If l0084.ListCount Then l0084.AddItem ""
  For l007C = 0 To gv02D6.ListCount - 1
    l0084.AddItem gv02D6.List(l007C)
    l0084.ItemData(l0084.NewIndex) = gv02D6.ItemData(l007C)
  Next
  frm19.Show
End If
End Sub

Sub mi_ToFile_Click()
  DbgMileStone "to file"
End Sub

Sub mi_ToPrinter_Click()
Dim l008A
  For l008A = 0 To gv02D6.ListCount - 1
    Printer.Print gv02D6.List(l008A)
  Next l008A
  Printer.EndDoc
End Sub

Sub mi_ToWin_Click()
Dim l0090
Dim l0094 As Form
Dim l0098 As ListBox
Dim l009C As New frm18
  If gv02D6 Is Nothing Then Exit Sub
  Set l0098 = l009C.control1
  Set l0094 = gv02D6.Parent
  l009C.Move l0094.Left + gv02D6.Left, l0094.Top + gv02D6.Top, gv02D6.Width, gv02D6.Height
  l009C.Caption = gv02D6.List(0)
  l0098.FontName = gv02D6.FontName
  l0098.FontSize = gv02D6.FontSize
  l0098.FontBold = gv02D6.FontBold
  For l0090 = 0 To gv02D6.ListCount - 1
    l0098.AddItem gv02D6.List(l0090)
    l0098.ItemData(l0098.NewIndex) = gv02D6.ItemData(l0090)
  Next
  l009C.Show
End Sub

Sub mi_Text_Click()
Dim hFileText As Integer
On Error Resume Next
  frmOpenFileText.Tag = "*.*"
  frmOpenFileText.Show 1
  
  If Len(frmOpenFileText.Tag) Then
    Unload frmText
    Load frmText
    frmText.Caption = frmOpenFileText.Tag
    hFileText = FreeFile
    
    Open frmOpenFileText.Tag For Input As hFileText
    frmText.control1 = Input$(LOF(hFileText), hFileText)
    Close hFileText
    
    frmText.Show
    Arrange m001A
  End If
End Sub

Sub MDIForm_Load()
  gLic_VerType = "Demo"
  GetLicFromFile "VB4Tools", 40
  Menu_Debug.Visible = False
  DisableMenuControls
  Set g_MainForm = Me
  m001A = 2
  StartupInit
  
  processCommandline
End Sub

Private Sub processCommandline()
   CommandlineMode = Len(Command)
   
   If CommandlineMode Then
      mi_OpenEXE_Click
   End If
   
   CommandlineMode = 0
   
End Sub

Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  UnloadMe
End Sub

Sub MDIForm_Resize()
  If WindowState = 1 Then Exit Sub
  If gv097C Then
    If frmMain.ActiveForm.WindowState <> 0 Then Exit Sub
    Arrange m001A
  Else
    Arrange m001A
  End If
End Sub

Sub MDIForm_Unload(Cancel As Integer)
  End
End Sub

Sub OpenMAKFile()
MousePointer = 11
  UnloadMe
  frmOpenFileText.Tag = "*.MAK;*.VBP"
  frmOpenFileText.Show 1
  If Len(frmOpenFileText.Tag) Then
    g_decl_num = 0
    gProjectDirPath = ""
    gv_ExeFilePath = ""
    g_ExeFilePath = ""
    Caption = App.Title
    frmScan.Show
    frmScan.SetFocus
    gv02EC_MAKName = frmOpenFileText.Tag
    frm3.control4.value = 1
    Assume Len(G_FullExeFileName), "exefile exists"
    Set g_WorkFrm = frmEXE
    g_WorkFrm.cmdDecompile.value = 1
    
    If gv045E < mc00D4 Then
      g_WorkFrm.Show
    End If
    frm3.Show
    Unload frmScan
    Arrange m001A
    DoEvents
  End If
MousePointer = 0
End Sub

Sub PictureBox_Resize()
  If WindowState = 1 Then Exit Sub
  CmdOpenProject.Width = PictureBox.ScaleWidth
End Sub

Sub DisableMenuControls()
  mi_Prototype.Visible = False
  mi_Show_UserTypes.Enabled = True
  mi_Show_Variables.Enabled = True
  mi_Show_Classes.Enabled = False
  mi_Show_Controls.Enabled = False
  mi_Show_Labels.Enabled = False
If g_IsDemoVersion Then
End If
If LicEnableLiteFeatures Then
  mi_Show_Resource.Enabled = True
Else
  mi_Show_Resource.Enabled = False
End If
End Sub
