VERSION 5.00
Begin VB.Form frmBlkFile 
   Caption         =   "BlkFile"
   ClientHeight    =   3930
   ClientLeft      =   12270
   ClientTop       =   540
   ClientWidth     =   3015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   3015
   WindowState     =   1  'Minimiert
   Begin VB.CommandButton cmd_alloc 
      Caption         =   "alloc"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   3420
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmd_locate 
      Caption         =   "locate"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   3060
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox ComboBox1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   0
      Style           =   1  'Einfaches Kombinationsfeld
      TabIndex        =   2
      Top             =   0
      Width           =   2775
   End
   Begin VB.CommandButton cmd_close 
      Caption         =   "close"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3420
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmd_open 
      Caption         =   "open"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   3060
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmBlkFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frm13.FRM
Option Explicit
Const va_ExpectedVer = 96 ' &H60%
Dim Entries As Integer
Dim mFilename As Integer
Dim NewVa As Integer
Dim va_Header As T5628
Dim MoreVaData() As T565C

Sub cmd_alloc_Click()
Dim DataSize As Integer
Dim lComboEntry As Integer
Dim l0066 As Integer
Dim l0068 As Integer
Dim l006A As Long
  If mFilename = 0 Then ErrBeep: Exit Sub
  DataSize = gDataSize
  If DataSize < 1 Then
    gv16C2 = 0
    gv16BE = 0
    Exit Sub
  End If
  
  If DataSize > Entries Then
    
    Entries = DataSize
    ReDim Preserve MoreVaData(Entries)
    
    NewVa = True
    
    lComboEntry = ComboBox1.ListCount
    While lComboEntry < Entries
      ComboBox1.AddItem toHex16(lComboEntry)
      lComboEntry = lComboEntry + 1
    Wend
    
    While lComboEntry > Entries
      lComboEntry = lComboEntry - 1
      ComboBox1.RemoveItem lComboEntry - 1
    Wend
  End If
  
  gv16BE = MoreVaData(DataSize).CurEntries
  l006A = MoreVaData(DataSize).Size
  
  If l006A = gv16C2 Then GoTo L56C0A
  If l006A Then
    l0066 = MoreVaData(DataSize).M54BB
    l0068 = MoreVaData(DataSize).Count
    MoreVaData(l0066).Count = l0068
    MoreVaData(l0068).M54BB = l0066
  End If
  
  If gv16C2 Then
    l0066 = 0
    Do
      gv16BE = MoreVaData(l0066).CurEntries + MoreVaData(l0066).Size
      l0068 = MoreVaData(l0066).Count
      If l0068 = 0 Then
        Exit Do
      End If
      l006A = MoreVaData(l0068).CurEntries - gv16BE
      If l006A >= gv16C2 Then
        Exit Do
      End If
      l0066 = l0068
    Loop
    MoreVaData(DataSize).M54BB = l0066: MoreVaData(l0066).Count = DataSize
    MoreVaData(DataSize).Count = l0068: MoreVaData(l0068).M54BB = DataSize
  Else
    gv16BE = 0
  End If
  MoreVaData(DataSize).CurEntries = gv16BE
  MoreVaData(DataSize).Size = gv16C2
  NewVa = True

L56C0A:
  Seek mFilename, gv16BE
  ComboBox1.List(DataSize - 1) = toHex16(DataSize) & " " & Hex$(gv16BE) & ":" & Hex$(gv16C2)
  ComboBox1.ItemData(DataSize - 1) = gv16BE
  ComboBox1.ListIndex = DataSize - 1
End Sub

Sub cmd_close_Click()
  SaveVA
End Sub

Sub cmd_locate_Click()
Dim l007A As Integer
  l007A = gDataSize
  If l007A <= 0 Then
    DbgMileStone "error blk#"
    gv16BE = 0
    gv16C2 = 0
    Exit Sub
  End If
  
  If l007A > Entries Then
    ReDim Preserve MoreVaData(l007A)
    While Entries < l007A
      Entries = Entries + 1
      ComboBox1.AddItem toHex16(Entries)
    Wend
    NewVa = True
  End If
  
  gv16BE = MoreVaData(l007A).CurEntries
  gv16C2 = MoreVaData(l007A).Size
  If gv16BE > 0 Then Seek mFilename, gv16BE
  ComboBox1.ListIndex = l007A - 1
End Sub

Sub cmd_open_Click()

Dim Filename As String
Dim i
Dim twoBytesTmp As Long
  
  If mFilename Then SaveVA
  
  Filename = gFilename
  mFilename = FreeFile
  
  Open Filename For Binary As mFilename
  Get mFilename, 1, va_Header
  
  If va_Header.Ver = va_ExpectedVer Then
    gDataVersion = va_Header.DataVer
    Entries = va_Header.Entries
    twoBytesTmp = va_Header.Size
  Else
    Entries = 0
    twoBytesTmp = 0
  End If
  
  ReDim MoreVaData(Entries)
  
  If twoBytesTmp Then
    
    Seek mFilename, twoBytesTmp
    
    For i = 0 To Entries
      Get mFilename, , MoreVaData(i)
    Next
    
    For i = 1 To Entries
      If MoreVaData(i).Size Then
        ComboBox1.AddItem toHex16(i) & " " & Hex$(MoreVaData(i).CurEntries) & ":" & Hex$(MoreVaData(i).Size)
      Else
        ComboBox1.AddItem toHex16(i)
      End If
      ComboBox1.ItemData(ComboBox1.NewIndex) = i
    Next
    
    NewVa = False
  
  Else
    
    twoBytesTmp = Len(va_Header) + 1
    va_Header.Ver = va_ExpectedVer
    va_Header.DataVer = gDataVersion
    va_Header.Entries = Entries
    va_Header.Size = twoBytesTmp
    Put mFilename, 1, va_Header
    
    Assume Seek(mFilename) = twoBytesTmp, "all loaded"
    
    MoreVaData(0).CurEntries = 1
    MoreVaData(0).Size = Len(va_Header)
    Put mFilename, , MoreVaData(0)
    
    NewVa = True
    
  End If
  
  i = FindLastOccurenceOf(Filename, "\")
  If i Then Filename = Mid$(Filename, i + 1)
  Caption = Filename
  gDataSize = Entries
  gVAFileName = mFilename
  
End Sub

Sub Form_Load()
  Me.Show
End Sub

Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  SaveVA
End Sub

Sub Form_Resize()
  If WindowState = 1 Then Exit Sub
  ComboBox1.Height = ScaleHeight
  ComboBox1.Width = ScaleWidth
End Sub

Sub SaveVA()
Dim i As Integer
Dim l0092 As Long
  If mFilename = 0 Then Exit Sub
  
  ComboBox1.Clear
  If NewVa Then
    i = MoreVaData(0).M54BB
    l0092 = MoreVaData(i).CurEntries + MoreVaData(i).Size
    va_Header.Entries = Entries
    va_Header.Size = l0092
    Put mFilename, 1, va_Header
    Seek mFilename, l0092
    
    For i = 0 To Entries
      Put mFilename, , MoreVaData(i)
    Next
    NewVa = 0
  End If
  Close mFilename
  mFilename = 0
  gVAFileName = 0
End Sub

