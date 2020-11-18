Attribute VB_Name = "MODULE33"
' Module33
Option Explicit
Const Lic_FileExt = ".lic"

Function Lic_CreateChecksum(Lic_Data As String, RegC As Long) As Long
Dim RetVal As Long
Dim Char As Integer

  
  RetVal = 0
  Char = Rnd(-RegC) * &H100
  
  Dim i
  For i = 1 To Len(Lic_Data)
   'Check Char from Lic_Data
    Char = Asc(Mid$(Lic_Data, i, 1))
    
    RetVal = (RetVal And &HFFFFFFF) * 5 + (Char Xor Rnd)
  Next
  
  Lic_CreateChecksum = RetVal
End Function

Sub GetLicFromFile(ByVal ProgName As String, Version As Integer)
Dim hFile As Integer
Dim Lic_FileName As String
Dim SplitPos As Integer
Dim Lic_VerType As String
Dim File_TextLine As String
Dim SplitPos2 As Integer
Dim Lic_FileData As String
Dim Lic_Checksum As Long
Dim Lic_VerTypeInt As Integer
Dim RegA1_Lic_VerTypeInt As Integer
Dim RegA2_Version As Integer
Dim RegB_Lic_Checksum As Long
Dim RegC_NumOfLic As Long
  g_IsDemoVersion = True
  gLic_VerType = "Demo"
  
  Lic_FileName = App.Path & "\" & ProgName & Lic_FileExt
' Quit if there is no 'VB4Tools.lic'
  If Len(Dir$(Lic_FileName)) = 0 Then
    Exit Sub
  End If
  
  gv20F8 = True
  hFile = FreeFile
  Lic_FileData = LCase$(ProgName)
  Open Lic_FileName For Input As hFile
  Do While Not EOF(hFile)
    Line Input #hFile, File_TextLine
    Select Case LCase$(Left$(File_TextLine, 1))
    
    'DoDi's VB4-Tools, Pro Edition
    Case "d"
      SplitPos = InStr(File_TextLine, ",")
      Lic_VerType = Trim$(Mid$(File_TextLine, SplitPos + 1))
      SplitPos = InStr(Lic_VerType, " ")
      If SplitPos Then Lic_VerType = Left$(Lic_VerType, SplitPos - 1)
      gLic_VerType = Lic_VerType
    'Licensed to: CW2K
    Case "l"
      SplitPos = InStr(File_TextLine, ":")
      Lic_Name = Trim$(Mid$(File_TextLine, SplitPos + 1))
      
    'Registration number: 402-738428932-0111
    '                       ^
    '                     ^^
    '                         ^^^^^^^^^^^^^^
    '
    '                         ^^^^^^^^^
    '                                   ^^^^^
    
    Case "r"
      SplitPos = InStr(File_TextLine, "-")
      RegA1_Lic_VerTypeInt = Val(Mid$(File_TextLine, SplitPos - 1))
      RegA2_Version = Val(Mid$(File_TextLine, SplitPos - 3, 2))
      Lic_VerType = Mid$(File_TextLine, SplitPos + 1)
      
      SplitPos2 = InStr(Lic_VerType, "-")
      
     'Cut away Lic_Checksum from Line
      File_TextLine = Left$(File_TextLine, SplitPos) & Mid$(Lic_VerType, SplitPos2)
      
      RegB_Lic_Checksum = Val(Lic_VerType)
      RegC_NumOfLic = Val(Mid$(Lic_VerType, SplitPos2 + 1))
      
   'End of Lic Area reached
    Case ""
      Exit Do
    End Select
    Lic_FileData = Lic_FileData & File_TextLine
  Loop
  
'RegRest Of LicFile
  While Not EOF(hFile)
    Line Input #hFile, File_TextLine
    Lic_FileData = Lic_FileData & File_TextLine
  Wend
  
  Close hFile
  
 'Quit if there are no Licenses
  If RegC_NumOfLic = 0 Then
    Exit Sub
  End If
  
  Lic_Checksum = Lic_CreateChecksum(Lic_FileData, RegC_NumOfLic)
  If Lic_Checksum <> RegB_Lic_Checksum Then
    Exit Sub
  End If
  
  If RegA2_Version <> Version Then
    Exit Sub
  End If
  
  Select Case LCase$(Left$(gLic_VerType, 1))
  Case "l"
    Lic_VerTypeInt = 1
  Case "p"
    Lic_VerTypeInt = 2
  Case Else
    Lic_VerTypeInt = 0
  End Select
  
  If Lic_VerTypeInt <> RegA1_Lic_VerTypeInt Then
    Lic_VerTypeInt = 0
  End If
  
  g_IsDemoVersion = False
  
  Select Case Lic_VerTypeInt
  Case 1
    LicEnableLiteFeatures = True
  Case 2
    LicEnableLiteFeatures = True
    LicEnableProFeatures = True
  Case Else
    g_IsDemoVersion = True
  End Select
  
  gv20F8 = False

End Sub
