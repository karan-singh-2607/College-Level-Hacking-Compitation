Attribute VB_Name = "MODULE20"
' Module20
Option Explicit

Sub sub15EE(p0008 As Long, p000A As Long)
Dim l000C As Long
Dim l000E As Integer
Dim l0010 As Integer
  Assume gImportDLLCount = gv21E2, "current mod"
  Seek g_hFileExe, p000A
  Do
    Get g_hFileExe, , l000C
    If l000C = 0 Then Exit Sub
    l000E = CreateSym(p0008)
    If l000C < 0 Then
      l000C = fn088F(l000C)
      l0010 = gv21E6(l000C)
    Else
      l0010 = CreateSym(l000C)
    End If
    sub13B9 l0010, p0008, 3
    gReloc_Arr(l000E).subRef_M3C6A = -l0010
    p0008 = p0008 + 4
  Loop
End Sub

Sub ApplyRelocChunkEntries(VA_CodeSection As Long, ByVal NumRelocEntrys As Integer)
Dim i%
Dim RelocPart%
Dim l0032 As Integer
ReDim Preserve gv161C(gv164E + NumRelocEntrys)
  
  For i = 1 To NumRelocEntrys
    Get g_hFileExe, , RelocPart
    If RelocPart = 0 Then Exit For
    
    l0032 = ApplyRelocChunkEntry(VA_CodeSection + (RelocPart And &HFFF), 3)
  Next
End Sub

Private Function ApplyRelocChunkEntry(VA_CodeSection As Long, RebaseType%) As Integer
Dim FilePos As Long
Dim AdrThatNeedsRebase As Long
 'Store FilePos
  FilePos = Seek(g_hFileExe)
  
  Get g_hFileExe, VAtoRaw(VA_CodeSection), AdrThatNeedsRebase
  
  Assume AdrThatNeedsRebase > 0, "no ords"
  sub13B9 CreateSym(AdrThatNeedsRebase), VA_CodeSection, RebaseType
  ApplyRelocChunkEntry = gv164E
  
 'Restore FilePos
  Seek g_hFileExe, FilePos
End Function
