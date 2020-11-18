Attribute VB_Name = "MODULE3"
' Module3
Option Explicit
Declare Function extfn02E0 Lib "Kernel" Alias "GetPrivateProfileString" (ByVal p1$, ByVal p2$, ByVal p3$, ByVal p4$, ByVal p5%, ByVal p6$) As Integer
Declare Function extfn02FB Lib "Kernel" Alias "GetSystemDirectory" () As Integer
Declare Function extfn0311 Lib "Kernel" Alias "GetWindowsDirectory" () As Integer
Declare Sub extsub0328 Lib "User" Alias "WinHelp" ()

Function fn0333(p0020 As String) As String
Dim l0022 As Integer
Dim l0024 As String
  l0024 = Space$(128)
  l0022 = extfn02E0("Visual Basic" & Chr$(0), p0020 & Chr$(0), "", l0024, 128, "vb.ini")
  fn0333 = Left$(l0024, l0022)
End Function

Sub GetAppPath()
  g_AppPath = AppendPathSlash((App.Path))
End Sub
