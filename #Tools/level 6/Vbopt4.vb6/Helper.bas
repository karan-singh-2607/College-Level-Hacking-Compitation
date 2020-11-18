Attribute VB_Name = "Helper"
Option Explicit


'Konstantendeklationen für Registry.cls

'Registrierungsdatentypen
Public Const REG_SZ As Long = 1                         ' String
Public Const REG_BINARY As Long = 3                     ' Binär Zeichenfolge
Public Const REG_DWORD As Long = 4                      ' 32-Bit-Zahl

'Vordefinierte RegistrySchlüssel (hRootKey)
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

Public Const ERROR_NONE = 0


Public Const ERR_FILESTREAM = &H1000000
Public Const ERR_OPENFILE = vbObjectError + ERR_FILESTREAM + 1
Public i, j As Integer

Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (src As Any, ByVal src As Any, ByVal Length&)
Public Declare Sub MemCopyAnyToStr Lib "kernel32" Alias "RtlMoveMemory" (ByVal src As String, Dest As Any, ByVal Length&)
Public Declare Sub MemCopyLngToStr Lib "kernel32" Alias "RtlMoveMemory" (ByVal Dest As String, src As Long, ByVal Length&)

Public Declare Sub MemCopyStrToLng Lib "kernel32" Alias "RtlMoveMemory" (src As Long, ByVal src As String, ByVal Length&)
'Public Declare Sub MemCopyLngToStr Lib "kernel32" Alias "RtlMoveMemory" (ByVal src As String, dest As Long, ByVal Length&)
Public Declare Sub MemCopyLngToInt Lib "kernel32" Alias "RtlMoveMemory" (src As Long, ByVal Dest As Integer, ByVal Length&)
    
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const SM_DBCSENABLED = 42
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Integer) As Integer


'Returns whether the user has DBCS enabled
Private Function isDBCSEnabled() As Boolean
   isDBCSEnabled = GetSystemMetrics(SM_DBCSENABLED)
End Function


Function LeftButton() As Boolean
    LeftButton = (GetAsyncKeyState(vbKeyLButton) And &H8000)
End Function

Function RightButton() As Boolean
    RightButton = (GetAsyncKeyState(vbKeyRButton) And &H8000)
End Function

Function MiddleButton() As Boolean
    MiddleButton = (GetAsyncKeyState(vbKeyMButton) And &H8000)
End Function

Function MouseButton() As Integer
    If GetAsyncKeyState(vbKeyLButton) < 0 Then
        MouseButton = 1
    End If
    If GetAsyncKeyState(vbKeyRButton) < 0 Then
        MouseButton = MouseButton Or 2
    End If
    If GetAsyncKeyState(vbKeyMButton) < 0 Then
        MouseButton = MouseButton Or 4
    End If
End Function

Function KeyPressed(Key) As Boolean
   KeyPressed = GetAsyncKeyState(Key)
End Function



Public Function HexvaluesToString$(Hexvalues$)
   Dim tmpchar
   For Each tmpchar In Split(Hexvalues)
      HexvaluesToString = HexvaluesToString & Chr("&h" & tmpchar)
   Next
End Function

Function Max(ParamArray values())
   Dim item
   For Each item In values
      Max = IIf(Max < item, item, Max)
   Next
End Function

Function Min(ParamArray values())
   Dim item
   Min = &H7FFFFFFF
   For Each item In values
      Min = IIf(Min > item, item, Min)
   Next
End Function

Function limit(value&, Optional ByVal upperLimit = &H7FFFFFFF, Optional lowerLimit = 0) As Long
   'limit = IIf(Value > upperLimit, upperLimit, IIf(Value < lowerLimit, lowerLimit, Value))

   If (value > upperLimit) Then _
      limit = upperLimit _
   Else _
      If (value < lowerLimit) Then _
         limit = lowerLimit _
      Else _
         limit = value
   
End Function

Function RangeCheck(ByVal value&, Max&, Optional Min& = 0, Optional errtext, Optional ErrSource$) As Boolean
   RangeCheck = (Min <= value) And (value <= Max)
   If (RangeCheck = False) And (IsMissing(errtext) = False) Then Err.Raise vbObjectError, ErrSource, errtext
End Function
Public Function H16(ByVal value As Long)
   H16 = Right(String(3, "0") & Hex(value), 4)
End Function

Public Function H32(ByVal value As Long)
   H32 = Right(String(7, "0") & Hex(value), 8)
End Function

Public Function Swap(ByRef A, ByRef B)
   Swap = B
   B = A
   A = Swap
End Function

'////////////////////////////////////////////////////////////////////////
'// BlockAlign_l  -  Erzeugt einen linksbündigen BlockString
'//
'// Beispiel1:     BlockAlign_l("Summe",7) -> "  Summe"
'// Beispiel2:     BlockAlign_l("Summe",4) -> "umme"
Public Function BlockAlign_l(RawString, Blocksize) As String
  'String kürzen lang wenn zu
   RawString = Left(RawString, Blocksize)
  'mit Leerzeichen auffüllen
   BlockAlign_l = Space(Blocksize - Len(RawString)) & RawString
End Function

Public Function BlockAlign_r(RawString, Blocksize) As String
  'String kürzen lang wenn zu
   RawString = Left(RawString, Blocksize)
  'mit Leerzeichen auffüllen
   BlockAlign_r = RawString & Space(Blocksize - Len(RawString))
End Function


Public Function qw()
   Do
      DoEvents
   Loop
End Function
Public Function szNullCut$(zeroString$)
   Dim nullCharPos&
   nullCharPos = InStr(1, zeroString, Chr(0))
   If nullCharPos Then
      szNullCut = Left(zeroString, nullCharPos - 1)
   Else
      szNullCut = zeroString
   End If
   
End Function

Public Function Inc(ByRef value, Optional Increment& = 1)
   value = value + Increment
   Inc = value
End Function

Public Function Dec(ByRef value, Optional DeIncrement& = 1)
   value = value - DeIncrement
   Dec = value
End Function



Public Function CollectionToArray(Collection As Collection) As Variant
   
   Dim tmp
   ReDim tmp(Collection.Count - 1)
   
   Dim i
   i = LBound(tmp)
   
   Dim item
   For Each item In Collection
      tmp(i) = item
      Inc i
   Next
   
   CollectionToArray = tmp
   
End Function
Public Function isString(StringToCheck) As Boolean
   'isString = False
   Dim i&
   For i = 1 To Len(StringToCheck)
      If RangeCheck(Asc(Mid(StringToCheck, i, 1)), &H7F, &H20) Then
      
      Else
         Exit Function
      End If
   Next
   
   isString = True
   
End Function

