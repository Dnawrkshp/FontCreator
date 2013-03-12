Attribute VB_Name = "Math"
'Math module for PS2FontCreator/Font Creator
Option Explicit
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'declarations for mathmatical formula variables
Public MaxHexInput As Long
Public MaxDecInput As Long
Public MaxBinInput As Long

Public DecBytes() As Byte
Public MaxDecLen As Long
Public HexString$
Public DecString$
Public BinString$    ' General

Public HexString1$, DecString1$, BinString1$ ' 1st number
Public HexString2$, DecString2$, BinString2$ ' 2nd number
Public HexResult$, DecResult$, BinResult$    ' Logic result
Public HexRemain$, DecRemain$, BinRemain$    ' Div remainder

Public HexBytes() As Byte
Public MaxHexLen As Long
Public Dec1()
Public Dec2()
Public PDec1() As Byte
Public BinBytes() As Byte
Public MaxByteLen As Long
Public BinBits() As Byte
Public MaxBinLen As Long
'public aVBASM As Boolean
Public ptrBinBytes As Long
Public ptrDecBytes As Long
Public ptrHexBytes As Long
Public ptrBinBits As Long
Public Counter() As Byte

'Written by ORCXodus
Function StrToHex(data As String) As String
Dim strTemp   As String
Dim strReturn As String
Dim I         As Long
    For I = 1 To Len(data)
        strTemp = Hex$(Asc(Mid$(data, I, 1)))
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        strReturn = strReturn & strTemp 'Space$(1) & strTemp
    Next I
    StrToHex = strReturn
End Function

'Written by ORCXodus
Function HexToString(ByVal HexToStr As String) As String
Dim strTemp   As String
Dim strReturn As String
Dim I         As Long
    For I = 1 To Len(HexToStr) Step 2
        strTemp = Chr$(Val("&H" & Mid$(HexToStr, I, 2)))
        strReturn = strReturn & strTemp
    Next I
    HexToString = strReturn
End Function

Function Pad(String1 As String, Size As Integer)

Pad = String1
If Len(String1) >= Size Then: Exit Function

Dim X As Integer
X = 1

Do While X <= (Size - Len(String1))
    Pad = "0" & Pad
    X = X + 1
Loop


End Function

'Written by Dnawrkshpwrkshp
Function StringByteReverse(HexString As String)
StringByteReverse = ""
Dim Counter As Integer
Counter = Len(HexString)

'Checks to see if it is even
If Not (Counter Mod 2) = 0 Then Exit Function

Do While Not Counter = 0
StringByteReverse = Mid(HexString, Counter, 1) & Mid(HexString, Counter - 1, 1) & StringByteReverse
Counter = Counter - 2
Loop

End Function

'Written by Dnawrkshpwrkshp
Function StringFlip(Bytes As String)
StringFlip = ""
Dim Counter As Integer
Counter = 1

Do While Counter <= Len(Bytes)

StringFlip = StringFlip & Mid(Bytes, Counter + 6, 2)
StringFlip = StringFlip & Mid(Bytes, Counter + 4, 2)
StringFlip = StringFlip & Mid(Bytes, Counter + 2, 2)
StringFlip = StringFlip & Mid(Bytes, Counter, 2)

Counter = Counter + 8
Loop

End Function

'Written by Dnawrkshpwrkshp
Function HexToLitEdn(Bytes As String)
HexToLitEdn = ""
Dim Counter As Integer
Counter = 1

Do While Counter <= Len(Bytes)

HexToLitEdn = HexToLitEdn & Mid(Bytes, Counter + 2, 2)
HexToLitEdn = HexToLitEdn & Mid(Bytes, Counter + 0, 2)
HexToLitEdn = HexToLitEdn & Mid(Bytes, Counter + 6, 2)
HexToLitEdn = HexToLitEdn & Mid(Bytes, Counter + 4, 2)

Counter = Counter + 8
Loop

End Function
Function BinToDec(Bin As String) As Long
Dim I As Double, Exp As Variant, TotOut As Variant

Exp = CDec(0)
TotOut = CDec(0)
For I = 0 To Len(Bin) - 1
    
    If Mid(Bin, Len(Bin) - I, 1) = "1" Then
        TotOut = CDec(CDec(TotOut) + CDec((2 ^ CDec(Exp))))
    End If
    
    Exp = CDec(CDec(Exp) + 1)
Next I
'BinString$ = TotOut
BinToDec = CDec(TotOut)
End Function
'---------------------------------------------------------------------------------------
' Procedure : Hex2Bin
' Author    : Xodus
' Date      : 9/23/2012
' Purpose   : Converts a hex value to a binary value
'---------------------------------------------------------------------------------------
Function Hex2Bin(HexValue As String)
method1:
    Hex2Bin2Dec (HexValue)
    Hex2Bin = BinString$
   
End Function
'====================================================================================
'====================================================================================
Private Sub Hex2Bin2Dec(A$)
' IN:  A$ = Hex string
' OUT: DecString$, BinString$
Dim k As Long, J As Long
Dim b As Byte
   b = 48
   FillMemory HexBytes(1), MaxHexLen, b    ' "0" to HexBytes()
   ' Fill HexBytes() from A$
   ' NB HexBytes(1) is Right char of A$ ie @ Len(A$)
   ' ie CopyMemory cannot be used here
   For k = 1 To Len(A$)
      HexBytes(k) = Asc(Mid$(A$, Len(A$) - (k - 1), 1))
   Next k
    
   Hex2Bytes A$
   A$ = ""
   Bytes2Bits   ' BinBits()
   Bytes2Dec    ' DecBytes()   ' Slow in VB
   
   ' Get Dec result
   DecString$ = ""
   For k = MaxDecLen To 1 Step -1
      If DecBytes(k) <> 48 Then Exit For
   Next k
   For J = k To 1 Step -1
         DecString$ = DecString$ + Chr$(DecBytes(J))
   Next J
   
   ' Get Binary result
   BinString$ = ""
   For k = MaxBinLen To 1 Step -1
      If BinBits(k) <> 48 Then Exit For
   Next k
   For J = k To 1 Step -1
         BinString$ = BinString$ + Chr$(BinBits(J))
   Next J
End Sub
Private Sub Bytes2Bits()
'IN:  BinBytes(MaxByteLen)
'OUT: BinBits(MaxBinLen)
Dim I As Long, J As Long, k As Long
Dim one As Byte
Dim b As Byte
   b = 48
   FillMemory BinBits(1), MaxBinLen, b    ' "0" to BinBits()
   
   
   ' VB routine
   one = 49 ' "1"
   I = 1
   For J = 1 To MaxByteLen
      b = BinBytes(J)
      For k = 0 To 7
         If (b And 1) <> 0 Then BinBits(I) = one
         b = b \ 2
         I = I + 1
      Next k
   Next J
End Sub

Private Sub Bytes2Dec()
' SLOW IN VB!
'IN:  BinBytes(MaxByteLen)
'OUT: DecBytes(MaxDecLen)
Dim I As Long, k As Long
Dim Carry1 As Byte, Carry2 As Byte
Dim bits As Integer, sum As Integer
Dim b As Byte
   b = 48
   ReDim DecBytes(MaxDecLen + 4)
   FillMemory DecBytes(1), MaxDecLen, b    ' "0" to DecBytes()
   
   ' VB routine
   k = 1
   Do
XX:
      bits = MaxBinLen
      sum = 0
      Do Until bits = 0
         bits = bits - 1
         Carry1 = 0
         ' Shift bits to left with carry
         For I = 1 To MaxByteLen
            ' Check if * 2 will give a carry
            If BinBytes(I) > 127 Then
               Carry2 = 1
               BinBytes(I) = BinBytes(I) - 128
            Else
               Carry2 = 0
            End If
            
            BinBytes(I) = BinBytes(I) * 2 + Carry1  ' Shift << 1 + 1/0
            Carry1 = Carry2
         Next I
         sum = sum * 2 + Carry1    ' Shift << 1 + 1/0
         If sum >= 10 Then
            sum = sum - 10
            BinBytes(1) = BinBytes(1) + 1
         End If
      Loop
      DecBytes(k) = sum + 48  ' Store ASCII digit
      k = k + 1
      ' Check if finished
      For I = MaxByteLen To 1 Step -1
         If BinBytes(I) <> 0 Then GoTo XX ' GoTo used for comparison with ASM
      Next I
      Exit Do
   Loop

End Sub

Private Sub Hex2Bytes(HexString$)
'IN:  HexString$
'OUT: BinBytes(MaxByteLen)
Dim A$
Dim LengthHexStr As Long
Dim k As Long, N As Long
Dim b As Byte
   
   ' Ensure LengthHexStr even
   If (Len(HexString$) And 1) <> 0 Then HexString$ = "0" & HexString$
   LengthHexStr = Len(HexString$)
   b = 48
   FillMemory DecBytes(1), MaxDecLen, b    ' "0" to DecBytes()
   ReDim BinBytes(MaxByteLen)  ' to zero
   ' Transfer 2 nybble values to BinBytes()
   N = 1
   For k = LengthHexStr To 2 Step -2
      A$ = Mid$(HexString$, (k - 1), 2)
      BinBytes(N) = Val("&H" & A$)
      N = N + 1
      If N > MaxByteLen Then Exit For
   Next k
End Sub
Sub SetLengths()
' IN: MaxHexInput
   MaxHexInput = 16
   MaxHexLen = MaxHexInput + 4
   MaxDecLen = MaxHexLen + (MaxHexLen \ 5) + 1
   MaxBinLen = 4 * MaxHexLen
   MaxBinLen = 8 * (MaxBinLen \ 8)  ' Make multiple of 8
   ReDim HexBytes(MaxHexLen)
   ReDim DecBytes(MaxDecLen)
   ReDim BinBits(MaxBinLen)
   MaxByteLen = MaxBinLen \ 8
   ReDim BinBytes(MaxByteLen)
   MaxDecInput = MaxDecLen - 2
   MaxBinInput = MaxBinLen - 8
End Sub
