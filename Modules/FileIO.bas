Attribute VB_Name = "FileIO"

'Performs file in and out functions for  PS2FontCreator.Ed/Font Creator
Option Explicit

' _______________________________________________________________________________________________________________
'|#######################################[  FONT DATA STRUCTURE     ]############################################|
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Type FONTDATA
    ByteData As String
    spacing As Integer
End Type

Type FontInfo
    Name As String
    fdata(255) As FONTDATA
End Type
Global fontz As FontInfo
Global CurrentLetter
' _______________________________________________________________________________________________________________
'|#######################################[   FONT COMPRESSION        ]###########################################|
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : COMPRESS_FONT / FileIO.bas
' Author    : Xodus
' Date      : 2/4/2013 14:36
' Purpose   : COMPRESSES 324BYTE HEX STRING INTO 20BYTE COMPRESSED DATA
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Function COMPRESS_FONT(data As String)
    '---------------------------------------------------------------------------------------------------------------
    'this is the compression function. It has undergone several revisions, however currently compresses the
    'font set to the rate of about 91%. The font begins # 324 Bytes per char and is reduced to an average of 16-20
    'Bytes per character through repetitive reduction compression. Overall this is a reduction of filesize from
    '82KB to 7.5KB average
    '---------------------------------------------------------------------------------------------------------------
    Dim nf$
    nf$ = data
    
    'COMPLETE "F-reduction COMPRESSION level 1
    If nf$ = "" Then COMPRESS_FONT = nf$: Exit Function 'This will save time if character is blank, no need to continue
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", "Z", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", "Y", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", "X", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", "W", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", "V", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", "U", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", "T", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", "S", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", "R", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", "Q", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", "P", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFFFFFFFFFFFFFF", "O", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFFFFFFFFFFFF", "N", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFFFFFFFFFF", "M", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFFFFFFFF", "L", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFFFFFF", "K", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFFFF", "J", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFF", "I", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFF", "H", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFF", "G", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFF", "E", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFF", "D", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFF", "C", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFF", "B", 1, -1)
    nf$ = Strings.Replace(nf$, "FF", "A", 1, -1)
  
    'F Compression level 2
    nf$ = Strings.Replace(nf$, "ZZZZZZZZZ", "2", 1, -1)
    nf$ = Strings.Replace(nf$, "ZZZZZZZ", "3", 1, -1)
    nf$ = Strings.Replace(nf$, "ZZZZZ", "4", 1, -1)
    nf$ = Strings.Replace(nf$, "ZZZ", "5", 1, -1)
    nf$ = Strings.Replace(nf$, "ZZ", "6", 1, -1)
    
    If Len(nf$) = 1 Then COMPRESS_FONT = nf$: Exit Function 'This will save time if character is blank, no need to continue
    
    'complete 01 Curve Compression - level 1
    nf$ = Strings.Replace(nf$, "010101010101010101010101010101", "o", 1, -1)
    nf$ = Strings.Replace(nf$, "0101010101010101010101010101", "n", 1, -1)
    nf$ = Strings.Replace(nf$, "01010101010101010101010101", "m", 1, -1)
    nf$ = Strings.Replace(nf$, "010101010101010101010101", "l", 1, -1)
    nf$ = Strings.Replace(nf$, "0101010101010101010101", "k", 1, -1)
    nf$ = Strings.Replace(nf$, "01010101010101010101", "j", 1, -1)
    nf$ = Strings.Replace(nf$, "010101010101010101", "i", 1, -1)
    nf$ = Strings.Replace(nf$, "0101010101010101", "h", 1, -1)
    nf$ = Strings.Replace(nf$, "01010101010101", "g", 1, -1)
    nf$ = Strings.Replace(nf$, "010101010101", "f", 1, -1)
    nf$ = Strings.Replace(nf$, "0101010101", "e", 1, -1)
    nf$ = Strings.Replace(nf$, "01010101", "d", 1, -1)
    nf$ = Strings.Replace(nf$, "010101", "c", 1, -1)
    nf$ = Strings.Replace(nf$, "0101", "b", 1, -1)
    nf$ = Strings.Replace(nf$, "01", "a", 1, -1)
    
    'complete "common pairs" compression
    nf$ = Strings.Replace(nf$, "Ab", "p", 1, -1)
    nf$ = Strings.Replace(nf$, "Bb", "q", 1, -1)
    nf$ = Strings.Replace(nf$, "Cb", "r", 1, -1)
    nf$ = Strings.Replace(nf$, "Db", "s", 1, -1)
    nf$ = Strings.Replace(nf$, "Eb", "t", 1, -1)
    nf$ = Strings.Replace(nf$, "Ac", "u", 1, -1)
    nf$ = Strings.Replace(nf$, "Bc", "v", 1, -1)
    nf$ = Strings.Replace(nf$, "Cc", "w", 1, -1)
    nf$ = Strings.Replace(nf$, "Dc", "x", 1, -1)
    nf$ = Strings.Replace(nf$, "Ec", "y", 1, -1)
    nf$ = Strings.Replace(nf$, "Ad", "z", 1, -1)
    nf$ = Strings.Replace(nf$, "Bd", "*", 1, -1)
    nf$ = Strings.Replace(nf$, "Cd", "^", 1, -1)
    nf$ = Strings.Replace(nf$, "Dd", "=", 1, -1)
    nf$ = Strings.Replace(nf$, "Ed", "+", 1, -1)
    nf$ = Strings.Replace(nf$, "Ae", "{", 1, -1)
    nf$ = Strings.Replace(nf$, "Be", "~", 1, -1)
    nf$ = Strings.Replace(nf$, "Ce", "`", 1, -1)
    nf$ = Strings.Replace(nf$, "De", "'", 1, -1)
    nf$ = Strings.Replace(nf$, "Ee", ":", 1, -1)
    nf$ = Strings.Replace(nf$, "Af", "?", 1, -1)
    nf$ = Strings.Replace(nf$, "Bf", ">", 1, -1)
    nf$ = Strings.Replace(nf$, "Cf", "Ç", 1, -1)
    nf$ = Strings.Replace(nf$, "Df", "ü", 1, -1)
    nf$ = Strings.Replace(nf$, "Ef", "é", 1, -1)
    nf$ = Strings.Replace(nf$, "Ra", "â", 1, -1)
    nf$ = Strings.Replace(nf$, "Na", "ä", 1, -1)
    nf$ = Strings.Replace(nf$, "Ba", "à", 1, -1)
    nf$ = Strings.Replace(nf$, "Ka", "å", 1, -1)
    nf$ = Strings.Replace(nf$, "Ja", "}", 1, -1)
    nf$ = Strings.Replace(nf$, "Ga", "|", 1, -1)
    nf$ = Strings.Replace(nf$, "Mb", "/", 1, -1)
    nf$ = Strings.Replace(nf$, "Hb", "\", 1, -1)
    nf$ = Strings.Replace(nf$, "Gb", "<", 1, -1)
    
    COMPRESS_FONT = nf$: Exit Function
    
    '   Ü¢£¥Pƒá

End Function

'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : EXPAND_FONT / FileIO.bas
' Author    : Xodus
' Date      : 2/4/2013 14:35
' Purpose   : DECOMPRESSES FONT BACK TO HEX DATA (is basically a mirror of the compression function)
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Function EXPAND_FONT(data As String)

    Dim nf$
    nf$ = data
           
    
    'complete "common pairs" DE-compression
    nf$ = Strings.Replace(nf$, "p", "Ab", 1, -1)
    nf$ = Strings.Replace(nf$, "q", "Bb", 1, -1)
    nf$ = Strings.Replace(nf$, "r", "Cb", 1, -1)
    nf$ = Strings.Replace(nf$, "s", "Db", 1, -1)
    nf$ = Strings.Replace(nf$, "t", "Eb", 1, -1)
    nf$ = Strings.Replace(nf$, "u", "Ac", 1, -1)
    nf$ = Strings.Replace(nf$, "v", "Bc", 1, -1)
    nf$ = Strings.Replace(nf$, "w", "Cc", 1, -1)
    nf$ = Strings.Replace(nf$, "x", "Dc", 1, -1)
    nf$ = Strings.Replace(nf$, "y", "Ec", 1, -1)
    nf$ = Strings.Replace(nf$, "z", "Ad", 1, -1)
    nf$ = Strings.Replace(nf$, "*", "Bd", 1, -1)
    nf$ = Strings.Replace(nf$, "^", "Cd", 1, -1)
    nf$ = Strings.Replace(nf$, "=", "Dd", 1, -1)
    nf$ = Strings.Replace(nf$, "+", "Ed", 1, -1)
    nf$ = Strings.Replace(nf$, "{", "Ae", 1, -1)
    nf$ = Strings.Replace(nf$, "~", "Be", 1, -1)
    nf$ = Strings.Replace(nf$, "`", "Ce", 1, -1)
    nf$ = Strings.Replace(nf$, "'", "De", 1, -1)
    nf$ = Strings.Replace(nf$, ":", "Ee", 1, -1)
    nf$ = Strings.Replace(nf$, "?", "Af", 1, -1)
    nf$ = Strings.Replace(nf$, ">", "Bf", 1, -1)
    nf$ = Strings.Replace(nf$, "Ç", "Cf", 1, -1)
    nf$ = Strings.Replace(nf$, "ü", "Df", 1, -1)
    nf$ = Strings.Replace(nf$, "é", "Ef", 1, -1)
    nf$ = Strings.Replace(nf$, "â", "Ra", 1, -1)
    nf$ = Strings.Replace(nf$, "ä", "Na", 1, -1)
    nf$ = Strings.Replace(nf$, "à", "Ba", 1, -1)
    nf$ = Strings.Replace(nf$, "å", "Ka", 1, -1)
    nf$ = Strings.Replace(nf$, "}", "Ja", 1, -1)
    nf$ = Strings.Replace(nf$, "|", "Ga", 1, -1)
    nf$ = Strings.Replace(nf$, "/", "Mb", 1, -1)
    nf$ = Strings.Replace(nf$, "\", "Hb", 1, -1)
    nf$ = Strings.Replace(nf$, "<", "Gb", 1, -1)
    
    'complete 01 Curve DECompression - level 1
    nf$ = Strings.Replace(nf$, "o", "010101010101010101010101010101", 1, -1)
    nf$ = Strings.Replace(nf$, "n", "0101010101010101010101010101", 1, -1)
    nf$ = Strings.Replace(nf$, "m", "01010101010101010101010101", 1, -1)
    nf$ = Strings.Replace(nf$, "l", "010101010101010101010101", 1, -1)
    nf$ = Strings.Replace(nf$, "k", "0101010101010101010101", 1, -1)
    nf$ = Strings.Replace(nf$, "j", "01010101010101010101", 1, -1)
    nf$ = Strings.Replace(nf$, "i", "010101010101010101", 1, -1)
    nf$ = Strings.Replace(nf$, "h", "0101010101010101", 1, -1)
    nf$ = Strings.Replace(nf$, "g", "01010101010101", 1, -1)
    nf$ = Strings.Replace(nf$, "f", "010101010101", 1, -1)
    nf$ = Strings.Replace(nf$, "e", "0101010101", 1, -1)
    nf$ = Strings.Replace(nf$, "d", "01010101", 1, -1)
    nf$ = Strings.Replace(nf$, "c", "010101", 1, -1)
    nf$ = Strings.Replace(nf$, "b", "0101", 1, -1)
    nf$ = Strings.Replace(nf$, "a", "01", 1, -1)
    
    'F DECompression level 2
    nf$ = Strings.Replace(nf$, "2", "ZZZZZZZZZ", 1, -1)
    nf$ = Strings.Replace(nf$, "3", "ZZZZZZZ", 1, -1)
    nf$ = Strings.Replace(nf$, "4", "ZZZZZ", 1, -1)
    nf$ = Strings.Replace(nf$, "5", "ZZZ", 1, -1)
    nf$ = Strings.Replace(nf$, "6", "ZZ", 1, -1)
    'F DECompression Level 1
    nf$ = Strings.Replace(nf$, "A", "FF", 1, -1)
    nf$ = Strings.Replace(nf$, "B", "FFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "C", "FFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "D", "FFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "E", "FFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "G", "FFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "H", "FFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "I", "FFFFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "J", "FFFFFFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "K", "FFFFFFFFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "L", "FFFFFFFFFFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "M", "FFFFFFFFFFFFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "N", "FFFFFFFFFFFFFFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "O", "FFFFFFFFFFFFFFFFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "P", "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "Q", "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "R", "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "S", "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "T", "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "U", "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "V", "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "W", "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "X", "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "Y", "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "Z", "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", 1, -1)
        
    EXPAND_FONT = nf$


End Function
' _______________________________________________________________________________________________________________
'|#######################################[ SAVE & LOAD ROUTINES      ]###########################################|
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : OPEN_FONT / FileIO.bas
' Author    : Xodus
' Date      : 2/4/2013 14:35
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Function OPEN_FONT(filename As String)
    Dim fr, ff
    If filename = "" Or ExtOf(filename) <> ".PFC" Then: Exit Function
    
    fr = FreeFile
    Open filename For Input As fr
        Input #1, fontz.Name
        
        For ff = 0 To 255
            Input #1, fontz.fdata(ff).spacing
            Input #1, fontz.fdata(ff).ByteData
            fontz.fdata(ff).ByteData = EXPAND_FONT(fontz.fdata(ff).ByteData)
            'fontz.fdata(ff).ByteData = fontz.fdata(ff).ByteData
        
        Next ff
    Close fr

End Function


'_______________________________________________________________________________________________________________
'---------------------------------------------------------------------------------------------------------------
' Procedure : Save_Font / FileIO.bas
' Author    : Xodus
' Date      : 2/4/2013 14:35
' Purpose   :
'---------------------------------------------------------------------------------------------------------------
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Function Save_Font(filename As String)
    Dim fr, ff, temp
    fr = FreeFile
    If filename = "" Then Exit Function
    Open filename For Output As fr
        Print #1, fontz.Name
        
        For ff = 0 To 255
            Print #1, fontz.fdata(ff).spacing
            temp = COMPRESS_FONT(fontz.fdata(ff).ByteData)
            'temp = fontz.fdata(ff).ByteData
            Print #1, temp
            
        Next ff
    Close fr
End Function

Function NameOf(File As String)
    NameOf = Right(File, Len(File) - InStrRev(File, "/"))
End Function

Function Save_BINFont(filename As String)
    Dim fr, ff, temp
    fr = FreeFile
    If filename = "" Then Exit Function
    Open filename For Output As fr
        Print #1, NameOf(GlobalFileName)
        
        For ff = 0 To 255
            Print #1, fontz.fdata(ff).spacing
            temp = COMPRESS_FONT(fontz.fdata(ff).ByteData)
            Print #1, temp
            
        Next ff
    Close fr
End Function

Function MakeDumFilePFC(filename As String)
    Dim fr, ff, temp
    fr = FreeFile
    If filename = "" Then Exit Function
    Open filename For Output As fr
        Print #1, "temp font"
        
        For ff = 0 To 255
            Print #1, 12
            temp = "25Y"
            Print #1, temp
            
        Next ff
    Close fr
End Function

Sub QuickSaveFont()
If ExtOf(GlobalFileName) = ".PFC" Then
    Save_Font GlobalFileName
Else
    SaveBIN GlobalFileName, GlobalSpaceName
End If
UpdateSaveFile
End Sub

'Function by Dnawrkshpwrkshp, returns the extension of a file as uppercase
Function ExtOf(File As String)
ExtOf = Right(File, Len(File) - InStrRev(File, ".") + 1)
ExtOf = UCase(ExtOf)
End Function

'Converts the PFC to a BIN and saves it to filename, by Dnawrkshpwrkshp
Sub SaveBINfromPFC(filename As String, spacename As String)
Dim temp As String, X As Long

If filename = "" Or spacename = "" Or Dir(GlobalFileName) = "" Then: MsgBox "Error: Missing or invalid files!": Exit Sub
OPEN_FONT (GlobalFileName)

'Save data to BIN
For X = 0 To 255
    temp = temp & HexToString(fontz.fdata(X).ByteData)
Next X

Open filename For Binary As #1
    Put #1, , temp
Close #1

temp = ""
For X = 0 To 255
    temp = temp & HexToString(Pad(Hex$(fontz.fdata(X).spacing), 2))
Next X

Open spacename For Binary As #1
    Put #1, , temp
Close #1

End Sub

'Saves a BIN to filename (BIN), by Dnawrkshpwrkshp
Sub SaveBIN(filename As String, spacename As String)
Dim temp As String, X As Long

If filename = "" Or spacename = "" Then: MsgBox "Error: Missing or invalid files!": Exit Sub

'Save data to BIN
For X = 0 To 255
    temp = temp & HexToString(fontz.fdata(X).ByteData)
Next X

Open filename For Binary As #1
    Put #1, , temp
Close #1

temp = ""
For X = 0 To 255
    temp = temp & HexToString(Hex$(fontz.fdata(X).spacing))
Next X

Open spacename For Binary As #1
    Put #1, , temp
Close #1

End Sub

'Updates entire PS2FontCreator.Ed.grid to file
Sub UpdateSaveFile()

Dim Output As String

    Output = ""
    Dim X
    Dim Offs As Long
    
    For X = 0 To 255
        Output = fontz.fdata(X).ByteData
        Offs = (X * 324) + 1
        SaveCharToFile Output, Offs, 324
    Next X
    
    
End Sub

Function AccessSpaceSegment(Offset As Integer)
    Dim GetChunk As Long, GetByte As Byte, filename As String
    'Offset = Offset + 1
    AccessSpaceSegment = ""
    filename = GlobalSpaceName
    If GlobalSpaceName = "" Then: filename = "temp SPACE.bin"
    If Dir(filename) <> "" Then

        If Offset <= 0 Then: Exit Function
            Open filename For Binary As #4
                For GetChunk = Offset To Offset
                    Get #4, GetChunk, GetByte
                    AccessSpaceSegment = GetByte
                Next GetChunk
            Close
        Else: AccessSpaceSegment = 0
    End If

End Function

Sub SaveSpaceMarker()
Dim Space As Integer, GetChunk As Long, filename As String, Offset As Integer
Space = max
Offset = PS2FontCreator.Ed.CharList1.ListIndex + 1

filename = GlobalSpaceName
If GlobalSpaceName = "" Then: filename = "temp SPACE.bin"
If Dir(filename) <> "" Then
If Offset <= 0 Then: Exit Sub

Open filename For Binary As #1
For GetChunk = Offset To Offset
    Put #1, GetChunk, HexToString(Hex$(Space))
Next GetChunk
Close
Else:
End If
End Sub

'Used by File -> New, makes a temp.bin that is deleted on termination
Sub MakeDumFile(filename As String, Size As Long, Optional Mode As Integer)

Dim GetChunk As Long

If filename = "" Then: Exit Sub
If Size <= 0 Then: Exit Sub

If Mode = 0 Then

Open filename For Binary As #123
For GetChunk = 1 To Size
    Put #123, GetChunk, HexToString("FF")
Next GetChunk
Close

Else
Open filename For Binary As #123
For GetChunk = 1 To Size
    Put #123, GetChunk, HexToString("0C")
Next GetChunk
Close

End If

End Sub

'Saves specific values to a file, called when a PS2FontCreator.Ed.grid box is clicked (BIN only)
Sub SaveCharToFile(data As String, Offset As Long, Size As Long)

Dim GetChunk As Long, filename As String, MidStart As Long
filename = GlobalFileName
If GlobalFileName = "" Then: filename = "temp.bin"

If Dir(filename) <> "" Then

MidStart = 1

If Offset <= 0 Then: Exit Sub
If Size <= 0 Then: Exit Sub

Open filename For Binary As #1
For GetChunk = Offset To (Offset + Size)
    Put #1, GetChunk, HexToString(Mid(data, MidStart, 2))
    MidStart = MidStart + 2
Next GetChunk
Close #1
Else:
End If
End Sub

'Grabs selected bytes from segment in a file, returns hex value as string
Function AccessBinSegment(Offset As Long, Size As Long)
Dim GetChunk As Long, GetByte As Byte, filename As String

AccessBinSegment = AccessBinSegmentFIX(Offset, Size)
Exit Function
AccessBinSegment = ""
filename = GlobalFileName
If GlobalFileName = "" Then: filename = "temp.bin"

    If Dir(filename) <> "" Then
            If Offset <= 0 Then: Exit Function
            If Size <= 0 Then: Exit Function

            Open filename For Binary As #12
                For GetChunk = (Offset) To (Offset + Size)
                    Get #12, GetChunk, GetByte
                    AccessBinSegment = AccessBinSegment & Pad(Hex$(GetByte), 2)
                Next GetChunk
            Close #12
        Else:
    End If
End Function

'Patched after PFC update, characters no longer loaded sideways. By Dnawrkshpwrkshp
'NOT USED JUST HERE FOR POSSIBLE FUTURE USE
Function AccessBinSegmentFIX(Offset As Long, Size As Long)
Dim GetChunk As Long, GetByte As Byte, filename As String, temp As String, temp2(19) As String, X As Integer

AccessBinSegmentFIX = ""
filename = GlobalFileName
If GlobalFileName = "" Then: filename = "temp.bin"

    If Dir(filename) <> "" Then
            If Offset <= 0 Then: Exit Function
            If Size <= 0 Then: Exit Function

            Open filename For Binary As #12
                For GetChunk = (Offset) To (Offset + Size)
                    Get #12, GetChunk, GetByte
                    temp = Pad(Hex$(GetByte), 2)
                    
                    Select Case X
                        Case 0: temp2(X) = temp2(X) & temp
                        Case 1: temp2(X) = temp2(X) & temp
                        Case 2: temp2(X) = temp2(X) & temp
                        Case 3: temp2(X) = temp2(X) & temp
                        Case 4: temp2(X) = temp2(X) & temp
                        Case 5: temp2(X) = temp2(X) & temp
                        Case 6: temp2(X) = temp2(X) & temp
                        Case 7: temp2(X) = temp2(X) & temp
                        Case 8: temp2(X) = temp2(X) & temp
                        Case 9: temp2(X) = temp2(X) & temp
                        Case 10: temp2(X) = temp2(X) & temp
                        Case 11: temp2(X) = temp2(X) & temp
                        Case 12: temp2(X) = temp2(X) & temp
                        Case 13: temp2(X) = temp2(X) & temp
                        Case 14: temp2(X) = temp2(X) & temp
                        Case 15: temp2(X) = temp2(X) & temp
                        Case 16: temp2(X) = temp2(X) & temp
                        Case 17
                            If Len(temp2(17)) = 34 Then
                                For X = 0 To 17
                                    AccessBinSegmentFIX = AccessBinSegmentFIX & temp2(X)
                                    temp2(X) = ""
                                Next X
                            Else
                                temp2(X) = temp2(X) & temp
                            End If
                            X = -1
                    End Select
                    X = X + 1
                    
                Next GetChunk
            Close #12
        Else:
    End If
End Function

Function LoadSett(filename As String, Mode As String)

If filename = "" Then: Exit Function
If Dir(filename) = "" Then: MakeSett filename

Dim GetLine As String

Select Case Mode
    
    Case "inv"
        Open filename For Binary As #2
            Line Input #2, GetLine
        Close
    Case ""
    
End Select

LoadSett = GetLine

End Function

Sub MakeSett(filename As String)

Open filename For Binary As #3
Close

StoreSett filename

End Sub

Sub StoreSett(filename As String)
    Dim StringDat As String
    Dim X
    Open filename For Binary As #3
        For X = 0 To 0
            StringDat = SettArray(X)
            StringDat = SettValue(StringDat) & vbCrLf
            Put #3, , StringDat
        Next X
    Close
End Sub
Function SettValue(Desc As String)
    Select Case Desc
        'Case "InvertColors ": SettValue = Desc & Str(PS2FontCreator.SettingsForm.InvertColors.Value): Exit Function
    End Select
End Function

Sub SaveAsCBIN(filename As String, filespace As String, FileArray As String, SpaceArray As String)
    Dim Res As String, GetString As String, tempFA As String, tempSA As String
    Res = ""

    If Dir(GlobalFileName) = "" Then: Exit Sub
    If Dir(GlobalSpaceName) = "" Then: Exit Sub

    tempFA = FileArray
    tempSA = SpaceArray

    'OPEN EXPORTING IMAGE
    Ed.ExportC.Visible = True
    
    Dim h As Integer
    h = FreeFile
    Open GlobalFileName For Input As #h
        GetString = Input$(LOF(h), h)
    Close #h

    'FileArray = "u32 " & FileArray & "[" & Str((FileLen(GlobalFileName) / 4) + (FileLen(GlobalFileName) Mod 4)) & "] = {" & vbCrLf
    Dim NUMBERSIZE As Long
    NUMBERSIZE = FileLen(GlobalFileName)
    FileArray = "u32 " & FileArray & "[" & Trim(Str(NUMBERSIZE)) & "] = {" & vbCrLf
    Res = FileArray & vbCrLf
    Dim X
    For X = 0 To (FileLen(GlobalFileName) / 4) - 1
        DoEvents
        Res = Res & "   0x" & StringFlip(Pad(StrToHex(Mid(GetString, (X * 4) + 1, 4)), 8)) & "," & vbCrLf
    Next X

    Res = Res & vbCrLf & "};" & vbCrLf

    If Dir(filename) <> "" Then: Kill filename

    Open filename For Binary As #5
        Put #5, , Res
    Close
    Res = ""

    h = FreeFile
    Open GlobalSpaceName For Input As #h
        GetString = Input$(LOF(h), h)
    Close #h

    For X = 0 To (FileLen(GlobalSpaceName) / 4) - 1
        Res = Res & "   0x" & StringFlip(Pad(StrToHex(Mid(GetString, (X * 4) + 1, 4)), 8)) & "," & vbCrLf
    Next X

    NUMBERSIZE = FileLen(GlobalSpaceName)
    SpaceArray = "u32 " & SpaceArray & "[" & Trim(Str(NUMBERSIZE)) & "] = {" & vbCrLf
    Res = SpaceArray & vbCrLf & Res & vbCrLf & "};" & vbCrLf

    If Dir(filespace) <> "" Then: Kill filespace

    Open filespace For Binary As #5
        Put #5, , Res
    Close

    FileArray = tempFA
    SpaceArray = tempSA

    'CLOSE EXPORTING IMAGE
    Ed.ExportC.Visible = False
    'PS2FontCreator.Ed.PicExporting.Visible = False
    'MsgBox "Successfully Exported As " & filename & " and " & filespace

End Sub

Sub SaveAsCPFC(filename As String, filespace As String, FileArray As String, SpaceArray As String)
    Dim Res As String, GetString As String, X As Long, Y As Long, StrCol(18) As String, tempf As String, temps As String
    Res = ""

    If Dir(GlobalFileName) = "" Then: Exit Sub
    
    tempf = FileArray
    temps = SpaceArray
    
    Ed.ExportC.Visible = True
    
    GetString = ""
    OPEN_FONT (GlobalFileName)

    X = 0
    For X = 0 To 255
        GetString = GetString & XtoY(fontz.fdata(X).ByteData)
    Next X

    'FileArray = "u32 " & FileArray & "[" & Str((FileLen(GlobalFileName) / 4) + (FileLen(GlobalFileName) Mod 4)) & "] = {" & vbCrLf
    Dim NUMBERSIZE As Long
    NUMBERSIZE = 82944
    FileArray = "u32 " & FileArray & "[" & Trim(Str(NUMBERSIZE)) & "] = {" & vbCrLf
    Res = FileArray & vbCrLf
    For X = 0 To (Len(GetString) - 1) / 8
        If X Mod 100 = 0 Then: DoEvents
        Res = Res & "   0x" & StringFlip(Pad(Mid(GetString, (X * 8) + 1, 8), 8)) & "," & vbCrLf
    Next X

    Res = Res & vbCrLf & "};" & vbCrLf

    'Save font .C
    If Dir(filename) <> "" Then: Kill filename
    Open filename For Binary As #5
        Put #5, , Res
    Close
    Res = ""
    
    For X = 0 To 255 Step 4
        Res = Res & "   0x" & Pad(Hex$(fontz.fdata(X + 3).spacing), 2) & Pad(Hex$(fontz.fdata(X + 2).spacing), 2) & Pad(Hex$(fontz.fdata(X + 1).spacing), 2) & Pad(Hex$(fontz.fdata(X).spacing), 2) & "," & vbCrLf
    Next X

    NUMBERSIZE = 256
    SpaceArray = "u32 " & SpaceArray & "[" & Trim(Str(NUMBERSIZE)) & "] = {" & vbCrLf
    Res = SpaceArray & vbCrLf & Res & vbCrLf & "};" & vbCrLf

    If Dir(filespace) <> "" Then: Kill filespace

    Open filespace For Binary As #5
        Put #5, , Res
    Close

    FileArray = tempf
    SpaceArray = temps

    'CLOSE EXPORTING IMAGE
    Ed.ExportC.Visible = False
    
    'MsgBox "Successfully Exported As " & filename & " and " & filespace

End Sub

Function XtoY(fdata As String)
Dim X As Long, StrCol(18) As String, EndData As String
X = 0
Do While X < 324
    EndData = Mid(fdata, (2 * X) + 1, 2)
    If EndData = "" Then: EndData = "FF"
    If (X Mod 18) = 0 Then: StrCol(0) = StrCol(0) & EndData: GoTo fontdone
    If (X Mod 18) = 1 Then: StrCol(1) = StrCol(1) & EndData: GoTo fontdone
    If (X Mod 18) = 2 Then: StrCol(2) = StrCol(2) & EndData: GoTo fontdone
    If (X Mod 18) = 3 Then: StrCol(3) = StrCol(3) & EndData: GoTo fontdone
    If (X Mod 18) = 4 Then: StrCol(4) = StrCol(4) & EndData: GoTo fontdone
    If (X Mod 18) = 5 Then: StrCol(5) = StrCol(5) & EndData: GoTo fontdone
    If (X Mod 18) = 6 Then: StrCol(6) = StrCol(6) & EndData: GoTo fontdone
    If (X Mod 18) = 7 Then: StrCol(7) = StrCol(7) & EndData: GoTo fontdone
    If (X Mod 18) = 8 Then: StrCol(8) = StrCol(8) & EndData: GoTo fontdone
    If (X Mod 18) = 9 Then: StrCol(9) = StrCol(9) & EndData: GoTo fontdone
    If (X Mod 18) = 10 Then: StrCol(10) = StrCol(10) & EndData: GoTo fontdone
    If (X Mod 18) = 11 Then: StrCol(11) = StrCol(11) & EndData: GoTo fontdone
    If (X Mod 18) = 12 Then: StrCol(12) = StrCol(12) & EndData: GoTo fontdone
    If (X Mod 18) = 13 Then: StrCol(13) = StrCol(13) & EndData: GoTo fontdone
    If (X Mod 18) = 14 Then: StrCol(14) = StrCol(14) & EndData: GoTo fontdone
    If (X Mod 18) = 15 Then: StrCol(15) = StrCol(15) & EndData: GoTo fontdone
    If (X Mod 18) = 16 Then: StrCol(16) = StrCol(16) & EndData: GoTo fontdone
    If (X Mod 18) = 17 Then: StrCol(17) = StrCol(17) & EndData
fontdone:
            X = X + 1
     Loop

DoEvents

X = 0
EndData = ""
Do While X <= 17
    EndData = EndData & StrCol(X)
    X = X + 1
Loop

XtoY = EndData

End Function

Sub SaveAsO(filename As String, filespace As String, FileArray As String, SpaceArray As String)
Dim FileArraySize As String, SpaceArraySize As String

If Dir(GlobalFileName) = "" Then: Exit Sub
If Dir(GlobalSpaceName) = "" Then: Exit Sub

FileCopy GlobalFileName, filename
FileCopy GlobalSpaceName, filespace

FileArraySize = "size_" & FileArray
SpaceArraySize = "size_" & SpaceArray

Dim FileSize As Long, SpaceSize As Long
FileSize = FileLen(filename)
SpaceSize = FileLen(filespace)

Open filename For Binary As #5
        Put #5, FileSize, HexToString("0000")
        Put #5, FileSize + 2, FileArraySize
        Put #5, FileSize + 2 + Len(FileArraySize), HexToString("00")
        Put #5, FileSize + 3 + Len(FileArraySize), FileArray
        Put #5, FileSize + 3 + Len(FileArraySize) + Len(FileArray), HexToString("00")
Close

Open filespace For Binary As #5
        Put #5, FileSize, HexToString("0000")
        Put #5, FileSize + 2, SpaceArraySize
        Put #5, FileSize + 2 + Len(SpaceArraySize), HexToString("00")
        Put #5, FileSize + 3 + Len(SpaceArraySize), SpaceArray
        Put #5, FileSize + 3 + Len(SpaceArraySize) + Len(SpaceArray), HexToString("00")
Close

End Sub
