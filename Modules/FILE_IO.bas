Attribute VB_Name = "FILE_IO"
Option Explicit
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

Function COMPRESS_FONT(data As String)
    
    Dim nf$
    nf$ = data
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFFFFFFFF", "$", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFFFFFF", "#", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFFFF", "@", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFFFF", "A", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFFFF", "B", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFFFF", "9", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFFFF", "C", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFFFF", "8", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFFFF", "D", 1, -1)
    nf$ = Strings.Replace(nf$, "FFFF", "7", 1, -1)
    nf$ = Strings.Replace(nf$, "FF", "E", 1, -1)
    nf$ = Strings.Replace(nf$, "0101010101010101", "6", 1, -1)
    nf$ = Strings.Replace(nf$, "01010101010101", "5", 1, -1)
    nf$ = Strings.Replace(nf$, "010101010101", "4", 1, -1)
    nf$ = Strings.Replace(nf$, "0101010101", "3", 1, -1)
    nf$ = Strings.Replace(nf$, "01010101", "G", 1, -1)
    nf$ = Strings.Replace(nf$, "010101", "H", 1, -1)
    nf$ = Strings.Replace(nf$, "0101", "I", 1, -1)
    nf$ = Strings.Replace(nf$, "01", "J", 1, -1)
    nf$ = Strings.Replace(nf$, "AAAA", "K", 1, -1)
    nf$ = Strings.Replace(nf$, "BBBB", "L", 1, -1)
    nf$ = Strings.Replace(nf$, "CCCC", "M", 1, -1)
    nf$ = Strings.Replace(nf$, "DDDD", "N", 1, -1)
    nf$ = Strings.Replace(nf$, "AA", "O", 1, -1)
    nf$ = Strings.Replace(nf$, "BB", "P", 1, -1)
    nf$ = Strings.Replace(nf$, "CC", "Q", 1, -1)
    nf$ = Strings.Replace(nf$, "DD", "R", 1, -1)
    nf$ = Strings.Replace(nf$, "AD", "S", 1, -1)
    nf$ = Strings.Replace(nf$, "AB", "T", 1, -1)
    nf$ = Strings.Replace(nf$, "AC", "U", 1, -1)
    nf$ = Strings.Replace(nf$, "AE", "V", 1, -1)
    nf$ = Strings.Replace(nf$, "GH", "W", 1, -1)
    nf$ = Strings.Replace(nf$, "GI", "X", 1, -1)
    nf$ = Strings.Replace(nf$, "GJ", "Y", 1, -1)
    nf$ = Strings.Replace(nf$, "JSIV", "Z", 1, -1)
    nf$ = Strings.Replace(nf$, "JAIB", "/", 1, -1)
    nf$ = Strings.Replace(nf$, "ZZZZ", "2", 1, -1)
    
    COMPRESS_FONT = nf$


End Function

Function EXPAND_FONT(data As String)

    Dim nf$
    nf$ = data
    nf$ = Strings.Replace(nf$, "2", "ZZZZ", 1, -1)
    nf$ = Strings.Replace(nf$, "/", "JAIB", 1, -1)
    nf$ = Strings.Replace(nf$, "Z", "JSIV", 1, -1)
    nf$ = Strings.Replace(nf$, "Y", "GJ", 1, -1)
    nf$ = Strings.Replace(nf$, "X", "GI", 1, -1)
    nf$ = Strings.Replace(nf$, "W", "GH", 1, -1)
    nf$ = Strings.Replace(nf$, "V", "AE", 1, -1)
    nf$ = Strings.Replace(nf$, "U", "AC", 1, -1)
    nf$ = Strings.Replace(nf$, "T", "AB", 1, -1)
    nf$ = Strings.Replace(nf$, "S", "AD", 1, -1)
    nf$ = Strings.Replace(nf$, "R", "DD", 1, -1)
    nf$ = Strings.Replace(nf$, "Q", "CC", 1, -1)
    nf$ = Strings.Replace(nf$, "P", "BB", 1, -1)
    nf$ = Strings.Replace(nf$, "O", "AA", 1, -1)
    nf$ = Strings.Replace(nf$, "N", "DDDD", 1, -1)
    nf$ = Strings.Replace(nf$, "M", "CCCC", 1, -1)
    
    nf$ = Strings.Replace(nf$, "L", "BBBB", 1, -1)
    nf$ = Strings.Replace(nf$, "K", "AAAA", 1, -1)
    nf$ = Strings.Replace(nf$, "J", "01", 1, -1)
    nf$ = Strings.Replace(nf$, "I", "0101", 1, -1)
    nf$ = Strings.Replace(nf$, "H", "010101", 1, -1)
    nf$ = Strings.Replace(nf$, "G", "01010101", 1, -1)
    nf$ = Strings.Replace(nf$, "3", "0101010101", 1, -1)
    nf$ = Strings.Replace(nf$, "4", "010101010101", 1, -1)
    nf$ = Strings.Replace(nf$, "5", "01010101010101", 1, -1)
    nf$ = Strings.Replace(nf$, "6", "0101010101010101", 1, -1)
    nf$ = Strings.Replace(nf$, "E", "FF", 1, -1)
    nf$ = Strings.Replace(nf$, "7", "FFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "D", "FFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "8", "FFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "C", "FFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "9", "FFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "B", "FFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "A", "FFFFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "@", "FFFFFFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "#", "FFFFFFFFFFFFFFFFFFFF", 1, -1)
    nf$ = Strings.Replace(nf$, "$", "FFFFFFFFFFFFFFFFFFFFFF", 1, -1)
    
    EXPAND_FONT = nf$


End Function


Function OPEN_FONT(filename As String)
    Dim fr, ff
    fr = FreeFile
    Open filename For Input As fr
        Input #1, fontz.Name
        
        For ff = 0 To 255
            Input #1, fontz.fdata(ff).spacing
            Input #1, fontz.fdata(ff).ByteData
            fontz.fdata(ff).ByteData = EXPAND_FONT(fontz.fdata(ff).ByteData)
        
        Next ff
    Close fr

End Function


Function Save_Font(filename As String)
    Dim fr, ff
    fr = FreeFile
    If filename = "" Then Exit Function
    Open filename For Output As fr
        Print #1, fontz.Name
        
        For ff = 0 To 255
            Print #1, fontz.fdata(ff).spacing
            fontz.fdata(ff).ByteData = COMPRESS_FONT(fontz.fdata(ff).ByteData)
            Print #1, fontz.fdata(ff).ByteData
            
        Next ff
    Close fr
End Function
