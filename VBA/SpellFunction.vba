Option Explicit

Public SpellArray(0 To 1, 0 To 99) As String
Public SpellSeprators(0 To 1, 0 To 4) As String
Public MiscWords(0 To 1, 0 To 2) As String

Function Spell(v As String, cur As String, Optional lng As Integer = 1) As String

    '@ upto 999 French-Billion
    '@ considered 2 decimal fraction value & rounding off by floor value
    '@ "Un" is used incase of One, Million, Milliard or Billion
    '@ Livre Sterling/Pound has Pense decimal
    '@ lng = 0 for english
    '@ lng = 1 for french---> (Default French conversion in case language not set)

    
    Dim ans(0 To 2) As String, vString() As String, cent As String, dec As String, initial As String
    v = Trim(v)
    vString = Split(v, ".")
    
    Call SpellWords
    Call Seprators
    Call Misc
    
    ans(0) = "": ans(1) = ""
    
    If vString(0) = "1" And lng = 1 Then
        ans(0) = "Un"
    Else
        ans(0) = Trim(Number_to_Word(vString(0), lng))
        initial = LCase(ans(0))
        If InStr(1, initial, " ") > 0 Then: initial = Trim(Left(initial, InStr(1, initial, " ")))
        If lng = 1 And (initial = "million" Or _
           initial = "milliard" Or _
           initial = "billion") _
        Then: ans(0) = "Un " & ans(0)
    End If
    
    ans(0) = ans(0) & " " & cur
    
    If UBound(vString) > 0 Then
        If Len(vString(1)) = 1 Then
            dec = vString(1) & "0"
        Else
            dec = Left(vString(1), 2)
        End If
        If CInt(dec) = 1 Then
            ans(1) = "Un"
        Else
            ans(1) = Number_to_Word(dec, lng)
        End If
    End If
        
    
    If LCase(cur) = "livre sterling" Or LCase(cur) = "pound" Then
        cent = "Pence"
    Else
        cent = MiscWords(lng, 2)
    End If
    
    ans(2) = ans(0) & IIf(ans(1) <> "", " " & MiscWords(lng, 1) & " " & ans(1) & " " & cent, "")
    Spell = Trim(ans(2))
End Function



Private Function Number_to_Word(vString As String, lng As Integer) As String

    Dim i As Integer
    Dim ans As String
    
    Dim vWord As String, fDigit As String, last2Digit As String, fWord As String, last2Word As String
    Dim sp As String
    Dim NofDigits As Integer, fValue As Integer, last2Value As Integer
    
    For i = 0 To 4 ' no of seprators
        NofDigits = Len(vString)
        vWord = Right(vString, 3) ' no of digits to interpret
        fWord = ""
        
        If NofDigits > 2 Then
            fDigit = Left(vWord, 1)
            fValue = CInt(fDigit)
            If fDigit <> "0" Then: fWord = SpellArray(lng, fValue) & " " & MiscWords(lng, 0) & " "
        Else: fValue = 0
        End If
        
        last2Digit = Right(vWord, 2)
        last2Value = CInt(last2Digit)
        last2Word = SpellArray(lng, last2Value)
        
        If fValue + last2Value > 0 Then: ans = fWord & last2Word & " " & SpellSeprators(lng, i) & " " & ans
        
        If NofDigits < 4 Then: Exit For
        vString = Left(vString, NofDigits - 3)
        
    Next
    ans = Replace(ans, "  ", " ")
    Number_to_Word = Trim(StrConv(ans, vbProperCase, 1036)) ' French Fanace ( https://analystcave.com/vba-reference-functions/vba-string-functions/vba-strconv-function/ )

End Function




Sub SpellWords()
    
    '-----english-----
    SpellArray(0, 0) = "" '"zero"
    SpellArray(0, 1) = "One"
    SpellArray(0, 2) = "Two"
    SpellArray(0, 3) = "Three"
    SpellArray(0, 4) = "Four"
    SpellArray(0, 5) = "Five"
    SpellArray(0, 6) = "Six"
    SpellArray(0, 7) = "Seven"
    SpellArray(0, 8) = "Eight"
    SpellArray(0, 9) = "Nine"
    SpellArray(0, 10) = "Ten"
    SpellArray(0, 11) = "Eleven"
    SpellArray(0, 12) = "Twelve"
    SpellArray(0, 13) = "Thirteen"
    SpellArray(0, 14) = "Fourteen"
    SpellArray(0, 15) = "Fifteen"
    SpellArray(0, 16) = "Sixteen"
    SpellArray(0, 17) = "Seventeen"
    SpellArray(0, 18) = "Eighteen"
    SpellArray(0, 19) = "Nineteen"
    SpellArray(0, 20) = "Twenty"
    SpellArray(0, 21) = "Twenty One"
    SpellArray(0, 22) = "Twenty Two"
    SpellArray(0, 23) = "Twenty Three"
    SpellArray(0, 24) = "Twenty Four"
    SpellArray(0, 25) = "Twenty Five"
    SpellArray(0, 26) = "Twenty Six"
    SpellArray(0, 27) = "Twenty Seven"
    SpellArray(0, 28) = "Twenty Eight"
    SpellArray(0, 29) = "Twenty Nine"
    SpellArray(0, 30) = "Thirty"
    SpellArray(0, 31) = "Thirty One"
    SpellArray(0, 32) = "Thirty Two"
    SpellArray(0, 33) = "Thirty Three"
    SpellArray(0, 34) = "Thirty Four"
    SpellArray(0, 35) = "Thirty Five"
    SpellArray(0, 36) = "Thirty Six"
    SpellArray(0, 37) = "Thirty Seven"
    SpellArray(0, 38) = "Thirty Eight"
    SpellArray(0, 39) = "Thirty Nine"
    SpellArray(0, 40) = "Forty"
    SpellArray(0, 41) = "Forty One"
    SpellArray(0, 42) = "Forty Two"
    SpellArray(0, 43) = "Forty Three"
    SpellArray(0, 44) = "Forty Four"
    SpellArray(0, 45) = "Forty Five"
    SpellArray(0, 46) = "Forty Six"
    SpellArray(0, 47) = "Forty Seven"
    SpellArray(0, 48) = "Forty Eight"
    SpellArray(0, 49) = "Forty Nine"
    SpellArray(0, 50) = "Fifty"
    SpellArray(0, 51) = "Fifty One"
    SpellArray(0, 52) = "Fifty Two"
    SpellArray(0, 53) = "Fifty Three"
    SpellArray(0, 54) = "Fifty Four"
    SpellArray(0, 55) = "Fifty Five"
    SpellArray(0, 56) = "Fifty Six"
    SpellArray(0, 57) = "Fifty Seven"
    SpellArray(0, 58) = "Fifty Eight"
    SpellArray(0, 59) = "Fifty Nine"
    SpellArray(0, 60) = "Sixty"
    SpellArray(0, 61) = "Sixty One"
    SpellArray(0, 62) = "Sixty Two"
    SpellArray(0, 63) = "Sixty Three"
    SpellArray(0, 64) = "Sixty Four"
    SpellArray(0, 65) = "Sixty Five"
    SpellArray(0, 66) = "Sixty Six"
    SpellArray(0, 67) = "Sixty Seven"
    SpellArray(0, 68) = "Sixty Eight"
    SpellArray(0, 69) = "Sixty Nine"
    SpellArray(0, 70) = "Seventy"
    SpellArray(0, 71) = "Seventy One"
    SpellArray(0, 72) = "Seventy Two"
    SpellArray(0, 73) = "Seventy Three"
    SpellArray(0, 74) = "Seventy Four"
    SpellArray(0, 75) = "Seventy Five"
    SpellArray(0, 76) = "Seventy Six"
    SpellArray(0, 77) = "Seventy Seven"
    SpellArray(0, 78) = "Seventy Eight"
    SpellArray(0, 79) = "Seventy Nine"
    SpellArray(0, 80) = "Eighty"
    SpellArray(0, 81) = "Eighty One"
    SpellArray(0, 82) = "Eighty Two"
    SpellArray(0, 83) = "Eighty Three"
    SpellArray(0, 84) = "Eighty Four"
    SpellArray(0, 85) = "Eighty Five"
    SpellArray(0, 86) = "Eighty Six"
    SpellArray(0, 87) = "Eighty Seven"
    SpellArray(0, 88) = "Eighty Eight"
    SpellArray(0, 89) = "Eighty Nine"
    SpellArray(0, 90) = "Ninety"
    SpellArray(0, 91) = "Ninety One"
    SpellArray(0, 92) = "Ninety Two"
    SpellArray(0, 93) = "Ninety Three"
    SpellArray(0, 94) = "Ninety Four"
    SpellArray(0, 95) = "Ninety Five"
    SpellArray(0, 96) = "Ninety Six"
    SpellArray(0, 97) = "Ninety Seven"
    SpellArray(0, 98) = "Ninety Eight"
    SpellArray(0, 99) = "Ninety Nine"
    
    '-----french------
    SpellArray(1, 0) = "" '"zéro"
    SpellArray(1, 1) = "" '"un"
    SpellArray(1, 2) = "deux"
    SpellArray(1, 3) = "trois"
    SpellArray(1, 4) = "quatre"
    SpellArray(1, 5) = "cinq"
    SpellArray(1, 6) = "six"
    SpellArray(1, 7) = "sept"
    SpellArray(1, 8) = "huit"
    SpellArray(1, 9) = "neuf"
    SpellArray(1, 10) = "dix"
    SpellArray(1, 11) = "onze"
    SpellArray(1, 12) = "douze"
    SpellArray(1, 13) = "treize"
    SpellArray(1, 14) = "quatorze"
    SpellArray(1, 15) = "quinze"
    SpellArray(1, 16) = "seize"
    SpellArray(1, 17) = "dix sept"
    SpellArray(1, 18) = "dix huit"
    SpellArray(1, 19) = "dix neuf"
    SpellArray(1, 20) = "vingt"
    SpellArray(1, 21) = "vingt et un"
    SpellArray(1, 22) = "vingt deux"
    SpellArray(1, 23) = "vingt trois"
    SpellArray(1, 24) = "vingt quatre"
    SpellArray(1, 25) = "vingt cinq"
    SpellArray(1, 26) = "vingt six"
    SpellArray(1, 27) = "vingt sept"
    SpellArray(1, 28) = "vingt huit"
    SpellArray(1, 29) = "vingt neuf"
    SpellArray(1, 30) = "trente"
    SpellArray(1, 31) = "Trente et un"
    SpellArray(1, 32) = "Trente deux"
    SpellArray(1, 33) = "Trente trois"
    SpellArray(1, 34) = "Trente quatre"
    SpellArray(1, 35) = "Trente cinq"
    SpellArray(1, 36) = "Trente six"
    SpellArray(1, 37) = "Trente sept"
    SpellArray(1, 38) = "Trente huit"
    SpellArray(1, 39) = "Trente neuf"
    SpellArray(1, 40) = "quarante"
    SpellArray(1, 41) = "quarante et un"
    SpellArray(1, 42) = "quarante deux"
    SpellArray(1, 43) = "quarante trois"
    SpellArray(1, 44) = "quarante quatre"
    SpellArray(1, 45) = "quarante cinq"
    SpellArray(1, 46) = "quarante six"
    SpellArray(1, 47) = "quarante sept"
    SpellArray(1, 48) = "quarante huit"
    SpellArray(1, 49) = "quarante neuf"
    SpellArray(1, 50) = "cinquante"
    SpellArray(1, 51) = "cinquante et un"
    SpellArray(1, 52) = "cinquante deux"
    SpellArray(1, 53) = "cinquante trois"
    SpellArray(1, 54) = "cinquante quatre"
    SpellArray(1, 55) = "cinquante cinq"
    SpellArray(1, 56) = "cinquante six"
    SpellArray(1, 57) = "cinquante sept"
    SpellArray(1, 58) = "cinquante huit"
    SpellArray(1, 59) = "cinquante neuf"
    SpellArray(1, 60) = "soixante"
    SpellArray(1, 61) = "soixante et un"
    SpellArray(1, 62) = "soixante deux"
    SpellArray(1, 63) = "soixante trois"
    SpellArray(1, 64) = "soixante quatre"
    SpellArray(1, 65) = "soixante cinq"
    SpellArray(1, 66) = "soixante six"
    SpellArray(1, 67) = "soixante sept"
    SpellArray(1, 68) = "soixante huit"
    SpellArray(1, 69) = "soixante neuf"
    SpellArray(1, 70) = "soixante dix"
    SpellArray(1, 71) = "soixante et onze"
    SpellArray(1, 72) = "soixante douze"
    SpellArray(1, 73) = "soixante treize"
    SpellArray(1, 74) = "soixante quatorze"
    SpellArray(1, 75) = "soixante quinze"
    SpellArray(1, 76) = "soixante seize"
    SpellArray(1, 77) = "soixante dix sept"
    SpellArray(1, 78) = "soixante dix huit"
    SpellArray(1, 79) = "soixante dix neuf"
    SpellArray(1, 80) = "quatre vingts"
    SpellArray(1, 81) = "quatre vingt un"
    SpellArray(1, 82) = "quatre vingt deux"
    SpellArray(1, 83) = "quatre vingt trois"
    SpellArray(1, 84) = "quatre vingt quatre"
    SpellArray(1, 85) = "quatre vingt cinq"
    SpellArray(1, 86) = "quatre vingt six"
    SpellArray(1, 87) = "quatre vingt sept"
    SpellArray(1, 88) = "quatre vingt huit"
    SpellArray(1, 89) = "quatre vingt neuf"
    SpellArray(1, 90) = "quatre vingt dix"
    SpellArray(1, 91) = "quatre vingt onze"
    SpellArray(1, 92) = "quatre vingt douze"
    SpellArray(1, 93) = "quatre vingt treize"
    SpellArray(1, 94) = "quatre vingt quatorze"
    SpellArray(1, 95) = "quatre vingt quinze"
    SpellArray(1, 96) = "quatre vingt seize"
    SpellArray(1, 97) = "quatre vingt dix sept"
    SpellArray(1, 98) = "quatre vingt dix huit"
    SpellArray(1, 99) = "quatre vingt dix neuf"

End Sub

Sub Seprators()
    '-----english-----
    SpellSeprators(0, 0) = "" '"Hundred"
    SpellSeprators(0, 1) = "Thousand"
    SpellSeprators(0, 2) = "Million"
    SpellSeprators(0, 3) = "Billion"
    SpellSeprators(0, 4) = "Trillion"
    
    '-----french------
    SpellSeprators(1, 0) = "" '"Cent"
    SpellSeprators(1, 1) = "Mille"
    SpellSeprators(1, 2) = "Million"
    SpellSeprators(1, 3) = "Milliard"
    SpellSeprators(1, 4) = "Billion"
End Sub

Sub Misc()
    '-----english-----
    MiscWords(0, 0) = "Hundred"
    MiscWords(0, 1) = "and"
    MiscWords(0, 2) = "Cents"
    
    '-----french------
    MiscWords(1, 0) = "Cent"
    MiscWords(1, 1) = "et"
    MiscWords(1, 2) = "Centimes"
    
End Sub
