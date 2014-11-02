Attribute VB_Name = "LIB_Substitutes"



Function SuperSubV(strOldText As String, ByVal Rng1 As Range, ByVal rng2 As Range) As String
    Dim cel As Range
    Dim strMyChar As String, strMyReplace As String
    For Each cel In Rng1
        strMyChar = cel.Value
        strMyReplace = cel.Offset(0, rng2.Column - Rng1.Column).Value
         ' Next line does not work.
         'It would require a VBA version of Excel's SUBSTITUTE to work
        strOldText = Replace(strOldText, strMyChar, strMyReplace)

    Next cel
    SuperSubV = strOldText
End Function

Function SuperSubH(strOldText As String, ByVal Rng1 As Range, ByVal rng2 As Range) As String
    Dim cel As Range
    Dim strMyChar As String, strMyReplace As String
    For Each cel In Rng1
        strMyChar = cel.Value
        strMyReplace = cel.Offset(rng2.Row - Rng1.Row, 0).Value
         ' Next line does not work.
         'It would require a VBA version of Excel's SUBSTITUTE to work
        strOldText = Replace(strOldText, strMyChar, strMyReplace)

    Next cel
    SuperSubH = strOldText
End Function


