Attribute VB_Name = "LIB_Range"
' Documentation on this library can be found on https://github.com/emilefyon/Excel-VBA-Productivity-Libraries/blob/master/docs/Lib_Range.md

Function getAddress(ByVal cell As Range)
    
    getAddress = cell.Address

End Function


Function selectXlDownRange(ByVal cell As Range) As Range
    Application.Volatile ' Force Excel to recalculate on workbook change
    If cell.Offset(1, 0) = "" Then
        Set selectXlDownRange = cell
        Exit Function
    End If
    Set selectXlDownRange = Range(cell, cell.End(xlDown))
    
End Function


