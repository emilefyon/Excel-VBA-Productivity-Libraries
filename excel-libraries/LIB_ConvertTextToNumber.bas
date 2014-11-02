Attribute VB_Name = "ConvertTextToNumber"

Sub run()
    
    For Each cell In Selection
        If cell.Value <> "" Then Call ConvertTextToNumber(cell)
    Next
    
End Sub

Sub run_WithDecimals()
    
    For Each cell In Selection
        If cell.Value <> "" Then Call ConvertTextToNumber_WithDecimals(cell)
    Next
    
End Sub


Sub ConvertTextToNumber(ByVal cell As Range)
    
    cell.Value = convertTextToNumber_GetCurrentDisplay(cell)
    cell.Value = convertTextToNumber_DashToZero(cell)
    cell.Value = convertTextToNumber_DeleteSeparators(cell.Value)
    cell.Value = convertTextToNumber_Parenthesis(cell.Value)
    cell.Value = convertTextToNumber_AddDecimals(cell.Value)
    cell.NumberFormat = "#,##0.00"
    
End Sub

Sub ConvertTextToNumber_WithDecimals(ByVal cell As Range)
    
    decimals = convertTextToNumber_CountDecimals(cell)
    cell.Value = convertTextToNumber_GetCurrentDisplay(cell)
    cell.Value = convertTextToNumber_DashToZero(cell)
    cell.Value = convertTextToNumber_DeleteSeparators(cell.Value)
    cell.Value = convertTextToNumber_Parenthesis(cell.Value)
    cell.Value = convertTextToNumber_AddDecimals(cell.Value, decimals)
    cell.NumberFormat = "#,##0.00"
    
End Sub

Function convertTextToNumber_CountDecimals(ByVal cell As Range) As Integer
    
    Dim regexMask As String
    Dim sepDecimal As String
    Dim val As String
    
    val = convertTextToNumber_GetCurrentDisplay(cell)
    regexMask = "([0-9\.,]+)"
    sepDecimal = "."
    
    val = findExpreg(val, regexMask)
    If InStr(val, sepDecimal) <> 0 Then
        convertTextToNumber_CountDecimals = Len(val) - InStr(val, sepDecimal)
    Else
        convertTextToNumber_CountDecimals = 0
    End If
    
    

End Function

Function convertTextToNumber_GetCurrentDisplay(ByVal cell As Range) As String
    
    convertTextToNumber_GetCurrentDisplay = WorksheetFunction.Text(cell.Value, cell.NumberFormat)
    
End Function

Function convertTextToNumber_DashToZero(ByVal val As String) As Double
    
    If val = "-" Or val = "--" Then val = 0
    convertTextToNumber_DashToZero = val
    
    
End Function

Function convertTextToNumber_Parenthesis(ByVal val As String) As Double
    
    Dim regexMask As String
    
    sepDecimal = Application.DecimalSeparator
    regexMask = ".*?\(([0-9" + sepDecimal + "]+)\).*"
    If (findExpreg(val, regexMask)) Then val = -CDbl(findExpreg(val, regexMask))

    convertTextToNumber_Parenthesis = val

End Function


Function convertTextToNumber_DeleteSeparators(ByVal val As String) As String

    val = Replace(val, ".", "")
    val = Replace(val, ",", "")
    convertTextToNumber_DeleteSeparators = val

End Function

Function convertTextToNumber_AddDecimals(ByVal val As Double, Optional ByVal numberDecimals As Variant) As Double
        
    If IsMissing(numberDecimals) Then numberDecimals = 0
    
    convertTextToNumber_AddDecimals = val / 10 ^ numberDecimals

End Function


Sub convertDecimalSeparator()
    
    sepDecimal = Application.DecimalSeparator
    Application.DecimalSeparator = "."
    Application.DecimalSeparator = sepDecimal

End Sub


Sub ChangeSystemSeparators()

    Range("A1").Formula = "1,234,567.89"
    MsgBox "The system separators will now change."

    ' Define separators and apply.
    Application.DecimalSeparator = ","
    Application.ThousandsSeparator = "."
    Application.UseSystemSeparators = False

End Sub
