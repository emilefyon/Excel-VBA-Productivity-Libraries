Attribute VB_Name = "FoldersFiles"
Private Sub ListFiles(control As IRibbonControl)

strPath = ActiveCell.Value

Dim counter As Integer
'Leave Extention blank for all files
   Dim file As String
   If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"
    If Trim$(Extention) = "" Then
        Extention = "*.*"
    ElseIf Left$(Extention, 2) <> "*." Then
        Extention = "*." & Extention
    End If
    file = Dir$(strPath & Extention)


    Do While Len(file)
        file = Dir$
        counter = counter + 1
    Loop
    ReDim listOfQuery(counter - 1)
    counter = 0


    If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"
    If Trim$(Extention) = "" Then
        Extention = "*.*"
    ElseIf Left$(Extention, 2) <> "*." Then
        Extention = "*." & Extention
    End If

    file = Dir$(strPath & Extention)

    ' listOfQuery(counter) = file
    ' counter = counter + 1
    Do While Len(file) And counter <= UBound(listOfQuery)
        
        ' listOfQuery(counter) = file
        ActiveCell.Offset(counter + 1, 0) = file
        ActiveCell.Offset(counter + 1, 1) = strPath & file
        file = Dir$
        counter = counter + 1
    Loop


End Sub


Sub moveSheetsInCurrentWorkbook(control As IRibbonControl)
   Dim BkName As String
   Dim NumSht As Integer
   Dim BegSht As Integer
   
   Set wsCurrent = ActiveSheet

    For Each cell In Selection
        Workbooks.Open Filename:=cell.Offset(0, 1).Value
        Set wk = Workbooks(cell.Value)
        For Each ws In wk.Worksheets
            If cell.Offset(0, -1).Value <> "" Then ws.Name = getSheetName(cell.Offset(0, -1).Text, ws, wk)
            ws.Move after:=wsCurrent
        Next
      'Moves second sheet in source to front of designated workbook.
      'Workbooks(cell.Value).Sheets(BegSht).Move _
      '   Before:=Workbooks("Test.xls").Sheets(1)
         'In each loop, the next sheet in line becomes indexed as number 2.
      'Replace Test.xls with the full name of the target workbook you want.
    Next
End Sub

Function getSheetName(ByVal pattern As String, ByVal ws As Worksheet, ByVal wk As Workbook)

    'r = ActiveCell.Value
    'Set ws = ActiveSheet
    
    sheetName = pattern
    With CreateObject("vbscript.regexp")
        .pattern = "\$(.+?)\$"
        .Global = True
        If .test(pattern) Then
            For Each s In .Execute(pattern)
                ' MsgBox (s)
                cellAddress = Replace(s, "$", "")
                sheetName = Replace(sheetName, s, ws.Range(cellAddress).Text)
                ' r = Replace(r, s, Replace(s, ",", "#"))
            Next 'extractBrackets = .Execute(r)(0)
        End If
    End With
    sheetName = Replace(sheetName, "#wsName", ws.Name)
    sheetName = Replace(sheetName, "#wkName", wk.Name)
    If sheetName = pattern Then sheetName = pattern & " " & ws.Name
    'MsgBox (r)
    
    
    
    getSheetName = Left(sheetName, 31)

End Function

Sub getReplacementPatterns(control As IRibbonControl)
    
    ActiveCell.Offset(0, 0) = "$A1$"
    ActiveCell.Offset(0, 1) = "Value of cell A1 in worksheet"
    
    ActiveCell.Offset(1, 0) = "#wsName"
    ActiveCell.Offset(1, 1) = "Name of the worksheet"
    
    ActiveCell.Offset(2, 0) = "#wkName"
    ActiveCell.Offset(2, 1) = "Name of the workbook"
    
    ActiveCell.Offset(3, 0) = "The worksheet name will be automatically trimed to the first 31 characters"
    ActiveCell.Offset(4, 0) = "If you don't use any pattern, the value will be used as a prefix for the new sheet name"
    
    

End Sub
