Attribute VB_Name = "LIB_Workbook"

'---------------------------------------------------------------------------------------------------------------------------------------------
'
'   Regex Library v0.1
'
'   + References dependencies
'       - Microsoft VBScript Regular Expressions 1.0
'
'
'   Functions lists
'   ---------------
'
'
'
'       Revisions history
'       -----------------
'           - Emile Fyon        11/07/2012      v0.1        Creation
'
'---------------------------------------------------------------------------------------------------------------------------------------------






'---------------------------------------------------------------------------------------------------------------------------------------------
'       + Function getCurrentWorkbookPath() As String
'           * Description : Return the path of the current workbook
'           * Specifications / limitations
'               - None
'           * Arguments
'               - None
'
'
'       Last edition date : 11/07/2012
'
'       Revisions history
'       -----------------
'           - Emile Fyon        11/07/2012      Creation
'
'---------------------------------------------------------------------------------------------------------------------------------------------

Function getCurrentWorkbookPath()
    
    getCurrentWorkbookPath = checkFolder(ActiveWorkbook.Path)

End Function





Sub moveSheetsInCurrentWorkbook()
   Dim BkName As String
   Dim NumSht As Integer
   Dim BegSht As Integer
   
   Set wsCurrent = ActiveSheet

    For Each cell In Selection
        Workbooks.Open filename:=cell.Offset(0, 1).Value
        Set wk = Workbooks(cell.Value)
        For Each ws In wk.Worksheets
            If cell.Offset(0, -1).Value <> "" Then ws.Name = getSheetName(cell.Offset(0, -1).text, ws, wk)
            ws.Move After:=wsCurrent
        Next
      'Moves second sheet in source to front of designated workbook.
      'Workbooks(cell.Value).Sheets(BegSht).Move _
      '   Before:=Workbooks("Test.xls").Sheets(1)
         'In each loop, the next sheet in line becomes indexed as number 2.
      'Replace Test.xls with the full name of the target workbook you want.
    Next
End Sub


Function getSheetNameRedo(ByVal pattern As String, ByVal ws As Worksheet, ByVal wk As Workbook) As String

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
                sheetName = Replace(sheetName, s, ws.Range(cellAddress).text)
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




Sub getReplacementPatterns()
    
    ActiveCell.Offset(0, 0) = "$A1$"
    ActiveCell.Offset(0, 1) = "Value of cell A1 in worksheet"
    
    ActiveCell.Offset(1, 0) = "#wsName"
    ActiveCell.Offset(1, 1) = "Name of the worksheet"
    
    ActiveCell.Offset(2, 0) = "#wkName"
    ActiveCell.Offset(2, 1) = "Name of the workbook"
    
    ActiveCell.Offset(3, 0) = "The worksheet name will be automatically trimed to the first 31 characters"
    ActiveCell.Offset(4, 0) = "If you don't use any pattern, the value will be used as a prefix for the new sheet name"
    
    

End Sub


Sub listSheets()
    
    i = 0
    For Each ws In ActiveWorkbook.Sheets
        ActiveCell.Offset(i, 0).Value = ws.Name
        i = i + 1
    Next

End Sub

Sub copySheets()
    
    Set ws = ActiveSheet
    For Each cell In Selection
        ws.Copy After:=ws
        Sheets(ws.Index + 1).Name = cell.Value
    Next

End Sub



Sub renameSheets()
    
    For Each cell In Selection
        Sheets(cell.Value).Name = cell.Offset(0, 1).Value
    Next

End Sub

Sub concatenateSheets()
    
    Set cell = ActiveCell
    
    Set wsExtract = Sheets(cell.Value)
    wsExtract.cells.ClearContents
    
    For Each cell In Range(cell.Offset(1, 0), cell.End(xlDown))
        If wsExtract.Range("A1").Value = "" Then
            Set extractStart = wsExtract.Range("A1")
        Else
            Set extractStart = wsExtract.Range("A1").End(xlDown).Offset(1, 0)
        End If
        Set ws = Sheets(cell.Value)
        Range(ws.Range("A1"), ws.Range("A1").End(xlToRight).End(xlDown)).Copy
        wsExtract.Activate
        If wsExtract.Range("A1") = "" Then
            Set pasteCell = wsExtract.Range("A1")
        Else
            Set pasteCell = wsExtract.Range("A1").End(xlDown).Offset(1, 0)
        End If
        pasteCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Next

End Sub


