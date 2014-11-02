Attribute VB_Name = "LIB_Workbook"
'---------------------------------------------------------------------------------------------------------------------------------------------
'
'   Workbook Library v0.1
'
'
'   Dependencies
'   ------------
'
'       + LIB_Worksheet
'       + LIB_Regex
'
'
'
'
'   Functions lists
'   ---------------
'
'       + Function writeFile (ByVal file As String, ByVal content As String) As String : overwrite the content specified in the file specified.
'           * Specifications / limitations
'               - If the file does not exists, the file is created
'               - The folder has to exist
'           * Arguments
'               - file as String : the full path of the file
'               - content as String : the content that has to be written into the file
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



'---------------------------------------------------------------------------------------------------------------------------------------------
'       + Function moveSheetsInCurrentWorkbook(ByVal wkFullPath As String, Optional ByVal namePattern As String) As String
'           * Description : Return the path of the current workbook
'           * Specifications / limitations
'               - None
'           * Arguments
'               - wkFullPath : the fullPath of the workbook to import the worksheets from
'               - namePattern : a custom name Pattern
'                   #wkName will be replaced is with the name of the destination workbook
'                   #wsName will be the current name of the worksheet
'
'
'       Last edition date : 11/07/2012
'
'       Revisions history
'       -----------------
'           - Emile Fyon        11/07/2012      Creation
'           - Emile Fyon        02/09/2012      Revision in order to make the function ActiveCell-free
'
'---------------------------------------------------------------------------------------------------------------------------------------------

Sub moveSheetsInCurrentWorkbook(ByVal wkFullPath As String, Optional ByVal namePattern As String)
    Dim wkFileName As String
    Dim wsCurrent As Worksheet
    Dim wk As Workbook
    Dim BkName As String
    Dim NumSht As Integer
    Dim BegSht As Integer
    
    wkFileName = fileNameFromFullPath(wkFullPath)

    Set wsCurrent = ActiveSheet
   
    Workbooks.Open Filename:=wkFullPath
    Set wk = Workbooks(wkFileName)
    
    For Each ws In wk.Worksheets
        If IsMissing(namePattern) = False Then
            ws.Name = Replace(ws.Name, "#wsName", ws.Name)
            ws.Name = Replace(ws.Name, "#wkName", wk.Name)
        End If
        ws.Move After:=wsCurrent
    Next
    
    wsCurrent.Select
      'Moves second sheet in source to front of designated workbook.
      'Workbooks(cell.Value).Sheets(BegSht).Move _
      '   Before:=Workbooks("Test.xls").Sheets(1)
         'In each loop, the next sheet in line becomes indexed as number 2.
      'Replace Test.xls with the full name of the target workbook you want.
End Sub






'

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

'---------------------------------------------------------------------------------------------------------------------------------------------
'       + Function listSheets() As String
'           * Description : Write the name of the worksheets of the current workbook in a destination cell
'           * Specifications / limitations
'               - None
'           * Arguments
'               - wkFullPath : the fullPath of the workbook to import the worksheets from
'               - namePattern : a custom name Pattern
'                   #wkName will be replaced is with the name of the destination workbook
'                   #wsName will be the current name of the worksheet
'
'
'       Last edition date : 11/07/2012
'
'       Revisions history
'       -----------------
'           - Emile Fyon        11/07/2012      Creation
'           - Emile Fyon        02/09/2012      Revision in order to make the function ActiveCell-free
'
'---------------------------------------------------------------------------------------------------------------------------------------------


Sub listSheets(Optional ByVal destRg As Range)
    
    Dim rg As Range
    
    If IsMissing(destRg) Then
        Do
            Set rg = Application.InputBox(Prompt:="Where do you want to copy the list of sheets ?", Title:="Choose a range", Type:=8)
        Loop While rg Is Nothing
    End If
    
    i = 0
    For Each ws In ActiveWorkbook.Sheets
        rg.Offset(i, 0).Value = ws.Name
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


