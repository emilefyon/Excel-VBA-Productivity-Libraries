Attribute VB_Name = "FoldersFiles"
'---------------------------------------------------------------------------------------------------------------------------------------------
'
'   FoldersFiles Library v0.1
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
'       + Function readFile(ByVal file As String) As String : read the content of a file and return a single line with all the content
'           * Specifications / limitations
'               - The file has to Exists, no error handling
'               - The content is retrieved without any line returns (line returns are replaced by space)
'           * Arguments
'               - file as String : the full path of the file
'       + Function readFileAndTruncate(ByVal file As String) As String : calls readFile() and then truncate the text to 30000 characters in order to avoid Excel limitations
'           * Specifications / limitations
'               - The file has to Exists, no error handling
'               - The content is retrieved without any line returns (line returns are replaced by space)
'               - Only the 30.000 first characters are retrieved
'           * Arguments
'               - file as String : the full path of the file
'
'       Revisions history
'       -----------------
'           - Emile Fyon        11/07/2012      v0.1        Creation
'
'---------------------------------------------------------------------------------------------------------------------------------------------





'---------------------------------------------------------------------------------------------------------------------------------------------
'       + ListFiles(ByVal strPath As String, ByVal cellDestination As Range)
'           * Description : List the files in the folder specified in argument and display the list of the files in the cells
'                           below the cell given in argument as well as the full path in the right column
'
'           * Specifications / limitations
'               - Has to be launched by an other macro
'           * Arguments
'               - strPath as String : the full path of the folder
'               - cellDestination as Range : the destination cell
'
'       Last edition date : 11/07/2012
'
'       Revisions history
'       -----------------
'           - Emile Fyon        11/07/2012      Creation
'
'---------------------------------------------------------------------------------------------------------------------------------------------



Sub ListFiles(ByVal strPath As String, ByVal cellDestination As Range)

    ' Local variables
    Dim counter As Integer
    Dim file As String
    ' Dim filesTab
    
    ' Add a trailing slash if needed
    strPath = checkFolder(strPath)
    
    file = Dir$(strPath & Extention)
    
    ' Count the number of files in the folder
    Do While Len(file)
        file = Dir$
        counter = counter + 1
    Loop
    ReDim filesTab(counter - 1)
    counter = 0

    ' Reset the counter of the array
    file = Dir$(strPath & Extention)

    ' List the files and display them in the cells
    Do While Len(file) And counter <= UBound(filesTab)
        cellDestination.Offset(counter, 0) = file
        cellDestination.Offset(counter, 1) = strPath & file
        file = Dir$
        counter = counter + 1
    Loop
    


End Sub


Function checkFolder(ByVal strPath As String) As String
    
    ' Add a trailing slash if needed
    If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"
    checkFolder = strPath

End Function


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
