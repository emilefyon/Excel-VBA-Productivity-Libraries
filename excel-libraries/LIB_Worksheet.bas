Attribute VB_Name = "LIB_Worksheet"
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

Function getSheetName(ByVal cell As Range)
    
    getSheetName = cell.Worksheet.Name

End Function

Sub extractPivot()
    
    Dim piv As PivotTable
    Dim wsExtract As Worksheet
    
    defaultShift = 20
    defaultSpace = 5
    
    
    Set wsExtract = Sheets("Pivots")
    
    ' Sheets("Pivots").Cells.Delete
    ' wsExtract.Pictures.Delete
    
    
    Set ws = Sheets("Sheet14")
    Set wsExtract = Sheets("Pivots")
    
    Set piv = ws.PivotTables(1)
    ' Set extractStart = wsExtract.Range("l3")
    Set extractCh = wsExtract.Range("b3")
    
    
    If wsExtract.Range("A1").Value = "" Then
        Set extractStart = wsExtract.Range("l3")
    Else
        Set extractStart = wsExtract.Range(wsExtract.Range("A1").Value)
    End If
    
    Set extractCh = wsExtract.cells(extractStart.Row, 2)
    
    
    ' piv.DataLabelRange
    
    'piv.DataLabelRange.Copy _
            Destination:=extractStart
    
    piv.ColumnRange.Copy _
            Destination:=extractStart.Offset(0, 1)
    
    piv.RowRange.Copy _
            Destination:=extractStart.Offset(piv.ColumnRange.Row - 1, 0)
    
    piv.DataBodyRange.Copy _
            Destination:=extractStart.Offset(piv.ColumnRange.Row, piv.RowRange.Column)
            
    Call extractChart(ws, wsExtract, extractCh)
    
    
    
    If extractStart.Offset(defaultShift, 0) = "" Then
        wsExtract.Range("A1").Formula = extractStart.Offset(defaultShift, 0).Address
    Else
        wsExtract.Range("A1").Formula = extractStart.Offset(defaultShift, 0).End(xlDown).Offset(defaultSpace, 0).Address
    End If
    
End Sub

Sub extractChart(ByVal ws As Worksheet, ByVal wsExtract As Worksheet, ByVal rgExtract As Range)
'
' Macro3 Macro
'

'
    'Dim ws As Worksheet
    'Dim wsExtract As Worksheet
    'Set ws = Sheets("Sheet14")
    'Set wsExtract = Sheets("Pivots")
    Set c = ws.ChartObjects(1).Chart
    
    c.ChartArea.Copy
    
    wsExtract.Activate
    rgExtract.Select
    ActiveSheet.PasteSpecial Format:="Picture (Enhanced Metafile)", Link:=False _
        , DisplayAsIcon:=False
        
    
    'Range("I25").Select
End Sub

