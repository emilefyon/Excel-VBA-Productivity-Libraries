

'---------------------------------------------------------------------------------------------------------------------------------------------
'
'   Range Actions Library v0.1
'
'
'   Functions lists
'   ---------------
'
'       + Sub addLinks () : add an hyperlink to all the cells in the current selection. The URL will be the content of the cell
'           * Specifications / limitations
'               - None
'           * Arguments
'               - None
'       + Sub addLink (ByVal url As String, ByVal cell As Range) : add an hyperlink to cell given in argument to the URL given in argument. The URL will be the content of the cell
'           * Specifications / limitations
'               - None
'           * Arguments
'               - ByVal url As String the URL the cell must point to
'               - ByVal cell As Range the cell to add the link
'
'       Revisions history
'       -----------------
'           - Emile Fyon        11/07/2012      v0.1        Creation
'
'---------------------------------------------------------------------------------------------------------------------------------------------




'---------------------------------------------------------------------------------------------------------------------------------------------
'       + Sub addLinks () : add an hyperlink to all the cells in the current selection. The URL will be the content of the cell
'           * Specifications / limitations
'               - None
'           * Arguments
'               - None
'
'       Last edition date : 11/07/2012
'
'       Revisions history
'       -----------------
'           - Emile Fyon        11/07/2012      Creation
'
'---------------------------------------------------------------------------------------------------------------------------------------------

Sub addLinks()
    
    For Each cell In Selection
        Call addLink(cell.Value, cell)
    Next

End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'       + Sub addLink (ByVal url As String, ByVal cell As Range) : add an hyperlink to cell given in argument to the URL given in argument. The URL will be the content of the cell
'           * Specifications / limitations
'               - None
'           * Arguments
'               - ByVal url As String the URL the cell must point to
'               - ByVal cell As Range the cell to add the link
'
'       Last edition date : 11/07/2012
'
'       Revisions history
'       -----------------
'           - Emile Fyon        11/07/2012      Creation
'
'---------------------------------------------------------------------------------------------------------------------------------------------

Sub addLink(ByVal url As String, ByVal cell As Range)

    cell.Worksheet.Hyperlinks.Add Anchor:=cell, Address:=url


End Sub



'---------------------------------------------------------------------------------------------------------------------------------------------
'       + Sub addLink (ByVal url As String, ByVal cell As Range) : add an hyperlink to cell given in argument to the URL given in argument. The URL will be the content of the cell
'           * Specifications / limitations
'               - None
'           * Arguments
'               - ByVal url As String the URL the cell must point to
'               - ByVal cell As Range the cell to add the link
'
'       Last edition date : 11/07/2012
'
'       Revisions history
'       -----------------
'           - Emile Fyon        11/07/2012      Creation
'
'---------------------------------------------------------------------------------------------------------------------------------------------

Function getAddress(ByVal cell As Range)
    
    getAddress = cell.Address

End Function



'------------------------------------------------------


Function selectXlDownRange(ByVal cell As Range) As Range
    Application.Volatile ' Force Excel to recalculate on workbook change
    If cell.Offset(1, 0) = "" Then
        Set selectXlDownRange = cell
        Exit Function
    End If
    Set selectXlDownRange = Range(cell, cell.End(xlDown))
    
End Function


