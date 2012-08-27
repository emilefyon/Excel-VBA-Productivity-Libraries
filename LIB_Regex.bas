Attribute VB_Name = "Regex"
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
'       + Function matchExpreg(ByVal txt As String, ByVal matchPattern As String, ByVal replacePattern As String) As String
'           * Description : Match the specified pattern in the text given in argument and apply the replacementPattern
'           * Specifications / limitations
'               - Multiline
'               - Not case sensitive
'           * Arguments
'               - ByVal txt As String : the text to search in
'               - ByVal matchPattern As String : the regular expression pattern
'               - ByVal replacePattern As String : the replacement pattern
'       + Function findExpreg(ByVal cellContent As Range, ByVal cellPattern As Range) As String
'           * Description : Return the first occurence of the regular expression pattern found in the given expression
'           * Specifications / limitations
'               - Multiline
'               - Not case sensitive
'           * Arguments
'               - ByVal txt As String : the text to search in
'               - ByVal matchPattern As String : the regular expression pattern
'       + Function stripTags(ByVal txt As String) As String
'           * Description : Strips all the tags within a given string
'           * Specifications / limitations
'               - None
'           * Arguments
'               - ByVal txt As String : the text to search in'
'
'
'       Revisions history
'       -----------------
'           - Emile Fyon        11/07/2012      v0.1        Creation
'
'---------------------------------------------------------------------------------------------------------------------------------------------




'---------------------------------------------------------------------------------------------------------------------------------------------
'       + Function matchExpreg(ByVal txt As String, ByVal matchPattern As String, ByVal replacePattern As String) As String
'           * Description : Match the specified pattern in the text given in argument and apply the replacementPattern
'           * Specifications / limitations
'               - Multiline
'               - Not case sensitive
'           * Arguments
'               - ByVal txt As String : the text to search in
'               - ByVal matchPattern As String : the regular expression pattern
'               - ByVal replacePattern As String : the replacement pattern
'
'
'       Last edition date : 11/07/2012
'
'       Revisions history
'       -----------------
'           - Emile Fyon        11/07/2012      Creation
'
'---------------------------------------------------------------------------------------------------------------------------------------------


Function matchExpreg(ByVal txt As String, ByVal matchPattern As String, ByVal replacePattern As String) As String
    Dim RE As Object, REMatches As Object
    
    ' Set cell = Range("e15")
    ' strData = cell.Value
     
    Dim reg_exp As New RegExp
    reg_exp.pattern = matchPattern
    reg_exp.IgnoreCase = True
    reg_exp.Global = True
    
    txt = reg_exp.Replace(txt, replacePattern)
    matchExpreg = txt

     
End Function


'---------------------------------------------------------------------------------------------------------------------------------------------
'       + Function findExpreg(ByVal cellContent As Range, ByVal cellPattern As Range) As String
'           * Description : Return the first occurence of the regular expression pattern found in the given expression
'           * Specifications / limitations
'               - Multiline
'               - Not case sensitive
'           * Arguments
'               - ByVal txt As String : the text to search in
'               - ByVal matchPattern As String : the regular expression pattern
'
'
'       Last edition date : 11/07/2012
'
'       Revisions history
'       -----------------
'           - Emile Fyon        11/07/2012      Creation
'           - Emile Fyon        27/08/2012      Handle response in case the string is not found (return False)
'
'---------------------------------------------------------------------------------------------------------------------------------------------


Function findExpreg(ByVal txt As String, ByVal matchPattern As String) As String

    On Error GoTo errorHandler
     
    Dim expReg As New RegExp
    expReg.pattern = matchPattern
    expReg.IgnoreCase = True
    expReg.Global = True
    
    Set res = expReg.Execute(txt)
    
    txt = res(0).submatches(0)
    findExpreg = txt
    Exit Function
errorHandler:
    findExpreg = False
     
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------
'       + Function stripTags(ByVal txt As String) As String
'           * Description : Strips all the tags within a given string
'           * Specifications / limitations
'               - None
'           * Arguments
'               - ByVal txt As String : the text to search in
'
'
'       Last edition date : 11/07/2012
'
'       Revisions history
'       -----------------
'           - Emile Fyon        11/07/2012      Creation
'
'---------------------------------------------------------------------------------------------------------------------------------------------


Function stripTags(ByVal txt As String) As String

    regMask = "(<.+?>)"
    stripTags = matchExpreg(txt, regMask, "")
    

End Function
