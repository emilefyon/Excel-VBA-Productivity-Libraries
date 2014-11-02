Attribute VB_Name = "LIB_String"
'---------------------------------------------------------------------------------------------------------------------------------------------
'
'   File Library v0.1
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
'
'       Revisions history
'       -----------------
'           - Emile Fyon        11/07/2012      v0.1        Creation
'
'---------------------------------------------------------------------------------------------------------------------------------------------






'---------------------------------------------------------------------------------------------------------------------------------------------
'       + implode (ByVal cellRange As Range, ByVal delimiter As String, Optional ignoreBlank As Boolean) :
'           * Specifications / limitations
'               -
'           * Arguments
'               -
'
'       Last edition date : 24/08/2012
'
'       Revisions history
'       -----------------
'           - Emile Fyon        11/07/2012      Creation
'           - Emile Fyon        24/08/2012      Corrected delimiter bug
'
'---------------------------------------------------------------------------------------------------------------------------------------------


Function implode(ByVal cellRange As Range, ByVal delimiter As String, Optional ignoreBlank As Boolean)
    
    If IsMissing(ignoreBlank) Then ignoreBlank = False
    
    If delimiter = "\n" Then
        delimiter = vbCrLf
    End If
    
    newText = ""
    For Each c In cellRange
        If c.Value = "" And ignoreBlank = True Then
        
        Else
            newText = newText & c.Value & delimiter
        End If
    Next
    newText = Left(newText, Len(newText) - Len(delimiter))
    
    implode = newText
    
    
    
End Function


Function protectForSQL(text As String, Optional ByVal isNumber As Boolean) As String
    
    If IsMissing(isNumber) Then isNumber = False
    
    
    text = Replace(text, "€", "&euro;")
    text = Replace(text, "'", "''")
    
    If isNumber Then text = Replace(text, ",", ".")
    
    protectForSQL = text
    
End Function



Function deleteCharReturns(ByVal txt As String) As String

    txt = Replace(txt, Chr(13), "")
    txt = Replace(txt, Chr(10), "")
    
    deleteCharReturns = txt

End Function

Function ScrubFileName(stringToScrub As String) As String
' remove illegal characters from filenames
Dim newString As String
 
  newString = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(stringToScrub, "|", ""), ">", ""), "<", ""), Chr(34), ""), "?", ""), "*", ""), ":", ""), "/", ""), "\", "")
 
  ScrubFileName = newString
End Function
