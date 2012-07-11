Attribute VB_Name = "Files"
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
'       + writeFile (ByVal file As String, ByVal content As String) : overwrite the content specified in the file specified.
'           * Specifications / limitations
'               - If the file does not exists, the file is created
'               - The folder has to exist
'           * Arguments
'               - file as String : the full path of the file
'               - content as String : the content that has to be written into the file
'
'       Last edition date : 11/07/2012
'
'       Revisions history
'       -----------------
'           - Emile Fyon        11/07/2012      Creation
'
'---------------------------------------------------------------------------------------------------------------------------------------------

Function writeFile(ByVal file As String, ByVal content As String) As String
    
   Open file For Output As #1
   Print #1, content
   Close #1
   
   writeFile = "File updated yet"

End Function

'---------------------------------------------------------------------------------------------------------------------------------------------
'       + readFile(ByVal file As String) As String : read the content of a file and return a single line with all the content
'           * Specifications / limitations
'               - The file has to Exists, no error handling
'               - The content is retrieved without any line returns (line returns are replaced by space)
'           * Arguments
'               - file as String : the full path of the file
'
'
'       Last edition date : 11/07/2012
'
'       Revisions history
'       -----------------
'           - Emile Fyon        11/07/2012      Creation
'
'---------------------------------------------------------------------------------------------------------------------------------------------

Function readFile(ByVal file As String) As String

    Dim MyString, MyNumber
    Open file For Input As #1 ' Open file for input.
    fileContent = ""
    Do While Not EOF(1) ' Loop until end of file.
        Line Input #1, MyString
        Debug.Print MyString
        fileContent = fileContent & MyString & " "
    Loop
    Close #1 ' Close file.
    readFile = fileContent
End Function


'---------------------------------------------------------------------------------------------------------------------------------------------
'       + readFileAndTruncate(ByVal file As String) As String : calls readFile() and then truncate the text to 30000 characters in order to avoid Excel limitations
'           * Specifications / limitations
'               - The file has to Exists, no error handling
'               - The content is retrieved without any line returns (line returns are replaced by space)
'               - Only the 30.000 first characters are retrieved
'           * Arguments
'               - file as String : the full path of the file
'
'
'       Last edition date : 11/07/2012
'
'       Revisions history
'       -----------------
'           - Emile Fyon        11/07/2012      Creation
'
'---------------------------------------------------------------------------------------------------------------------------------------------


Function readFileAndTruncate(ByVal file As String) As String

    readFileAndTruncate = Left(readFile(file), 30000)

End Function
