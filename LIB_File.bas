Attribute VB_Name = "LIB_File"
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

Function writeFile(ByVal File As String, ByVal content As String) As String
    
   Open File For Output As #1
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
'               - Optional createFile as Boolean : indicates if we have to create the file if it does not exists (default False)
'
'
'       Last edition date : 11/07/2012
'
'       Revisions history
'       -----------------
'           - Emile Fyon        11/07/2012      Creation
'           - Emile Fyon        12/07/2012      Added functionality in order to create the file if it does not exists
'
'---------------------------------------------------------------------------------------------------------------------------------------------

Function readFile(ByVal File As String, Optional createFile As Boolean) As String
    
    If (IsMissing(createFile)) Then createFile = False
    
    If (fileExists(File) = False) Then
        If (createFile = True) Then
            temp = writeFile(File, "")
        Else
            readFile = "Error : File does not exists"
            Exit Function
        End If
    End If
    
    Dim MyString, MyNumber
    Open File For Input As #1 ' Open file for input.
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


Function readFileAndTruncate(ByVal File As String, Optional createFile As Boolean) As String

        If (IsMissing(createFile)) Then createFile = False
    readFileAndTruncate = Left(readFile(File, createFile), 30000)

End Function

'---------------------------------------------------------------------------------------------------------------------------------------------
'       + readFileAndTruncate(ByVal file As String) As String :
'           * Description : It is used to validate file path/link from a string variable. Return True if the file exist.
'           * Specifications / limitations
'               - The file has to Exists, no error handling
'               - The content is retrieved without any line returns (line returns are replaced by space)
'               - Only the 30.000 first characters are retrieved
'           * Arguments
'               - strFileFullPath As String : Full path to the file input file including the file name and the file extension.
'           * Output
'
'
'
'       Last edition date : 11/07/2012
'
'       Revisions history
'       -----------------
'           - Georgiev Velin    20/12/2011      Creation    velin.georgiev@gmail.com
'           - Emile Fyon        11/07/2012      Creation    emilefyon@gmail.com
'
'
'---------------------------------------------------------------------------------------------------------------------------------------------

Function fileExists(strFileFullPath As String) As Boolean


    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.fileExists(strFileFullPath) Then fileExists = True
    Set objFSO = Nothing


End Function

