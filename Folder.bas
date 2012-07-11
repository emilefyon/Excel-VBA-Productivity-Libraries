Attribute VB_Name = "Folder"
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
    Dim File As String
    ' Dim filesTab
    
    ' Add a trailing slash if needed
    strPath = checkFolder(strPath)
    
    File = Dir$(strPath & Extention)
    
    ' Count the number of files in the folder
    Do While Len(File)
        File = Dir$
        counter = counter + 1
    Loop
    ReDim filesTab(counter - 1)
    counter = 0

    ' Reset the counter of the array
    File = Dir$(strPath & Extention)

    ' List the files and display them in the cells
    Do While Len(File) And counter <= UBound(filesTab)
        cellDestination.Offset(counter, 0) = File
        cellDestination.Offset(counter, 1) = strPath & File
        File = Dir$
        counter = counter + 1
    Loop
    


End Sub


'---------------------------------------------------------------------------------------------------------------------------------------------
'       + checkFolder(ByVal strPath As String) As String
'           * Description : check that the folder has a trailing slash and add one if needed
'
'           * Specifications / limitations
'               - Should be nice to create the folder (and parent folders) if it does not exists
'           * Arguments
'               - strPath as String : the full path of the folder
'
'       Last edition date : 11/07/2012
'
'       Revisions history
'       -----------------
'           - Emile Fyon        11/07/2012      Creation
'
'---------------------------------------------------------------------------------------------------------------------------------------------


Function checkFolder(ByVal strPath As String) As String
    
    ' Add a trailing slash if needed
    If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"
    checkFolder = strPath

End Function

