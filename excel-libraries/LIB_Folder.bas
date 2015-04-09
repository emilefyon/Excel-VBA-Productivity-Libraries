Attribute VB_Name = "LIB_Folder"
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
'       + Sub listFolder(ByVal strPath As String, byval cellDestination as range) As String
'           * Description : List the folders in the folder specified in argument and display the list of the folders in the cells
'                           below the cell given in argument
'
'           * Specifications / limitations
'               - Should be nice to create the folder (and parent folders) if it does not exists
'           * Arguments
'               - strPath as String : the full path of the folder
'               - cellDestination as Range : the destination cell
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
    If (counter = 0) Then Exit Sub
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
'           * Description : check that the folder has a trailing slash and add one if needed. Create the folders if needed
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
    strPath = Replace(strPath, "/", "\")
    strPath = Replace(strPath, "\\", "\")
    ' createDirs (strPath)
    checkFolder = strPath

End Function

'---------------------------------------------------------------------------------------------------------------------------------------------
'       + Sub listFolder(ByVal strPath As String, byval cellDestination as range) As String
'           * Description : List the folders in the folder specified in argument and display the list of the folders in the cells
'                           below the cell given in argument
'
'           * Specifications / limitations
'               - Should be nice to create the folder (and parent folders) if it does not exists
'           * Arguments
'               - strPath as String : the full path of the folder
'               - cellDestination as Range : the destination cell
'
'       Last edition date : 27/07/2012
'
'       Revisions history
'       -----------------
'           - Emile Fyon        27/07/2012      Creation
'
'---------------------------------------------------------------------------------------------------------------------------------------------


Sub ListFolder(sFolderPath As String, ByVal cellDestination As Range)
     
    Dim fs As New FileSystemObject
    Dim FSfolder As Folder
    Dim subfolder As Folder
    Dim i As Integer
     
    Set FSfolder = fs.GetFolder(sFolderPath)
    
    i = 0
    For Each subfolder In FSfolder.SubFolders
        DoEvents
        i = i + 1
        cellDestination.Offset(i, 0) = subfolder.Name
    Next subfolder
     
    Set FSfolder = Nothing

     
End Sub

Function getOldestFileInDir(ByVal path As String, ByVal fileNameMask As String) As Date
  
  Dim FileName As String, FileDir As String, FileSearch As String
  Dim MaxDate As Date, interDate As Date, dteFile As Date
  
  MaxDate = DateSerial(1900, 1, 1)
  FileDir = path
  FileName = fileNameMask
  FileSearch = Dir(FileDir & FileName)
  
  While Len(FileSearch) > 0
    dteFile = FileDateTime(FileDir & FileSearch)
    If dteFile > MaxDate Then
      MaxDate = dteFile
    End If
    FileSearch = Dir()
  Wend
  
  getOldestFileInDir = MaxDate

End Function

'-----------------------------------------------


Function createFolder(ByVal fullPath As String) As Boolean

    
    If (folderExists(fullPath) = False) Then MkDir (fullPath)
     
End Function

Function folderExists(ByVal fullPath As String) As String

Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
folderExists = fs.folderExists(fullPath)

End Function


Function createDirs(ByVal fullPath As String) As String

    fullPath = checkFolder(fullPath)
    paths = Split(fullPath, "\")
    
    currentPath = paths(0) & "\"
    folderCreated = 0
    For i = 1 To UBound(paths) - 1
        currentPath = currentPath & paths(i) & "\"
        If folderExists(currentPath) = False Then
            createFolder (currentPath)
            folderCreated = folderCreated + 1
        End If
    Next
    createDirs = folderCreated & " folder(s) has/have been generated"
    

End Function




'---------------------------------------------------------


 
Sub TestListFolders()
     
    Application.ScreenUpdating = False
     
     'create a new workbook for the folder list
     
     'commented out by dr
     'Workbooks.Add
     
     'line added by dr to clear old data
    cells.Delete
     
     ' add headers
    With Range("A1")
        .Formula = "Folder contents:"
        .Font.Bold = True
        .Font.Size = 12
    End With
     
    Range("A3").Formula = "Folder Path:"
    Range("B3").Formula = "Folder Name:"
    Range("C3").Formula = "Size:"
    Range("D3").Formula = "Subfolders:"
    Range("E3").Formula = "Files:"
    Range("F3").Formula = "Short Name:"
    Range("G3").Formula = "Short Path:"
    Range("A3:G3").Font.Bold = True
     
     'ENTER START FOLDER HERE
     ' and include subfolders (true/false)
    listFoldersFullInfo "H:\User\02. Projects\", False
     
    Application.ScreenUpdating = True
     
End Sub
 
Sub listFoldersFullInfo(SourceFolderName As String, IncludeSubfolders As Boolean)
     ' lists information about the folders in SourceFolder
     ' example: ListFolders "C:\", True
    Dim FSO As Scripting.FileSystemObject
    Dim SourceFolder As Scripting.Folder, subfolder As Scripting.Folder
    Dim r As Long
     
    Set FSO = New Scripting.FileSystemObject
    Set SourceFolder = FSO.GetFolder(SourceFolderName)
     
     'line added by dr for repeated "Permission Denied" errors
     
    On Error Resume Next
     
     ' display folder properties
    r = Range("A65536").End(xlUp).Row + 1
    cells(r, 1).Formula = SourceFolder.Path
    cells(r, 2).Formula = SourceFolder.Name
    cells(r, 3).Formula = SourceFolder.Size
    cells(r, 4).Formula = SourceFolder.SubFolders.Count
    cells(r, 5).Formula = SourceFolder.Files.Count
    cells(r, 6).Formula = SourceFolder.ShortName
    cells(r, 7).Formula = SourceFolder.ShortPath
    If IncludeSubfolders Then
        For Each subfolder In SourceFolder.SubFolders
            listFolders subfolder.Path, True
        Next subfolder
        Set subfolder = Nothing
    End If
     
    Columns("A:G").AutoFit
     
    Set SourceFolder = Nothing
    Set FSO = Nothing
     
     'commented out by dr
     'ActiveWorkbook.Saved = True
     
End Sub



