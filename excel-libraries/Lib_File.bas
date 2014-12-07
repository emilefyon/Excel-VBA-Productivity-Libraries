Attribute VB_Name = "LIB_File"

' Documentation on this library can be found on https://github.com/emilefyon/Excel-VBA-Productivity-Libraries/blob/master/docs/Lib_File.md

Function writeFile(ByVal File As String, ByVal content As String) As String
    
   Open File For Output As #1
   Print #1, content
   Close #1
   
   writeFile = "File updated yet"

End Function


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


Function readFileAndTruncate(ByVal File As String, Optional createFile As Boolean) As String

        If (IsMissing(createFile)) Then createFile = False
    readFileAndTruncate = Left(readFile(File, createFile), 30000)

End Function


Function fileExists(file As String) As Boolean

	fileExists = false
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.fileExists(file) Then fileExists = True
    Set objFSO = Nothing

End Function

Function getFileUpdateTime(ByVal fullPath As String)
    
    getFileUpdateTime = FileDateTime(fullPath)

End Function
