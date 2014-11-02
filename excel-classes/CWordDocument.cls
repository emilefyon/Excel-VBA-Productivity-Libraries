VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWordDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private pDocumentFullPath As String
Private pAppWrd As Object
Private pFileSystem As Object
Private pObjectDocument As Object

Public Property Get documentFullPath() As String
    documentFullPath = pDocumentFullPath
End Property

Public Property Let documentFullPath(value As String)
    pDocumentFullPath = value
    Call openDocument
End Property



Public Sub openDocument()

    Set pFileSystem = CreateObject("Scripting.FileSystemObject")
    
    'Create an instance of Word for us to use
    Set pAppWrd = CreateObject("Word.Application")
    
    'Open our specified Word file, On Error is used in case the file is not there
    On Error Resume Next
        Set pObjectDocument = pAppWrd.Documents.Open(pDocumentFullPath)
    On Error GoTo 0

    'If the file is not found, we need to end the sub and let the user know
    If pObjectDocument Is Nothing Then
        MsgBox "Unable to find the Word file.", vbCritical, "File Not Found"
        pAppWrd.Quit
        Set pAppWrd = Nothing
        Exit Sub
    End If
    
        
    'Make our Word session visible
    pAppWrd.Visible = True
           
End Sub

Public Sub copySelection(Optional ByVal bookmarkName As String)

    If IsMissing(bookmarkName) = False And bookmarkName <> "" Then
        If checkBookmarksExist(bookmarkName) = False Then
            MsgBox ("The bookmark does not exists")
            Exit Sub
        End If
        pAppWrd.Selection.GoTo what:=wdGoToBookmark, Name:=bookmarkName
    End If
    Selection.Copy
    pAppWrd.Selection.PasteSpecial DataType:=wdPasteEnhancedMetafile, Placement:=wdInLine

End Sub



Public Function checkBookmarksExist(ByVal bookmarkName) As Boolean
    
    Dim b As bookMark
    
    checkBookmarksExist = False
    For Each b In pObjectDocument.Bookmarks
        If b.Name = bookmarkName Then checkBookmarksExist = True
    Next
    

End Function



