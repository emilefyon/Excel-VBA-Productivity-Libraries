Attribute VB_Name = "LIB_ToImport"
'Title:libSimpleVBAFunctions
'Description: Very simple Visual Basic for Applications (VBA) functions
'for Excel and Access reports automation.
'All the functions are enclosed in this file and they can be included to every automation project.


'@Name: SheetExists
'@Description: Checks if a Worksheet exist.
'@Version: 1.0
'@Autor: velin.georgiev@gmail.com
'@Date: 20.12.2011
'Input parameters:
    '@Param strWSName: String. Worksheet name.
'Output patameters:
    '@Param SheetExists: Boolean.
Function SheetExists(strWSName As String) As Boolean
    
    Dim intCountSheet As Integer
    For intCountSheet = 1 To ActiveWorkbook.Sheets.Count
        If LCase(Sheets(intCountSheet).Name) = LCase(strWSName) Then
            SheetExists = True
            Exit Function
        End If
    Next intCountSheet
    
End Function





'@Name: FolderExists
'@Description: It is used to validate file path from a string variable. Returns True if the folder exist.
'@Version:1.0
'@Autor:velin.georgiev@gmail.com
'@Date: 20.12.2011
'Input parameters
    '@Param strFileFullPath: String. Full path to the folder.
    'Local path example: "C:/myFolder"
    'SharePoint path example: "\\my.sharepoint.com\Shared%20Documents"
        'Local network path example: "\\SOMEONES-PC\Users\Public\Documents"
'Output Parameters
    '@Param FolderExists: Boolean.
Function folderExists(strFolderPath As String) As Boolean


    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.folderExists(strFolderPath) Then folderExists = True
    Set objFSO = Nothing


End Function


'@Name: GetFileName
'@Description: Used to get the file name (with the extension) from a string variable that refers to a full path(local or network).
'Please ensure that the input variable contains a file name at the end of the full path string,
' because the function would catch the string after the last slash "/" or "\".
' If the input is incorrect it may result in a directory for example: \\sharepoint.com\dir so the result would be 'dir'
'@Version:1.0
'@Autor:
'@Date: 20.12.2011
'Input parameters
    '@Param strFilePath: String. Contains full path for a file including the file name and the extension.
    'Local path example: "C:/myFile.txt"
    'SharePoint path example: "\\my.sharepoint.com\Shared%20Documents\Raw_Data.xlsx"
        'Local network path example: "\\SOMEONES-PC\Users\Public\Documents\Raw_Data.xlsx"
'Output Parameters
    '@Param GetFileName: String. Contains the file name including the file extension or empty string. Example: Book1.xls
Function GetFileName(strFilePath As String) As String


    Dim strFileName As String
    GetFileName = ""
    If InStr(1, strFilePath, "\") > 0 Then
        strFileName = Split(strFilePath, "\")(UBound(Split(strFilePath, "\")))
        GetFileName = strFileName
    ElseIf InStr(1, strFilePath, "/") > 0 Then
        strFileName = Split(strFilePath, "/")(UBound(Split(strFilePath, "/")))
        GetFileName = strFileName
    End If


End Function


'@Name: IsWbOpen
'@Description: Used to check if an instance of an excel file is open by excel.
'@Version:1.0
'@Autor: http://www.vbaexpress.com/kb/getarticle.php?kb_id=443 'Zack Barresse
'@Date:  20.12.2011
'Input parameters
    '@Param wbName: String. Contains the file name including the file extension. Example: Book1.xls
'Output Parameters
    '@Param IsWbOpen: Boolean.
Function IsWbOpen(wbName As String) As Boolean
    
    Dim intCountWb As Integer
    For intCountWb = Workbooks.Count To 1 Step -1
        If Workbooks(intCountWb).Name = wbName Then
            IsWbOpen = True
            Exit Function
        End If
    Next intCountWb
    
End Function


'@Name: FindString
'@Description: Used to catch a string in a cell that match and activate the cell.
' It needs to be applied range before use of the function.
'@Version:1.0
'@Autor: velin.georgiev@gmail.com
'@Date: 20.12.2011
'Input parameters
    '@Param strFind: String. Word that has to be performed search for.
'Output Parameters
    '@Param FindString: Boolean
        '@Action. Activates the first match (search order is by column).
Function FindString(strFind As String) As Boolean
    
    Dim FoundRange
    Set FoundRange = cells.Find(What:=strFind, After:=ActiveCell, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)
    If Not (FoundRange Is Nothing) Then
        cells.FindNext(After:=ActiveCell).Select
        FindString = True
    End If
    
End Function


'@Name: Zip
'@Description: Creates a Zip file and returns the Zip file path.
'@Version:1.0
'@Autor:velin.georgiev@gmail.com 'keepITcool  'http://www.rondebruin.nl/windowsxpzip.htm
'@Date: 20.12.2011
'Input parameters
    '@Param strZipPath: String. Contains folder full path where the zip file should be stored
    '@Param strFilePath1: String. Contains the full path to the file that should be Zipped including the file name and exstension. Example: C:\Book1.xls
    '@Param strZipFileName: Optional String. Name of the Zip. If this param is empty the function will apply as a name the name of the file that should be zipped.
'Output Parameters
    '@Param Zip: String. Contains file full path on the zip file including the file name and extension.
        '@Action: Creates a Zip file.
Function Zip(strZipPath As String, strFilePath1 As String, Optional strZipFileName As String) As String
    
    Dim objApp As Object
    Dim intCount As Integer
    Dim arryFiles, ZipFile
    arryFiles = Array(strFilePath1) 'You can add additional param strFilePath2 as function input and add it to the array so it would zip two files...
    If Right(strZipPath, 1) <> "\" Then strZipPath = strZipPath & "\"
    If strZipFileName <> "" Then
        ZipFile = strZipPath & strZipFileName & ".zip"
    Else
        ZipFile = strZipPath & GetFileName(strFilePath1) & ".zip"
    End If
    If IsArray(arryFiles) Then
        If Len(Dir(ZipFile)) > 0 Then Kill ZipFile
        Open ZipFile For Output As #1
        Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
        Close #1
        Set objApp = CreateObject("Shell.Application")


        For intCount = LBound(arryFiles) To UBound(arryFiles)
            objApp.Namespace(ZipFile).CopyHere arryFiles(intCount)
        Next intCount
    End If
    Set objApp = Nothing
    Zip = ZipFile
    
End Function


'@Name: Mail
'@Description: Display or send an email through the Microsoft Outlook client. You should have MS Outlook installed and configured
'Note: Please note when you set "True" to the 4th argument the function will use your outlook to send directly the email.
'If you get warrning about a program accessing e-mail address information or sending e-mail on my behalf you can visit this link:
'http://office.microsoft.com/en-us/outlook-help/i-get-warnings-about-a-program-accessing-e-mail-address-information-or-sending-e-mail-on-my-behalf-HA001229943.aspx
'@Version:1.0
'@Autor:velin.georgiev@gmail.com 'Dick Kusleika 'http://www.rondebruin.nl/mail/folder3/signature.htm
'@Date: 20.12.2011
'Input parameters
    '@Param strTo: String. TO recipients of the email. Example: 'velin.georgiev@gmail.com;john.smith@hotmail.com'
    '@Param strSubject: String. The subject of the email.
    '@Param strBody: Optional String. The body content of the email.
    '@Param bSend: Optional Boolean. True would directly send the email. False would open the email in display mode in the outlook.
    '@Param strCC: Optional String. CC recipients of the email. Example: 'velin.georgiev@gmail.com;john.smith@hotmail.com'
    '@Param strSignName: Optional String.
    '@Param strAttachPath1: Optional String. Email attachment.
    '@Param strAttachPath2: Optional String. Email attachment.
    '@Param strAttachPath3: Optional String. Email attachment.
    '@Param strAttachPath4: Optional String. Email attachment.
    '@Param strAttachPath5: Optional String. Email attachment.
'Output Parameters
    '@Action. Displays or sends an email though the MS Outlook
Function Mail( _
    strTo As String, _
    strSubject As String, _
    strBody As String, _
    Optional bSend As Boolean, _
    Optional strCC As String, _
    Optional strSignName As String, _
    Optional strAttachPath1 As String, _
    Optional strAttachPath2 As String, _
    Optional strAttachPath3 As String, _
    Optional strAttachPath4 As String, _
    Optional strAttachPath5 As String)
        
    Dim objOutApp As Object
    Dim objOutMail As Object
    Dim objFSO As Object
    Dim txtStream As Object
    Dim Signature As String
    Dim strSignature As String
    
    Set objOutApp = CreateObject("Outlook.Application")
    Set objOutMail = objOutApp.CreateItem(0)
    'Get the outlook signature by its default path
    strSignature = "C:\Documents and Settings\" & Environ("username") & "\Application Data\Microsoft\Signatures\" & strSignName & ".htm"
    'strSignature = "C:\Users\" & Environ("username") & "\AppData\Roaming\Microsoft\Signatures\Mysig.htm"


    If Dir(strSignature) <> "" Then
        'Dick Kusleika
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set txtStream = objFSO.GetFile(strSignature).OpenAsTextStream(1, -2)
        Signature = txtStream.readall
        txtStream.Close
    End If


    On Error Resume Next
    With objOutMail
        .To = strTo
        .CC = strCC
        '.BCC = strBCC
        .Subject = strSubject
        .HTMLBody = "<style type='text/css'>.style1{font-family:'Futura Bk',Times,serif;font-size:95%;}</style><div class='style1'>" & strBody & "</div>" & Signature
        'You can add files also like this
        If strAttachPath1 <> "" Then .Attachments.Add (strAttachPath1)
        If strAttachPath2 <> "" Then .Attachments.Add (strAttachPath2)
        If strAttachPath3 <> "" Then .Attachments.Add (strAttachPath3)
        If strAttachPath4 <> "" Then .Attachments.Add (strAttachPath4)
        If strAttachPath5 <> "" Then .Attachments.Add (strAttachPath5)
        If bSend Then
            .Send
        Else
            .Display
        End If
    End With


    On Error GoTo 0
    Set objOutMail = Nothing
    Set objOutApp = Nothing
        
End Function


'@Name: GetFilePath
'@Description: Opens a file search window. Used to get the file full path.
'@Version:1.0
'@Autor:velin.georgiev@gmail.com
'@Date: 20.12.2011
'Input parameters
    '@Param strWindowMsg: Optional String. Message that would appear at the top of the search window.
'Output Parameters
    '@Param GetFilePath: String. The full path of the file including the file name and extension.
Function BrowseForFile(Optional strWindowMsg As String) As String


    Dim strWindowFilter As String
    If strWindowMsg = "" Then strWindowMsg = "Please select file."
    strWindowFilter = "All Files (*.*),*.*,Excel 2007 Files (*.xlsx),*.xlsx,Excel Files (*.xls),*.xls,Excel Macro Enabled Files (*.xlsm),*.xlsm"
    BrowseForFile = Application.GetOpenFilename(strWindowFilter, , strWindowMsg, , False)


End Function


'@Name: BrowseForFolder
'@Description: Opens a folder search window. Returns path to a folder.
'@Version:1.0
'@Autor:unknown, but provided by Margarita Yordanova
'@Date: 22.02.2011
'Input parameters
    '@Param Hwnd: Long.
    '@Param sTitle: String.
    '@Param BIF_Options: Intager.
    '@Param vRootFolder: Variant.
'Output Parameters
    '@Param BrowseForFolderShell: String. Returns path to a folder.
'// Minimum DLL version shell32.dll version 4.71 or later
'// Minimum operating systems   Windows 2000, Windows NT 4.0 with Internet Explorer 4.0,
'// Windows 98, Windows 95 with Internet Explorer 4.0
'// objFolder = objShell.BrowseForFolder(Hwnd, sTitle, BIF_Options [, vRootFolder])


Public Function BrowseForFolder( _
    Optional Hwnd As Long = 0, _
    Optional sTitle As String = "Please, select a folder", _
    Optional BIF_Options As Integer, _
    Optional vRootFolder As Variant) As String


    'Optional BIF_Options As Integer = BIF_VALIDATE, _


    Dim objShell As Object
    Dim objFolder As Variant
    Dim strFolderFullPath As String


    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(Hwnd, sTitle, BIF_Options, vRootFolder)


    If (Not objFolder Is Nothing) Then
        '// NB: If SpecFolder= 0 = Desktop then ....
        On Error Resume Next
        If IsError(objFolder.Items.Item.Path) Then strFolderFullPath = CStr(objFolder): GoTo GotIt
        On Error GoTo 0
        '// Is it the Root Dir?...if so change
        If Len(objFolder.Items.Item.Path) > 3 Then
            strFolderFullPath = objFolder.Items.Item.Path & Application.PathSeparator
        Else
            strFolderFullPath = objFolder.Items.Item.Path
        End If
    Else
    '// User cancelled
        GoTo XitProperly
    End If


GotIt:
    BrowseForFolder = strFolderFullPath
    
XitProperly:
    Set objFolder = Nothing
    Set objShell = Nothing


End Function


'@Name: GetDate
'@Description: Used to format the date the way it needs to stand in my reports.
'This function has been added to the lib because my everyday automated excel reports had to be renamed with a different date and format.
'so I decided to put this in one function.
'@Version:1.0
'@Autor:velin.georgiev@gmail.com
'@Date: 20.12.2011
'Input parameters
    '@Param strFormat: String. Implements date format. Example: "dd-mm-yyyy","mm-yy","mmmm-yyyy"
    '@Param strDiversionFrom: String. Possible values are "d" as day, "m" as month, "yyyy" as year.
    '@Param intDiversionValue: Integer. Implements timeframe (past, now, future) that have to be subtracted from or added to todays date.
    'The default is 0 = Now.
'Output Parameters
    '@Param GetDate: String. Formated date.
'@Note:
'@Example: For example if you specify @Param strDiversionFrom = "m" and intDiversionValue = -1.
'It will show you the current month( now) - 1. Result is previous month.
'dDate = GetDate("MMMM", "m", -1), The result shown the MONTH only ,but one MONTH past now (For example: December)
'dDate = GetDate("dd-mmm", "m", -1).The result should be something like 10-Dec.
'dDate = GetDate("mmm-yy", "m", -1).The result should be something like Dec-10.
'dDate = GetDate("dd-mmm-yy", "m").The result should be something like 10-Jan-10.
'dDate = GetDate("dd-mmm-yy").The result should be something like 10-Jan-10.
Function GetDate(Optional strDateFormat As String, Optional strDiversionFrom As String, Optional intDiversionValue As Integer) As String


    Dim dDate As Date
    If strDateFormat = "" Then strDateFormat = " mmm-yy"
    If strDiversionFrom = "" Then strDiversionFrom = "m"
    dDate = DateAdd(strDiversionFrom, intDiversionValue, CDate(Now))
    GetDate = Format(dDate, strDateFormat)
    
End Function


'@Name: Copy
'@Description: Copy file or files from a folder in another destination.
'Note: If the file already exist it will overwrite existing files in this folder.
'@Version:1.0
'@Autor:unknown
'@Date: 20.12.2011
'Input parameters
    '@Param strSourceFullPath: String. The full path to the file (that has to be copied). If you would like to copy all files
    'from a folder you can type *.* right after the source folder slash.
    'Example: strSourceFullPath="C:/*.*" would copy all files in "C:/"
    'Example: strSourceFullPath="C:/*.xlsx" would copy all excel 2007 files in "C:/"
    '@Param strCopyToDestination: String. The full path to a folder.
'Output Parameters
    '@Action. Performs a copy.
    '@Param Copy: String. Returns a message that a copy has been performed.
Function Copy(strSourceFullPath As String, strCopyToDestination As String) As String
    
    Dim objFSO As Object
    If Right(strCopyToDestination, 1) <> "\" Then strCopyToDestination = strCopyToDestination & "\"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If GetFileName(strSourceFullPath) <> "" And objFSO.folderExists(strCopyToDestination) Then
        objFSO.CopyFile Source:=strSourceFullPath, Destination:=strCopyToDestination, overwritefiles:=True
    End If
    
    Set objFSO = Nothing
    Copy = "CopyFile performed. Source=" & strSourceFullPath & " Destination=" & strCopyToDestination
    
End Function


'@Name: PivotCacheRefresh
'@Description: Used to refresh all pivot cached lists and clean up the 'dirty' xml cached items.
'I am using this function on Workbook_Open() when a pivot contains cached items from some outdated data.
'@Version:1.0
'@Autor: Margarita Yordanova
'@Date: 20.12.2011
'Input parameters
    '@Param:
'Output Parameters
    '@Param PivotCacheRefresh: Boolean.
Function PivotCacheRefresh() As Boolean


    Dim pvt
    For Each pvt In ActiveSheet.PivotTables
        pvt.PivotCache.MissingItemsLimit = xlMissingItemsNone
        pvt.PivotCache.Refresh
    Next pvt
    ThisWorkbook.RefreshAll
    PivotCacheRefresh = True
    
End Function


'@Name: OpenURL
'@Description: Opens an url in the internet explorer. This function could be assigned to button.
'@Version:1.0
'@Autor:
'@Date: 20.12.2011
'Input parameters
    '@Param strURL: String. Url.
'Output Parameters
    'Action: Opens the IExplorer with url address the input argument.
        '@Param OpenURL: String. The opened link.
Function OpenURL(strURL As String) As String
    
    Dim objIE As Object
    Set objIE = CreateObject("Internetexplorer.Application")
    objIE.Visible = True
    objIE.Navigate strURL
    Set objIE = Nothing
    
    OpenURL = strURL
    
End Function


'@Name: NewLog
'@Description: Used to create a sheet that would be used by the function Log() where log information would be stored.
'@Version: 1.0
'@Autor: velin.georgiev@gmail.com
'@Date: 20.12.2011
'Input parameters:
    '@Param strSheetName: Optional String. Name of a sheet where the log information where log information would be stored. Default: Log
'Output patameters:
    '@Param NewLog: Boolean. True is success.
        '@Param: Action. Creates a sheet.
    
Function NewLog(Optional strSheetName As String) As Boolean
    
    Dim func
    ThisWorkbook.Activate
    If strSheetName = "" Then strSheetName = "Log"
    'SheetExists function should be available within this module
    If SheetExists(strSheetName) = False Then
        ThisWorkbook.Worksheets.Add.Name = strSheetName
        With ThisWorkbook.Worksheets(strSheetName)
            .cells(1, 1).Value = "Date"
            .cells(1, 2).Value = "Time"
            .cells(1, 3).Value = "Log"
        End With
        func = Log("New log named " & strSheetName & " has been created.")
                NewLog = True
    End If
    
End Function


'@Name: Log
'@Description: Used to record action notes (from macro execution) in a sheet named 'Log' by default.
' It depends on the programmer what should be logged in the log
' so this function is used depending on the programmer needs.
' It can be applied on every line of the macro if an event information has to be recorded in the Log sheet.
' Before the use of this function a new sheet should be created to store the logs. The NewLog() function could do this for you.
'@Version: 1.0
'@Autor: velin.georgiev@gmail.com
'@Date: 20.12.2011
'Input parameters:
    '@Param strLogInfo: String. Describes error or some information of a taken action or event within the vba module.
    '@Param strSheetName: Optional String. Name of a sheet where the log information would be stored.
'Output patameters:
    '@Param Date: Data Record. Enters the current date in the Log sheet
    '@Param Time: Data Record. Enters the current time in the Log sheet
    '@Param strLogInfo: Data Record. Enters the strLogInfo string as a text in the Log sheet
Function Log(strLogInfo As String, Optional strSheetName As String) As String


    Dim rngLastRow
    If strSheetName = "" Then strSheetName = "Log"
    rngLastRow = ThisWorkbook.Worksheets(strSheetName).UsedRange.Rows.Count + 1
    With ThisWorkbook.Worksheets(strSheetName)
        .cells(rngLastRow, 1).Value = Date
        .cells(rngLastRow, 2).Value = Time
        .cells(rngLastRow, 3).Value = strLogInfo
    End With
    Log = "Date=" & Date & " Time=" & Time & " Message=" & strLogInfo
    
End Function



