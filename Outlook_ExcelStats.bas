Attribute VB_Name = "excelStats"

' Add Microsoft Excel 14.0 Object Library Reference

Sub ExportStatsToExcel()
    ' On Error GoTo ErrHandler
    
    Dim appExcel As Excel.Application
    Dim wkb As Excel.Workbook
    Dim wks As Excel.Worksheet
    Dim rng As Excel.Range
    Dim strSheet As String
    Dim strPath As String
    Dim intRowCounter As Integer
    Dim intColumnCounter As Integer
    Dim msg As Outlook.MailItem
    Dim nms As Outlook.NameSpace
    Dim fld As Outlook.MAPIFolder
    Dim itm As Object
    
    Debug.Print strSheet
    
    'Select export folder
    Set nms = Application.GetNamespace("MAPI")
    Set fld = nms.PickFolder
    
    'Handle potential errors with Select Folder dialog box.
    If fld Is Nothing Then
        MsgBox "There are no mail messages to export", vbOKOnly, _
        "Error"
        Exit Sub
        
    ElseIf fld.DefaultItemType <> olMailItem Then
        MsgBox "There are no mail messages to export", vbOKOnly, _
        "Error"
        Exit Sub
        
    ElseIf fld.Items.Count = 0 Then
        MsgBox "There are no mail messages to export", vbOKOnly, _
        "Error"
        Exit Sub
    End If
    
    'Open and activate Excel workbook.
    Set appExcel = CreateObject("Excel.Application")
    Set wkb = appExcel.Workbooks.Add
    Set wks = wkb.Sheets(1)
    wks.Activate
    appExcel.Application.Visible = True
    
    ' Copy headers
    wks.Cells(1, 1).Value = "Folder"
    wks.Cells(1, 2).Value = "Sender"
    wks.Cells(1, 3).Value = "Subject"
    wks.Cells(1, 4).Value = "Sent time"
    wks.Cells(1, 5).Value = "Received time"
    wks.Cells(1, 6).Value = "Size (ko)"
    
    
    'Copy field items in mail folder.
    
    recursive = True
    If recursive = True Then
        Call processFolder(fld, wks)
    Else
    
        intRowCounter = 0
        For Each itm In fld.Items
            intColumnCounter = 1
            intRowCounter = intRowCounter + 1
            
            wks.Cells(intRowCounter, 1).Value = itm.To
            wks.Cells(intRowCounter, 2).Value = itm.SenderEmailAddress
            wks.Cells(intRowCounter, 3).Value = itm.Subject
            wks.Cells(intRowCounter, 4).Value = itm.SentOn
            wks.Cells(intRowCounter, 5).Value = itm.ReceivedTime
            wks.Cells(intRowCounter, 6).Value = itm.Size
            
        Next
    End If
    
    MsgBox ("Import done")
    appExcel.Application.Visible = True

    Set appExcel = Nothing
    Set wkb = Nothing
    Set wks = Nothing
    Set rng = Nothing
    Set msg = Nothing
    Set nms = Nothing
    Set fld = Nothing
    Set itm = Nothing
    Exit Sub

ErrHandler:      If Err.Number = 1004 Then
    MsgBox strSheet & " doesn't exist", vbOKOnly, _
    "Error"
    Else
    MsgBox Err.Number & "; Description: ", vbOKOnly, _
    "Error"
    End If
    
    Set appExcel = Nothing
    Set wkb = Nothing
    Set wks = Nothing
    Set rng = Nothing
    Set msg = Nothing
    Set nms = Nothing
    Set fld = Nothing
    Set itm = Nothing
End Sub


Private Sub processFolder(ByVal oParent As Outlook.MAPIFolder, ByVal wks As Worksheet)

        Debug.Print oParent.Name & "|" & oParent.Items.Count
        
        Dim oFolder As Outlook.MAPIFolder
        Dim oMail As Outlook.MailItem

        

        'Get your data here ...
        If wks.Range("A1") = "" Then
            intRowCounter = 1
        ElseIf wks.Range("A3") = "" Then
            intRowCounter = 2
        Else
            intRowCounter = wks.Range("A1").End(xlDown).Offset(1, 0).Row
        End If
        For Each itm In oParent.Items
            intColumnCounter = 1
            
            wks.Cells(intRowCounter, 1).Value = oParent.Name
            wks.Cells(intRowCounter, 2).Value = itm.SenderName
            wks.Cells(intRowCounter, 3).Value = itm.Subject
            wks.Cells(intRowCounter, 4).Value = itm.SentOn
            wks.Cells(intRowCounter, 5).Value = itm.ReceivedTime
            wks.Cells(intRowCounter, 6).Value = itm.Size / 1024
            
            intRowCounter = intRowCounter + 1
        Next itm


        If (oParent.Folders.Count > 0) Then
            For Each oFolder In oParent.Folders
                Call processFolder(oFolder, wks)
                
            Next
        End If
End Sub
