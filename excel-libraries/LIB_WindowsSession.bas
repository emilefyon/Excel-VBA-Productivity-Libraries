Attribute VB_Name = "LIB_WindowsSession"
' Work with 32bit version
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
 (ByVal lpBuffer As String, nSize As Long) As Long

Public Function GetLogonName() As String
 
' Dimension variables
Dim lpBuff As String * 255
Dim ret As Long

' Get the user name minus any trailing spaces found in the name.
ret = GetUserName(lpBuff, 255)

If ret > 0 Then
  GetLogonName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
Else
  GetLogonName = vbNullString
End If

End Function


