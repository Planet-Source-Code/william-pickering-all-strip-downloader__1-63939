Attribute VB_Name = "modINIHandling"
'API Declares
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'This function gets the setting
Function GetINISetting(strSectionHeader As String, strVariableName As String, strFileName As String, Optional strDefault As String = "") As String
Dim strReturn As String
strReturn = String(255, Chr(0))
GetINISetting = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, strDefault, strReturn, Len(strReturn), strFileName))
End Function

'this function saves a setting
Function WriteINISetting(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
WriteINISetting = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function

'this one deletes a setting
Function DelINISetting(txtSection As String, txtSetting As String, txtFile As String) As Long
DelINISetting = WritePrivateProfileString(txtSection, txtSetting, 0&, txtFile)
End Function

'and this one delets an entire section
Function DelINISection(txtSection As String, txtFile As String) As Long
DelINISection = WritePrivateProfileString(txtSection, 0&, 0&, txtFile)
End Function

'this sets a setting to blank
Function ClearINISetting(txtSection As String, txtSetting As String, txtFile As String) As Long
ClearINISetting = WritePrivateProfileString(txtSection, txtSetting, "", txtFile)
End Function

Function FileExist(dPath As String) As Boolean
On Error GoTo errorH

FileExist = False

If FileLen(dPath) >= 0 Then
FileExist = True
End If

Exit Function

errorH:
FileExist = False

End Function
