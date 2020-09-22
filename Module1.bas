Attribute VB_Name = "CommonFunction"
Option Explicit

Public FavouriteDir As String
Public BackupDir As String
Public vbDialog As New vbDialog
Public vbFileIO As New vbFileIO
Public vbComboBox As New vbComboBox

Public ListLine() As String
Public ListIndex() As Long
Public ListCtr As Long
Public ListBM() As String


Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public dlgSettings As vbDialog_MESSAGEDLG_SETTINGS
Public cData As vbHTTP_CONNECTION_INFO

Public vbSystem As New vbSystem
Public des As String
Public vbHTTP As New vbHTTP
Public ResultGet As Boolean


Function GetDesc(ResponseCode As Integer) As String
Select Case ResponseCode
Case 200:
GetDesc = "Link Active"
Case 400:
GetDesc = "Bad Request"
Case 403:
GetDesc = "Forbidden"
Case 404:
GetDesc = "Not Found"
Case 501:
GetDesc = "Not Implemented"
Case 408:
GetDesc = "Time Out"
Case 502:
GetDesc = "Time Out"
Case 503:
GetDesc = "Not Found"
Case 404:
GetDesc = "Not Found"
Case 504:
GetDesc = "Cannot Reach"
Case 407:
GetDesc = "Authorization Error"
Case 302
GetDesc = "Temporarily Moved"
Case Else
GetDesc = "Not Defined"
End Select
End Function

Sub UpdateLink(Path As String, LinkName As String, NewURL As String)
Open Path & "\" & LinkName & ".url" For Output As #1
Print #1, "[DEFAULT]"
Print #1, "BASEURL=" & NewURL
Print #1, "[InternetShortcut]"
Print #1, "URL=" & NewURL
Close #1
End Sub

Sub CreateLink(Path As String, LinkName As String, NewURL As String)
Open Path & "\" & LinkName & ".url" For Output As #1
Print #1, "[DEFAULT]"
Print #1, NewURL
Print #1, "[InternetShortcut]"
Print #1, "URL=" & NewURL
Close #1
End Sub

Sub DeleteLink(Path As String, LinkName As String)
Kill Path & "\" & LinkName & ".url"
End Sub

Sub OpenURL(url As String)
    Dim lngReturnNumber As Long
    
    'Launch File
    lngReturnNumber = ShellExecLaunchFile(url, "", "")
    
    'Check for Errors
    If lngReturnNumber < 33 Then
        Call ShellExecLaunchErr(lngReturnNumber, True)
        Exit Sub
    End If
End Sub
