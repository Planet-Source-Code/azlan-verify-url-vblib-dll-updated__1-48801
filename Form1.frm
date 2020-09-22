VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E7E3DE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bookmark Cleaner"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FBF3EE&
      Caption         =   "Setting"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Program Setting"
      Top             =   0
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2280
      Top             =   1320
   End
   Begin MSWinsockLib.Winsock wscHttp 
      Left            =   2280
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1800
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FBF3EE&
      Caption         =   "Start Cleaning"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Start Bookmark Cleaning"
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FBF3EE&
      Caption         =   "Stop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Stop or Continue Cleaning"
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FBF3EE&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Quit Program"
      Top             =   0
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E7E3DE&
      Caption         =   "Show Erronous"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   5
      ToolTipText     =   "Show erronous links"
      Top             =   0
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00E7E3DE&
      Caption         =   "Show All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13440
      TabIndex        =   4
      ToolTipText     =   "Show all links"
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FBF3EE&
      Caption         =   "Restore All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Restore Favourite Folder with the backup"
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FBF3EE&
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Refresh Favourite List"
      Top             =   0
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7440
      Visible         =   0   'False
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1800
      Top             =   1320
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   6840
      Visible         =   0   'False
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   6255
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   11033
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   10711308
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Link Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Url"
         Object.Width           =   12348
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Response"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Action"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6255
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   11033
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   10711308
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Link Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Url"
         Object.Width           =   12348
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Response"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Action"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   7695
      Width           =   15015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Out"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   14160
      TabIndex        =   1
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu VerifyLink 
         Caption         =   "Verify Link"
      End
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu DeleteLink 
         Caption         =   "Delete Link"
      End
      Begin VB.Menu l2 
         Caption         =   "-"
      End
      Begin VB.Menu ModifyLink 
         Caption         =   "Modify Link"
      End
      Begin VB.Menu l3 
         Caption         =   "-"
      End
      Begin VB.Menu OpenLink 
         Caption         =   "Open"
      End
   End
   Begin VB.Menu Menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu RestoreLink 
         Caption         =   "Restore Link"
         Enabled         =   0   'False
      End
      Begin VB.Menu l4 
         Caption         =   "-"
      End
      Begin VB.Menu DeleteLink2 
         Caption         =   "Delete Link"
         Enabled         =   0   'False
      End
      Begin VB.Menu l5 
         Caption         =   "-"
      End
      Begin VB.Menu ChangeAddress 
         Caption         =   "Change Address"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CurrentIndex As Integer
Public IndexCount As Integer
Dim Quit As Boolean

Public Username As String
Public Password As String
Public Proxy As String
Public Port As String
Public url As String
Public HeaderCode As Integer
Public HeaderDesc As String
Public AltURL As String
Public JustOneItem As Boolean, TimeOut As Boolean, ForceStop As Boolean
Dim TimeOutCounter As Integer
Dim Started As Boolean, LastStop As Integer
Dim CurrentListIndex2 As Integer

Sub VerifyLinks()
ListView1.ListItems(IndexCount).Bold = True
ListView1.ListItems(IndexCount).ForeColor = vbBlue
ListView1.ListItems(IndexCount).ListSubItems(1).Bold = True
ListView1.ListItems(IndexCount).ListSubItems(1).ForeColor = vbBlue
Label4.Caption = "Validating " & ListView1.ListItems(IndexCount).SubItems(1)
ListView1.ListItems(IndexCount).SubItems(2) = "Checking..."
ListView1.Refresh
Timer1.Enabled = True   'Timeout starts here

    With wscHttp
        .Close
        .LocalPort = 0
        .Connect Proxy, Port
    End With
End Sub

Sub VerifyLinks2()
ListView2.ListItems(IndexCount).Bold = True
ListView2.ListItems(IndexCount).ForeColor = vbBlue
ListView2.ListItems(IndexCount).ListSubItems(1).Bold = True
ListView2.ListItems(IndexCount).ListSubItems(1).ForeColor = vbBlue
Label4.Caption = "Validating " & ListView2.ListItems(IndexCount).SubItems(1)
ListView2.ListItems(IndexCount).SubItems(2) = "Checking..."
ListView2.Refresh
Timer2.Enabled = True   'Timeout starts here

    With Winsock1
        .Close
        .LocalPort = 0
        .Connect Proxy, Port
    End With
End Sub

Private Sub ChangeAddress_Click()
Dim AddPath As String

AddPath = ListView2.ListItems(CurrentListIndex2).Tag

    If AddPath = "-" Then
    AddPath = FavouriteDir
    Else
    AddPath = FavouriteDir & "\" & AddPath
    End If
CommonFunction.UpdateLink AddPath, ListView2.ListItems(CurrentListIndex2).Text, ListView2.ListItems(CurrentListIndex2 + 1).SubItems(1)
ListView2.ListItems(CurrentListIndex2).SubItems(3) = "Changed"
End Sub

Private Sub Command1_Click()
RestoreFav
End Sub

Sub BackupFav()
   Dim FSO As New FileSystemObject
   On Error GoTo err:

   If FSO.FolderExists(FavouriteDir) = True Then
         FSO.CopyFolder FavouriteDir, BackupDir & "\Favourite", True
   End If
   Set FSO = Nothing
   Exit Sub
err:
   MsgBox err.Description
End Sub

Sub RestoreFav()
   Dim FSO As New FileSystemObject
   On Error GoTo err:

   If FSO.FolderExists(BackupDir & "\Favourite") = True Then
         FSO.CopyFolder BackupDir & "\Favourite", FavouriteDir, True
   End If
   Exit Sub
err:
   MsgBox err.Description
End Sub

Private Sub Command5_Click()
ListFavorite
'ListView2.ListItems.Clear
Option2.Value = True
End Sub


Private Sub DeleteLink2_Click()
Dim AddPath As String
If MsgBox("Are you sure to delete '" & ListView2.ListItems(CurrentListIndex2).Text & "' from the FAVOURITE list?", vbYesNo, "Delete Link") = vbNo Then Exit Sub
AddPath = ListView2.ListItems(CurrentListIndex2).Tag
    If AddPath = "-" Then
    AddPath = FavouriteDir
    Else
    AddPath = FavouriteDir & "\" & AddPath
    End If
CommonFunction.DeleteLink AddPath, ListView2.ListItems(CurrentListIndex2).Text
ListView2.ListItems(CurrentListIndex2).SubItems(3) = "Deleted"
End Sub

Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
CurrentListIndex2 = Item.Index
If CurrentListIndex2 > 1 Then
If Mid(ListView2.ListItems(CurrentListIndex2 - 1).SubItems(2), 1, 3) = "302" Then CurrentListIndex2 = CurrentListIndex2 - 1
End If
End Sub

Private Sub ListView2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 And ListView2.ListItems.Count > 0 Then

If Mid(ListView2.ListItems(CurrentListIndex2).SubItems(2), 1, 3) = "302" Then
DeleteLink2.Enabled = True
RestoreLink.Enabled = False
ChangeAddress.Enabled = True
ElseIf InStr(ListView2.ListItems(CurrentListIndex2).SubItems(3), "Prompt") > 0 Then
DeleteLink2.Enabled = True
RestoreLink.Enabled = False
ChangeAddress.Enabled = False
ElseIf InStr(ListView2.ListItems(CurrentListIndex2).SubItems(3), "Delete") > 0 Then
DeleteLink2.Enabled = False
RestoreLink.Enabled = True
ChangeAddress.Enabled = False
End If

Me.PopupMenu Menu2
End If
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
ListView2.Visible = True
ListView1.Visible = False
Else
ListView1.Visible = True
ListView2.Visible = False
End If
End Sub

Private Sub Option2_Click()
If Option1.Value = True Then
ListView2.Visible = True
ListView1.Visible = False
Else
ListView1.Visible = True
ListView2.Visible = False
End If
End Sub

Private Sub RestoreLink_Click()
Dim SourceLink As String
Dim DestLink As String
Dim AddPath As String
On Error GoTo err:
AddPath = ListView2.ListItems(CurrentListIndex2).Tag
    If AddPath = "-" Then
    SourceLink = BackupDir & "\Favourite" & "\" & ListView2.ListItems(CurrentListIndex2).Text & ".url"
    DestLink = FavouriteDir & "\" & ListView2.ListItems(CurrentListIndex2).Text & ".url"
    Else
    SourceLink = BackupDir & "\Favourite" & "\" & AddPath & "\" & ListView2.ListItems(CurrentListIndex2).Text & ".url"
    DestLink = FavouriteDir & "\" & AddPath & "\" & ListView2.ListItems(CurrentListIndex2).Text & ".url"
    End If

FileCopy SourceLink, DestLink
ListView2.ListItems(CurrentListIndex2).SubItems(3) = "Restored"
ListView2.ListItems(CurrentListIndex2).ForeColor = vbBlue
ListView2.ListItems(CurrentListIndex2).ListSubItems(1).ForeColor = vbBlue
ListView2.ListItems(CurrentListIndex2).ListSubItems(2).ForeColor = vbBlue
ListView2.ListItems(CurrentListIndex2).ListSubItems(3).ForeColor = vbBlue
Exit Sub
err:
MsgBox err.Description, vbOKOnly, "Error!"
End Sub

Private Sub Timer2_Timer()
Dim x As Integer
TimeOutCounter = TimeOutCounter + 1
ProgressBar2.Value = ProgressBar2.Value + 1
If IndexCount = ListView2.ListItems.Count + 1 Then
Timer2.Enabled = False
Winsock1.Close
ProgressBar1.Visible = False
ProgressBar2.Visible = False
Label1.Visible = False
ProgressBar2.Value = 0
Exit Sub
End If
If TimeOutCounter >= (TimeOutCtr * 10) Then
TimeOut = True
TimeOutCounter = 0
Timer2.Enabled = False
ProgressBar2.Value = 0
Winsock1.Close
HeaderCode = 408
CheckResponse2 True
    IndexCount = IndexCount + 1
    For x = IndexCount To ListView2.ListItems.Count
        If ListView2.ListItems(IndexCount).SubItems(3) = "Change" Then
        IndexCount = IndexCount + 1
        GoTo ReStart:
        End If
    Next x
    ProgressBar1.Visible = False
    Started = True
    Timer2.Enabled = False
    TimeOutCounter = 0
    Command2.Enabled = False
    Command2.Caption = "Stop"
    ProgressBar2.Visible = False
    Label1.Visible = False
    ProgressBar2.Value = 0
    Exit Sub
ReStart:
    VerifyLinks2
End If
End Sub

Private Sub Winsock1_Connect()
ConnectToServer2
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim x As Integer
    CheckResponse2
    Label4.Caption = ""
    ProgressBar1.Value = ProgressBar1.Value + 1
hell:
    If IndexCount = ListView2.ListItems.Count Then
    ProgressBar1.Visible = False
    Started = True
    Timer2.Enabled = False
    TimeOutCounter = 0
    Command2.Enabled = False
    Command2.Caption = "Stop"
    ProgressBar2.Visible = False
    Label1.Visible = False
    ProgressBar2.Value = 0
    Command3.Enabled = True
    Exit Sub
    End If
    IndexCount = IndexCount + 1
    For x = IndexCount To ListView2.ListItems.Count
        If ListView2.ListItems(x).SubItems(3) = "Change" Then
        IndexCount = x + 1
        GoTo ReStart:
        End If
    Next x
    IndexCount = ListView2.ListItems.Count
    GoTo hell:
    
ReStart:
    TimeOutCounter = 0
    ProgressBar2.Value = 0
    VerifyLinks2
End Sub

Private Sub wscHttp_Connect()
ConnectToServer
End Sub

Private Sub ConnectToServer()
url = ListView1.ListItems(IndexCount).SubItems(1)
    Dim strHttpRequest As String
    '
    'create the HTTP Request
    '
    'build request line that contains the HTTP method, 
    'path to the file to retrieve,
    'and HTTP version info. Each line of the request 
    'must be completed by the vbCrLf
    strHttpRequest = "GET " & url & " HTTP/1.1" & vbCrLf
    '
    'add HTTP headers to the request
    '
    'add required header - "Host", that contains the remote host name
    '
    strHttpRequest = strHttpRequest & "Host: " & url & vbCrLf
    '
    'add the "Connection" header to force the server to close the connection
    '
    strHttpRequest = strHttpRequest & "Connection: close" & vbCrLf
    '
    'add optional header "Accept"
    '
    strHttpRequest = strHttpRequest & "Accept: */*" & vbCrLf
    '
    'add other optional headers
    
    strHttpRequest = strHttpRequest & "Proxy-Authorization: Basic " & base64_encode(Username & ":" & Password) & vbCrLf
    strHttpRequest = strHttpRequest & vbCrLf
    '
    'send the request
    wscHttp.SendData strHttpRequest

End Sub

Private Sub ConnectToServer2()
url = ListView2.ListItems(IndexCount).SubItems(1)
    Dim strHttpRequest As String
    '
    'create the HTTP Request
    '
    'build request line that contains the HTTP method, 
    'path to the file to retrieve,
    'and HTTP version info. Each line of the request 
    'must be completed by the vbCrLf
    strHttpRequest = "GET " & url & " HTTP/1.1" & vbCrLf
    '
    'add HTTP headers to the request
    '
    'add required header - "Host", that contains the remote host name
    '
    strHttpRequest = strHttpRequest & "Host: " & url & vbCrLf
    '
    'add the "Connection" header to force the server to close the connection
    '
    strHttpRequest = strHttpRequest & "Connection: close" & vbCrLf
    '
    'add optional header "Accept"
    '
    strHttpRequest = strHttpRequest & "Accept: */*" & vbCrLf
    '
    'add other optional headers
    
    strHttpRequest = strHttpRequest & "Proxy-Authorization: Basic " & base64_encode(Username & ":" & Password) & vbCrLf
    strHttpRequest = strHttpRequest & vbCrLf
    '
    'send the request
    Winsock1.SendData strHttpRequest

End Sub

Private Sub wscHttp_DataArrival(ByVal bytesTotal As Long)
    CheckResponse
    Label4.Caption = ""
    If JustOneItem = False Then ProgressBar1.Value = ProgressBar1.Value + 1
    If IndexCount = ListView1.ListItems.Count Or JustOneItem = True Then
    ProgressBar1.Visible = False
    Started = True
    Timer1.Enabled = False
    TimeOutCounter = 0
    Command2.Enabled = False
    Command2.Caption = "Stop"
    ProgressBar2.Visible = False
    Label1.Visible = False
    ProgressBar2.Value = 0
    VerifyNewAddress
    Exit Sub
    End If
    IndexCount = IndexCount + 1
    TimeOutCounter = 0
    ProgressBar2.Value = 0
    VerifyLinks
End Sub

Sub CheckResponse(Optional SkipGetData As Boolean = False)
    Dim header As String
    Dim temp() As String, temp2() As String
    Dim x As Integer
    On Error Resume Next
    '
    Dim strData As String
    If InStr(strData, "HTTP") <> 0 Then Exit Sub
    '
    'get arrived data from winsock buffer
    '
    If SkipGetData = True Then GoTo skip:
    wscHttp.GetData strData, vbString
    temp() = Split(strData, Chr(Asc(vbNewLine)))
    temp2 = Split(temp(0), " ")
    header = Trim(temp2(1))
    HeaderCode = CInt(header)
    
Dim TempURL As String

    If HeaderCode = 302 Or HeaderCode = 301 Then
        For x = 1 To UBound(temp)
            If InStr(temp(x), "Location") = 2 Then
                AltURL = Trim(Mid(temp(x), InStr(temp(x), " ") + 1))
                If InStr(AltURL, "http://") = 0 Then
                
                    TempURL = Mid(url, InStr(url, "//") + 2)        'Get the www.aaa.com\xxxxxx\xxxxx\xxxx
                    'MsgBox TempURL
                    TempURL = Mid(TempURL, 1, InStr(TempURL, "/") - 1)  'Get the www.aaa.com
                    'MsgBox TempURL
                    TempURL = Mid(url, 1, InStr(url, "//") + 2) & TempURL 'Get the httpx\\www.aaa.com

                    If Left$(AltURL, 1) = "/" Then
                    AltURL = TempURL & AltURL
                    Else
                    AltURL = TempURL & "/" & AltURL
                    End If
                    'MsgBox AltURL
                End If
            Exit For
            End If
        Next x
    End If
skip:
    HeaderDesc = GetDesc(HeaderCode)
    ListView1.ListItems(IndexCount).SubItems(2) = HeaderCode & " (" & HeaderDesc & ")"
    ListView1.ListItems(IndexCount).SubItems(3) = GetAction(HeaderCode)
    
    If HeaderCode <> 200 Then
    ListView1.ListItems(IndexCount).ForeColor = vbRed
    ListView1.ListItems(IndexCount).ListSubItems(1).Bold = True
    ListView1.ListItems(IndexCount).ListSubItems(1).ForeColor = vbRed
    ListView1.ListItems(IndexCount).ListSubItems(2).Bold = True
    ListView1.ListItems(IndexCount).ListSubItems(2).ForeColor = vbRed
'    MsgBox HeaderCode
'    MsgBox GetAction(HeaderCode)
    
    If GetAction(HeaderCode) <> "No Action" Then
    ListView2.ListItems.Add , , ListView1.ListItems(IndexCount).Text
    ListView2.ListItems(ListView2.ListItems.Count).Tag = ListView1.ListItems(IndexCount).Tag
    ListView2.ListItems(ListView2.ListItems.Count).SubItems(1) = ListView1.ListItems(IndexCount).SubItems(1)
    ListView2.ListItems(ListView2.ListItems.Count).SubItems(2) = HeaderCode & " (" & HeaderDesc & ")"
    ListView2.ListItems(ListView2.ListItems.Count).ForeColor = vbRed
    ListView2.ListItems(ListView2.ListItems.Count).ListSubItems(1).Bold = True
    ListView2.ListItems(ListView2.ListItems.Count).ListSubItems(1).ForeColor = vbRed
    ListView2.ListItems(ListView2.ListItems.Count).ListSubItems(2).Bold = True
    ListView2.ListItems(ListView2.ListItems.Count).ListSubItems(2).ForeColor = vbRed
    ListView2.ListItems(ListView2.ListItems.Count).SubItems(3) = GetAction(HeaderCode)
    
    If GetAction(HeaderCode) = "Deleted" Then ExecuteAction IndexCount, GetAction(HeaderCode), AltURL
    
        If HeaderCode = 302 Or HeaderCode = 301 Then
            If Trim(AltURL) <> ListView1.ListItems(IndexCount).ListSubItems(1) Then
            ListView2.ListItems.Add , , "New Address:"
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(1) = AltURL
            ListView2.Refresh
            Else
            HeaderCode = 200
            ListView2.ListItems(ListView2.ListItems.Count).SubItems(2) = 202 & " (" & "Link Active" & ")"
            ListView2.ListItems(ListView2.ListItems.Count).ForeColor = vbBlue
            ListView2.ListItems(ListView2.ListItems.Count).ListSubItems(1).Bold = True
            ListView2.ListItems(ListView2.ListItems.Count).ListSubItems(1).ForeColor = vbBlue
            ListView2.ListItems(ListView2.ListItems.Count).ListSubItems(2).Bold = True
            ListView2.ListItems(ListView2.ListItems.Count).ListSubItems(2).ForeColor = vbBlue
            ListView1.ListItems(IndexCount).SubItems(2) = HeaderCode & " (" & HeaderDesc & ")"
            ListView2.Refresh
            GoTo AssumeOK
            End If
        End If
    End If
    Else
AssumeOK:
    ListView1.ListItems(IndexCount).ForeColor = vbBlue
    ListView1.ListItems(IndexCount).ListSubItems(1).Bold = True
    ListView1.ListItems(IndexCount).ListSubItems(1).ForeColor = vbBlue
    ListView1.ListItems(IndexCount).ListSubItems(2).Bold = True
    ListView1.ListItems(IndexCount).ListSubItems(2).ForeColor = vbBlue
    End If
wscHttp.Close
End Sub

Sub CheckResponse2(Optional SkipGetData As Boolean = False)
    Dim header As String
    Dim temp() As String, temp2() As String
    Dim x As Integer
    On Error Resume Next
    
    Dim strData As String
    If InStr(strData, "HTTP") <> 0 Then Exit Sub
    '
    'get arrived data from winsock buffer
    '
    If SkipGetData = True Then GoTo skip:
    Winsock1.GetData strData, vbString
    temp() = Split(strData, Chr(Asc(vbNewLine)))
    temp2 = Split(temp(0), " ")
    header = Trim(temp2(1))
    HeaderCode = CInt(header)
    ListView2.ListItems(IndexCount).SubItems(2) = GetDesc(HeaderCode)
    If HeaderCode = 200 Then ExecuteAction IndexCount, "Change", ListView2.ListItems(IndexCount).SubItems(1)
    Winsock1.Close
    Exit Sub
skip:
ListView2.ListItems(IndexCount - 1).SubItems(2) = "InValid"
Winsock1.Close
End Sub


Sub ExecuteAction(Index As Integer, ActionCode As String, Optional NewURL As String = "")
Dim AddPath As String
If ActionCode = "Deleted" Then
AddPath = ListView1.ListItems(Index).Tag
Else
AddPath = ListView2.ListItems(Index - 1).Tag
End If
    If AddPath = "-" Then
    AddPath = FavouriteDir
    Else
    AddPath = FavouriteDir & "\" & AddPath
    End If
Select Case ActionCode:
Case "Deleted":
CommonFunction.DeleteLink AddPath, ListView1.ListItems(Index).Text
Case "Change":
CommonFunction.UpdateLink AddPath, ListView2.ListItems(Index - 1).Text, NewURL
ListView2.ListItems(Index - 1).SubItems(3) = "Changed"
End Select
End Sub

Sub ListFavorite()
Dim Result As Long
Dim TotalFile As Long
Dim totalLine As Long
Dim fileBuffer() As String
Dim Bookmarks() As String
Dim tmpindex() As Long
Dim i As Long
Dim j As Long
Dim lLen As Long
Dim url As String
Dim Desc As String
Dim Action As Long
Dim RealName As String
Dim AddPath As String

ListView1.ListItems.Clear
    TotalFile = vbFileIO.GetAllFileName(FavouriteDir, "*.url", Bookmarks, RETURN_FILE, True)
    ReDim ListLine(TotalFile - 1)
    ReDim ListIndex(TotalFile - 1)
    ReDim ListBM(TotalFile - 1)
    ListCtr = 0
    If TotalFile > 0 Then
      For i = 0 To TotalFile - 1
      
        totalLine = vbFileIO.ReadFileText(Bookmarks(i), fileBuffer)
        If totalLine > 0 Then
        
        For j = 0 To totalLine - 1
          lLen = Len(fileBuffer(j))
          If lLen > 4 Then
            If Mid(fileBuffer(j), 1, 4) = "URL=" Then
              url = Right(fileBuffer(j), lLen - 4)
              RealName = Bookmarks(i)
              AddPath = Mid(Bookmarks(i), Len(FavouriteDir) + 2)
              If InStr(AddPath, "\") > 0 Then
              AddPath = Mid(AddPath, 1, InStr(AddPath, "\") - 1)
              AddPath = Trim(AddPath)
              Else
              AddPath = "-"
              End If
             RealName = StrReverse(Bookmarks(i))
             RealName = Mid(Bookmarks(i), Len(Bookmarks(i)) - InStr(RealName, "\") + 2)
             RealName = Mid(RealName, 1, Len(RealName) - 4)
              ListView1.ListItems.Add , , RealName
              ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = url
              ListView1.ListItems(ListView1.ListItems.Count).Tag = AddPath
              DoEvents
            End If
            End If
        Next j
        
        End If
        
      Next i
      End If
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Stop" Then
If JustOneItem = False Then
LastStop = IndexCount
Else
LastStop = CurrentIndex
End If
IndexCount = ListView1.ListItems.Count + 1
Command2.Caption = "Resume"
Else
Command2.Caption = "Stop"
JustOneItem = False
ProgressBar1.Visible = True
IndexCount = LastStop
VerifyLinks
End If
End Sub

Private Sub Command3_Click()
BackupFav
If Started = True Then ListFavorite
Started = True
JustOneItem = False
ListView2.ListItems.Clear
ProgressBar1.Visible = True
ProgressBar1.Max = ListView1.ListItems.Count
ProgressBar1.Value = 0
ProgressBar2.Visible = True
Label1.Visible = True
ProgressBar2.Max = TimeOutCtr * 10
ProgressBar2.Value = 0
IndexCount = 1
Command2.Enabled = True
Command2.Caption = "Stop"
VerifyLinks
Command3.Enabled = False
Command2.Enabled = True
End Sub

Sub VerifyNewAddress()
Dim NewAdd As Integer, StartNewAdd As Integer, x As Integer
StartNewAdd = 0
For x = 1 To ListView2.ListItems.Count
If ListView2.ListItems(x).SubItems(3) = "Change" Then
If StartNewAdd = 0 Then StartNewAdd = x
NewAdd = NewAdd + 1
End If
Next x
If NewAdd = 0 Then
Command3.Enabled = True
Exit Sub
End If
ProgressBar1.Visible = True
ProgressBar1.Max = NewAdd
ProgressBar1.Value = 0
ProgressBar2.Visible = True
Label1.Visible = True
ProgressBar2.Max = TimeOutCtr * 10
ProgressBar2.Value = 0
IndexCount = StartNewAdd + 1
Command2.Enabled = True
Command2.Caption = "Stop"
VerifyLinks2
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Command6_Click()
frmSetting.Show
frmSetting.Text3 = FavouriteDir
frmSetting.Text4 = BackupDir
End Sub

Private Sub DeleteLink_Click()
Dim AddPath As String
If MsgBox("Are you sure to delete '" & ListView1.ListItems(CurrentIndex).Text & "' from the FAVOURITE list?", vbYesNo, "Delete Link") = vbNo Then Exit Sub
AddPath = ListView1.ListItems(CurrentIndex).Tag
    If AddPath = "-" Then
    AddPath = FavouriteDir
    Else
    AddPath = FavouriteDir & "\" & AddPath
    End If
CommonFunction.DeleteLink AddPath, ListView1.ListItems(CurrentIndex).Text
ListView1.ListItems.Remove CurrentIndex
End Sub

Private Sub Form_Load()
TimeOutCounter = 0
Dim temp As String
Open App.Path & "\info.ini" For Input As #1
Input #1, temp
IfRedirected = CInt(temp)
Input #1, temp
IfTimeOut = CInt(temp)
Input #1, temp
IfUnathorized = CInt(temp)
Input #1, temp
IfNotFound = CInt(temp)
Input #1, temp
IsForbidden = CInt(temp)
Input #1, temp
IfBadRequest = CInt(temp)
Input #1, temp
IfOther = CInt(temp)
Input #1, temp
TimeOutCtr = CInt(temp)
Input #1, temp
HideDetails = CInt(temp)
Input #1, temp
UseProxy = CInt(temp)
Input #1, temp
UserLogin = base64_decode(temp)
Input #1, temp
UserPwd = base64_decode(temp)
Input #1, temp
Proxy = temp
Input #1, temp
Password = temp
Close #1

Dim x
Dim tmp As Long

  FavouriteDir = GetSetting("BookmarkCleaner", "Directory", "Favourites")
  If Len(FavouriteDir) = 0 Then
    vbDialog.BrowseForFolder Me.hwnd, FavouriteDir, "Please select Favourites folder:"
  End If
  
  BackupDir = GetSetting("BookmarkCleaner", "Directory", "Backup")
  If Len(BackupDir) = 0 Then
    vbDialog.BrowseForFolder Me.hwnd, BackupDir, "Please select Backup folder:"
  End If
'  Text1(0).Text = FavouriteDir
'  Text1(1).Text = BackupDir
  
If HideDetails = -1 Then
ListView1.Visible = False
ListView2.Visible = True
Option1.Value = True
Else
ListView2.Visible = False
ListView1.Visible = True
Option2.Value = True
End If
ListFavorite

Proxy = "proxy"
Port = "8080"
Username = UserLogin
Password = UserPwd
url = "http://microsoft.com/downloads/details.aspx?FamilyId=9996B314-0364-4623-9EDE-0B5FBB133652&displaylang=en"
Load frmSetting
Action(0) = "No Action"
Action(1) = "Deleted"
Action(2) = "Prompt"
Action(3) = "Change"
Action(4) = "Error"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim x
  
  SaveSetting "BookmarkCleaner", "Directory", "Favourites", FavouriteDir
  SaveSetting "BookmarkCleaner", "Directory", "Backup", BackupDir
  
  For Each x In Me
    If x.Name = "Combo1" Then
      If x.ListIndex = -1 Then
        x.ListIndex = 0
      End If
      SaveSetting "BookmarkCleaner", "Action", CStr(x.Index), CStr(x.ListIndex)
    End If
  Next
  
End Sub

Private Sub ListView1_DblClick()
Dim x As Integer
For x = 1 To ListView1.ListItems.Count
If ListView1.ListItems(x).Selected = True Then
OpenURL ListView1.ListItems(x).SubItems(1)
Exit For
End If
Next x
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
CurrentIndex = Item.Index
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Selected As Boolean
If Button = 2 Then
For x = 1 To ListView1.ListItems.Count
If ListView1.ListItems(x).Selected = True Then
    If ListView1.ListItems(x).SubItems(3) = "Deleted" Then
    ModifyLink.Enabled = False
    DeleteLink.Enabled = False
    VerifyLink.Enabled = False
    OpenLink.Enabled = False
    Else
    ModifyLink.Enabled = True
    DeleteLink.Enabled = True
    VerifyLink.Enabled = True
    OpenLink.Enabled = True
    End If
Selected = True
Exit For
End If
Next x
If Selected = True Then Me.PopupMenu Menu
End If
End Sub

Private Sub ModifyLink_Click()
frmEditLink.Show
frmEditLink.ListIndex = CurrentIndex
frmEditLink.OldLinkName = ListView1.ListItems(CurrentIndex).Text
frmEditLink.Text1 = ListView1.ListItems(CurrentIndex).Text
frmEditLink.Text2 = ListView1.ListItems(CurrentIndex).SubItems(1)
End Sub

Private Sub OpenLink_Click()
OpenURL ListView1.ListItems(CurrentIndex).SubItems(1)
End Sub

Private Sub Timer1_Timer()
TimeOutCounter = TimeOutCounter + 1
ProgressBar2.Value = ProgressBar2.Value + 1
If IndexCount = ListView1.ListItems.Count + 1 Then
Timer1.Enabled = False
wscHttp.Close
ProgressBar1.Visible = False
ProgressBar2.Visible = False
Label1.Visible = False
ProgressBar2.Value = 0
Exit Sub
End If
If TimeOutCounter >= (TimeOutCtr * 10) Then
TimeOut = True
TimeOutCounter = 0
Timer1.Enabled = False
ProgressBar2.Value = 0
wscHttp.Close
HeaderCode = 408
CheckResponse True
    If JustOneItem = False Then
    IndexCount = IndexCount + 1
    VerifyLinks
    End If
End If
End Sub

Private Sub VerifyLink_Click()
Command2.Enabled = True
Command2.Caption = "Stop"
JustOneItem = True
IndexCount = CurrentIndex
ListView2.ListItems.Clear
VerifyLinks
url = ListView1.ListItems(CurrentIndex).SubItems(1)
ProgressBar2.Visible = True
Label1.Visible = True
ProgressBar2.Value = 0
VerifyLinks
End Sub

Function GetAction(ResponseID As Integer) As String
Select Case ResponseID:
Case 400:
GetAction = Action(IfBadRequest)
Case 501:
GetAction = Action(IfBadRequest)
Case 403:
GetAction = Action(IsForbidden)
Case 404:
GetAction = Action(IfNotFound)
Case 504:
GetAction = Action(IfTimeOut)
Case 503:
GetAction = Action(IfNotFound)
Case 502:
GetAction = Action(IfTimeOut)
Case 408:
GetAction = Action(IfTimeOut)
Case 407:
GetAction = Action(4)
Case 302:
GetAction = Action(IfRedirected)
Case 301:
GetAction = Action(IfRedirected)
Case Else
GetAction = Action(IfOther)
End Select
End Function
