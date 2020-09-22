VERSION 5.00
Begin VB.Form frmSetting 
   BackColor       =   &H00E7E3DE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setting"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FBF3EE&
      Caption         =   "&Proxy Setting"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Proxy Server's Configuration"
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FBF3EE&
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Cancel Setting"
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FBF3EE&
      Caption         =   "&Save"
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Save Setting"
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E7E3DE&
      Caption         =   "Cleaning Display Option"
      Height          =   735
      Left            =   120
      TabIndex        =   21
      Top             =   4320
      Width           =   7215
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E7E3DE&
         Caption         =   "Show all links"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   23
         Top             =   360
         Width           =   3135
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E7E3DE&
         Caption         =   "Show erronous links only"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Value           =   -1  'True
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E7E3DE&
      Caption         =   "Link Cleaning Action"
      Height          =   2055
      Left            =   0
      TabIndex        =   10
      Top             =   2160
      Width           =   7215
      Begin VB.TextBox txtTimeOut 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   6000
         TabIndex        =   31
         Text            =   "30"
         Top             =   1560
         Width           =   975
      End
      Begin VB.ComboBox Combo7 
         Height          =   360
         ItemData        =   "frmSetting.frx":0ECA
         Left            =   1560
         List            =   "frmSetting.frx":0ED7
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1440
         Width           =   1815
      End
      Begin VB.ComboBox Combo6 
         Height          =   360
         ItemData        =   "frmSetting.frx":0EF9
         Left            =   5160
         List            =   "frmSetting.frx":0F06
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox Combo5 
         Height          =   360
         ItemData        =   "frmSetting.frx":0F28
         Left            =   5160
         List            =   "frmSetting.frx":0F35
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox Combo4 
         Height          =   360
         ItemData        =   "frmSetting.frx":0F57
         Left            =   5160
         List            =   "frmSetting.frx":0F64
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox Combo3 
         Height          =   360
         ItemData        =   "frmSetting.frx":0F86
         Left            =   1560
         List            =   "frmSetting.frx":0F93
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox Combo2 
         Height          =   360
         ItemData        =   "frmSetting.frx":0FB5
         Left            =   1560
         List            =   "frmSetting.frx":0FC2
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frmSetting.frx":0FE4
         Left            =   1560
         List            =   "frmSetting.frx":0FF4
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Time Out Period (Sec):"
         Height          =   255
         Index           =   11
         Left            =   3720
         TabIndex        =   30
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Other:"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bad Request:"
         Height          =   255
         Index           =   9
         Left            =   3720
         TabIndex        =   27
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Unauthorized:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Forbidden:"
         Height          =   255
         Index           =   7
         Left            =   3720
         TabIndex        =   14
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Not Found:"
         Height          =   255
         Index           =   6
         Left            =   3720
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "New Address:"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Timout:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1320
      Width           =   7215
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   240
      Width           =   7215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E7E3DE&
      Caption         =   "I use proxy server to connect to the Internet"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Check here if you use proxy server to connect to the Internet"
      Top             =   5160
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E7E3DE&
      Caption         =   "Proxy Server Information"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   840
         Width           =   5775
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Backup folder path:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   3
      Left            =   2640
      MouseIcon       =   "frmSetting.frx":1024
      MousePointer    =   99  'Custom
      TabIndex        =   8
      ToolTipText     =   "Browse for backup folder path"
      Top             =   1080
      Width           =   1875
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Favourite folder path:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   2
      Left            =   2520
      MouseIcon       =   "frmSetting.frx":132E
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "Browse for favourite folder path"
      Top             =   0
      Width           =   2085
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
Me.Height = 7890
Command1.Top = 6960
Command2.Top = Command1.Top
Frame1.Visible = True
Command3.Enabled = True
Else
Command1.Top = 5520
Command2.Top = Command1.Top
Me.Height = 6465
Frame1.Visible = False
Command3.Enabled = False
End If
End Sub

Private Sub Command1_Click()
SaveSettings
  SaveSetting "BookmarkCleaner", "Directory", "Favourites", FavouriteDir
  SaveSetting "BookmarkCleaner", "Directory", "Backup", BackupDir
End Sub

Sub SaveSettings()
Open App.Path & "\info.ini" For Output As #1
Write #1, Combo1.ListIndex
Write #1, Combo2.ListIndex
Write #1, Combo3.ListIndex
Write #1, Combo4.ListIndex
Write #1, Combo5.ListIndex
Write #1, Combo6.ListIndex
Write #1, Combo7.ListIndex
Write #1, txtTimeOut.Text
Write #1, CInt(Option1(0).Value)
Write #1, Check1.Value
Write #1, base64_encode(Text1.Text)
Write #1, base64_encode(Text2.Text)
Write #1, Form1.Proxy
Write #1, Form1.Port
Close #1

IfRedirected = Combo1.ListIndex
IfTimeOut = Combo2.ListIndex
IfUnathorized = Combo3.ListIndex
IfNotFound = Combo4.ListIndex
IsForbidden = Combo5.ListIndex
IfBadRequest = Combo6.ListIndex
IfOther = Combo7.ListIndex
TimeOutCtr = CInt(txtTimeOut.Text)
HideDetails = Option1(0).Value
UseProxy = Check1.Value
UserLogin = Text1
UserPwd = Text2
Form1.Username = UserLogin
Form1.Password = UserPwd

With Form1
If HideDetails = -1 Then
.Option2 = False
.Option1 = True
Else
.Option1 = False
.Option2 = True
End If
End With

Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
frmProxy.Show
End Sub

Private Sub Form_Load()
'Exit Sub
Combo1.ListIndex = IfRedirected
Combo2.ListIndex = IfTimeOut
Combo3.ListIndex = IfUnathorized
Combo4.ListIndex = IfNotFound
Combo5.ListIndex = IsForbidden
Combo6.ListIndex = IfBadRequest
Combo7.ListIndex = IfOther
txtTimeOut.Text = TimeOutCtr
Option1(0).Value = CBool(HideDetails)
If Option1(0).Value = False Then Option1(1).Value = True
Check1.Value = UseProxy
Text1.Text = UserLogin
Text2.Text = UserPwd
End Sub

Private Sub Label1_Click(Index As Integer)
If Index = 2 Then
    vbDialog.BrowseForFolder Me.hwnd, FavouriteDir, "Select Favourite Folder:"
    Text3.Text = FavouriteDir
ElseIf Index = 3 Then
    vbDialog.BrowseForFolder Me.hwnd, BackupDir, "Select Backup Folder:"
    Text4.Text = BackupDir
End If
End Sub

