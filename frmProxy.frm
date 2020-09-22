VERSION 5.00
Begin VB.Form frmProxy 
   BackColor       =   &H00E7E3DE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proxy Setting"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProxy.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E7E3DE&
      Caption         =   "Proxy Setting"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FBF3EE&
         Caption         =   "&Save"
         Height          =   495
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FBF3EE&
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1560
         TabIndex        =   3
         Text            =   "8080"
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1560
         TabIndex        =   1
         Text            =   "proxy"
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Proxy Server:"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Proxy = Text1
Form1.Port = Text2
frmSetting.SaveSettings
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = Form1.Proxy
Text2.Text = Form1.Port
End Sub
