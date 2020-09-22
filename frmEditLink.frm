VERSION 5.00
Begin VB.Form frmEditLink 
   BackColor       =   &H00E7E3DE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Link"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditLink.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   10410
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E7E3DE&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FBF3EE&
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FBF3EE&
         Caption         =   "&Save"
         Height          =   495
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   8775
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   8775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Link Name:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmEditLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OldLinkName As String
Public ListIndex As Integer

Private Sub Command1_Click()
Dim AddPath As String
AddPath = Form1.ListView1.ListItems(ListIndex).Tag
    If AddPath = "-" Then
    AddPath = FavouriteDir
    Else
    AddPath = FavouriteDir & "\" & AddPath
    End If

CommonFunction.DeleteLink AddPath, Form1.ListView1.ListItems(ListIndex).Text
CommonFunction.CreateLink AddPath, Text1, Text2

With Form1
.ListView1.ListItems(ListIndex).Text = Text1
.ListView1.ListItems(ListIndex).SubItems(1) = Text2
End With
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
