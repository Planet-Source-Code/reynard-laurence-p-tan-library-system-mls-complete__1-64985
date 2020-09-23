VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00800000&
   Caption         =   "Patron"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5790
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "EXIT"
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "LOGOUT"
      Height          =   615
      Left            =   360
      Picture         =   "Form3.frx":0000
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SIMPLE SEARCH"
      DownPicture     =   "Form3.frx":00D6
      Height          =   855
      Left            =   3480
      Picture         =   "Form3.frx":05CB
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROFILE"
      DownPicture     =   "Form3.frx":084D
      Height          =   855
      Left            =   480
      Picture         =   "Form3.frx":0CD6
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   5535
      Begin VB.CommandButton Command5 
         DownPicture     =   "Form3.frx":0E92
         Height          =   735
         Left            =   4680
         Picture         =   "Form3.frx":1847
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2160
         Visible         =   0   'False
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''THIS IS THE PATRON MAIN FORM'''''''''''''''''''''''''''''''''

Private Sub Command1_Click()    'Profile Button
pprofile.Show
Unload Me
End Sub
Private Sub Command2_Click()    'Search Button
With MLSDB.rsBookDbase
If MLSDB.rsBookDbase.RecordCount = 0 Then
    MsgBox "The Record is empty", vbCritical
    Form3.Refresh
Else
    PSearch.Show
    Unload Me
End If
End With
End Sub
Private Sub Command3_Click()    'Logout Button
Plog.Show
Unload Me
End Sub
Private Sub Command4_Click()    'Exit Button
MLSDB.rsCirculation.Close
MLSDB.rsBookDbase.Close
MLSDB.rsPatronDB.Close
MLSDB.rsAdministrator.Close
Form0.Show
Unload Me
End Sub

Private Sub Command5_Click()
helpindex2.Show
End Sub
