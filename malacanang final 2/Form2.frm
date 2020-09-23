VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00800000&
   Caption         =   "Administrator"
   ClientHeight    =   4035
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   5895
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      DownPicture     =   "Form2.frx":0000
      Height          =   735
      Left            =   4800
      Picture         =   "Form2.frx":09B5
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BOOK INFORMATION"
      DownPicture     =   "Form2.frx":1370
      Height          =   975
      Index           =   0
      Left            =   2160
      Picture         =   "Form2.frx":17EC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "LOGOUT"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ADMINISTRATOR PROFILE"
      DownPicture     =   "Form2.frx":1C5C
      Height          =   975
      Left            =   3240
      Picture         =   "Form2.frx":20E5
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "BARCODE SEARCH"
      DisabledPicture =   "Form2.frx":22A1
      DownPicture     =   "Form2.frx":26E4
      Height          =   975
      Left            =   1200
      Picture         =   "Form2.frx":2B27
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ADMINISTRATOR SEARCH"
      DownPicture     =   "Form2.frx":2C47
      Height          =   975
      Left            =   4080
      Picture         =   "Form2.frx":313C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PATRON PROFILE"
      DownPicture     =   "Form2.frx":33BE
      Height          =   975
      Left            =   240
      Picture         =   "Form2.frx":37CE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1575
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
      Height          =   2775
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5655
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00800000&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   480
      TabIndex        =   8
      Top             =   3000
      Width           =   3135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''THIS IS THE ADMINISTRATOR MAIN FORM'''''''''''''''''''''''''''

Private Sub Command1_Click()    'Patron Button
Form4.Show
Unload Me
End Sub

Private Sub Command2_Click(Index As Integer)    'Book Info Button
Form5.Show
Unload Me
End Sub

Private Sub Command3_Click()    'Simple Search Button
With MLSDB.rsBookDbase
If MLSDB.rsBookDbase.RecordCount = 0 Then
    MsgBox "The Record is empty", vbCritical
    Form2.Refresh
Else
    Form8.Show
    Unload Me
End If
End With
End Sub
Private Sub Command4_Click()    'Barcode Search Button
Form9.Show
Unload Me
End Sub

Private Sub Command5_Click()
    helpindex1.Show
End Sub

Private Sub Command6_Click()    'Profile Button
Form6.Show
Unload Me
End Sub
Private Sub Command7_Click()    'Logout Button
Alog.Show
Unload Me
End Sub
Private Sub Command8_Click()    'Exit Button
MLSDB.rsCirculation.Close
MLSDB.rsBookDbase.Close
MLSDB.rsPatronDB.Close
MLSDB.rsAdministrator.Close
Form0.Show
Unload Me
End Sub
