VERSION 5.00
Begin VB.Form Form0 
   BackColor       =   &H00800000&
   Caption         =   "Malacanang  Library"
   ClientHeight    =   2580
   ClientLeft      =   7350
   ClientTop       =   4905
   ClientWidth     =   4680
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PATRON"
      DownPicture     =   "Form0.frx":0000
      Height          =   1095
      Left            =   2400
      Picture         =   "Form0.frx":0410
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADMINISTRATOR"
      DownPicture     =   "Form0.frx":051F
      Height          =   1095
      Left            =   120
      Picture         =   "Form0.frx":08F4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Welcome to Malacanang Library. Please select an account type."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "Form0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''THIS IS THE PRE MAIN LOGIN FORM'''''''''''''''''''''''''''''

Private Sub Command1_Click()    'Admin Button
With MLSDB.rsAdministrator
If MLSDB.rsAdministrator.RecordCount = 0 Then
    AdminReg.Show
    Unload Me
Else
    Alog.Show
    Unload Me
End If
End With
End Sub

Private Sub Command2_Click()    'Patron Button
With MLSDB.rsPatronDB
If MLSDB.rsPatronDB.RecordCount = 0 Then
    MsgBox "Please see an Administrator to register", vbOKOnly
Else
    Plog.Show
    Unload Me
End If
End With
End Sub

Private Sub Exit_Click()
'Credits.Show
Unload Me
End Sub

Private Sub Form_Load()
MLSDB.rsCirculation.Open
MLSDB.rsBookDbase.Open
MLSDB.rsPatronDB.Open
MLSDB.rsAdministrator.Open
End Sub
