VERSION 5.00
Begin VB.Form help2 
   BackColor       =   &H00800000&
   Caption         =   "Help Index"
   ClientHeight    =   7485
   ClientLeft      =   5460
   ClientTop       =   630
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   8580
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command10 
      Caption         =   "RETURN"
      Height          =   615
      Left            =   1560
      TabIndex        =   5
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      Height          =   615
      Left            =   4920
      TabIndex        =   6
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GETTING STARTED"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   2400
      Width           =   4335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PATRON LOGIN"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   3120
      Width           =   4335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PATRON MAIN MENU"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   3840
      Width           =   4335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "PATRON PROFILE"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   4560
      Width           =   4335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PATRON SEARCH"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   5280
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "HELP INDEX"
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
      Height          =   4695
      Left            =   1560
      TabIndex        =   7
      Top             =   1560
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "WELCOME TO THE ONLINE HELP INDEX MENU. HOW MAY I HELP YOU?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1335
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "help2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command10_Click()
helpindex2.Show
Unload Me
End Sub

Private Sub Command2_Click()
helpclientstart.Show
Unload Me
End Sub

Private Sub Command3_Click()
helpclientlog.Show
Unload Me
End Sub

Private Sub Command4_Click()
helpclientmain.Show
Unload Me
End Sub

Private Sub Command5_Click()
helpclientprofile2.Show
Unload Me
End Sub

Private Sub Command6_Click()
helpclientsearch.Show
Unload Me
End Sub

