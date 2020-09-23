VERSION 5.00
Begin VB.Form helpclientmain 
   BackColor       =   &H00800000&
   Caption         =   "Help - Patron Main Menu"
   ClientHeight    =   9720
   ClientLeft      =   2490
   ClientTop       =   630
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   14385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00800000&
      Caption         =   "RELATED TOPICS"
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
      Height          =   3855
      Left            =   11160
      TabIndex        =   13
      Top             =   1320
      Width           =   3135
   End
   Begin VB.CommandButton Command9 
      Caption         =   "PATRON SEARCH"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CommandButton Command8 
      Caption         =   "PATRON PROFILE"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   3840
      Width           =   2655
   End
   Begin VB.CommandButton Command7 
      Caption         =   "PATRON MAIN MENU"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PATRON LOGIN"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "GETTING STARTED"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "HELP MAIN"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   9120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   7320
      Picture         =   "helpclientmain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   8640
      Picture         =   "helpclientmain.frx":026E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   3840
      Picture         =   "helpclientmain.frx":041D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9120
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Left            =   4080
      Picture         =   "helpclientmain.frx":0675
      ScaleHeight     =   4275
      ScaleWidth      =   6165
      TabIndex        =   9
      Top             =   1560
      Width           =   6225
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
      Height          =   3855
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "THE PATRON MAIN MENU"
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
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   14175
   End
   Begin VB.Label Label1 
      Caption         =   $"helpclientmain.frx":C907
      Height          =   2055
      Left            =   3840
      TabIndex        =   10
      Top             =   6960
      Width           =   6735
   End
End
Attribute VB_Name = "helpclientmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    helpclientlog.Show
    Unload Me
End Sub

Private Sub Command2_Click()
    helpclientprofile2.Show
    Unload Me
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    help2.Show
    Unload Me
End Sub
Private Sub Command5_Click()
helpclientstart.Show
Unload Me
End Sub

Private Sub Command6_Click()
helpclientlog.Show
Unload Me
End Sub

Private Sub Command7_Click()
helpclientmain.Show
Unload Me
End Sub

Private Sub Command8_Click()
helpclientprofile2.Show
Unload Me
End Sub

Private Sub Command9_Click()
helpclientsearch.Show
Unload Me
End Sub




