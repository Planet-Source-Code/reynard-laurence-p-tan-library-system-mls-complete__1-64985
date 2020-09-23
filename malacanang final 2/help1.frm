VERSION 5.00
Begin VB.Form help1 
   BackColor       =   &H00800000&
   Caption         =   "Help Index"
   ClientHeight    =   9690
   ClientLeft      =   5460
   ClientTop       =   630
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   8580
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command10 
      Caption         =   "RETURN"
      Height          =   615
      Left            =   1560
      TabIndex        =   8
      Top             =   8640
      Width           =   2175
   End
   Begin VB.CommandButton Command9 
      Caption         =   "ADMINISTRATOR PROFILE"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   7200
      Width           =   4335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "BARCODE SEARCH"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   6480
      Width           =   4335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "ADMINISTRATOR SEARCH"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   5760
      Width           =   4335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "BOOK INFORMATION"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   5040
      Width           =   4335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "PATRON PROFILE"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   4320
      Width           =   4335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ADMINISTRATOR MAIN MENU"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   3600
      Width           =   4335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ADMINISTRATOR LOGIN"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2880
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GETTING STARTED"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   2160
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      Height          =   615
      Left            =   4920
      TabIndex        =   9
      Top             =   8640
      Width           =   2175
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
      Height          =   6735
      Left            =   1560
      TabIndex        =   10
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
      TabIndex        =   11
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "help1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command10_Click()
    helpindex1.Show
    Unload Me
End Sub

Private Sub Command2_Click()
        helpstartup.Show
        Unload Me
End Sub

Private Sub Command3_Click()
        helpadminlog.Show
        Unload Me
End Sub

Private Sub Command4_Click()
        helpadminmain.Show
        Unload Me
End Sub

Private Sub Command5_Click()
        helpclientprofile.Show
        Unload Me
End Sub

Private Sub Command6_Click()
        helpbookinfo.Show
        Unload Me
End Sub

Private Sub Command7_Click()
        helpadminsearch.Show
        Unload Me
End Sub

Private Sub Command8_Click()
        helpbcode.Show
        Unload Me
End Sub

Private Sub Command9_Click()
        helpadminprofile.Show
        Unload Me
End Sub

