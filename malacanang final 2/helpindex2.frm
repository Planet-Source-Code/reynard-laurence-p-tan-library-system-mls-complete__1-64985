VERSION 5.00
Begin VB.Form helpindex2 
   BackColor       =   &H00800000&
   Caption         =   "Help"
   ClientHeight    =   3090
   ClientLeft      =   7335
   ClientTop       =   4320
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HELP INDEX"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CREDITS"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Welcome to the Online Help Menu"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "helpindex2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
help2.Show
Unload Me
End Sub

Private Sub Command2_Click()
about.Show
End Sub

Private Sub Command3_Click()
Credits.Show
End Sub


Private Sub Command4_Click()
Unload Me
End Sub
