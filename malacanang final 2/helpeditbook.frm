VERSION 5.00
Begin VB.Form helpeditbook 
   BackColor       =   &H00400000&
   Caption         =   "Help - Edit Book"
   ClientHeight    =   9690
   ClientLeft      =   5460
   ClientTop       =   630
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   8580
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "RETURN"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   9120
      Width           =   1095
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   5880
      Left            =   1440
      Picture         =   "helpeditbook.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   $"helpeditbook.frx":12F9E
      Height          =   2055
      Left            =   960
      TabIndex        =   2
      Top             =   6960
      Width           =   6735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "THE EDIT BOOK INFORMATION WINDOW"
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
      Left            =   -2760
      TabIndex        =   1
      Top             =   0
      Width           =   14175
   End
End
Attribute VB_Name = "helpeditbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
Unload Me
End Sub
