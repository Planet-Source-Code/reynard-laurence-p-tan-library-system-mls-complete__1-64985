VERSION 5.00
Begin VB.Form password 
   Caption         =   "DELETION PASSWORD"
   ClientHeight    =   1050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   ScaleHeight     =   1050
   ScaleWidth      =   3180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "ENTER PASSWORD:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
With MLSDB.rsPatronDB
    If Text1.Text = "tito" Then
        MLSDB.rsPatronDB.Delete
        .MoveFirst
        MsgBox "Record Deleted Successfully", vbInformation
        Form4.Show
    Else
        MsgBox "Incorrect Password", vbCritical
        Form4.Show
        Unload Me
    End If
    
End With
End Sub

Private Sub Command2_Click()
Form4.Show
Unload Me
End Sub
