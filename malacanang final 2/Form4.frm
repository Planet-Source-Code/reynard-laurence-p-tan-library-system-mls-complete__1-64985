VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00800000&
   Caption         =   "Patron Profile"
   ClientHeight    =   7875
   ClientLeft      =   5250
   ClientTop       =   2775
   ClientWidth     =   8490
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   8490
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      DataMember      =   "PatronDB"
      DataSource      =   "MLSDB"
      Height          =   315
      Left            =   2640
      TabIndex        =   7
      Top             =   1200
      Width           =   3135
   End
   Begin VB.CommandButton Command6 
      DownPicture     =   "Form4.frx":0000
      Height          =   735
      Left            =   7680
      Picture         =   "Form4.frx":09B5
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MAIN MENU"
      DownPicture     =   "Form4.frx":1370
      Height          =   855
      Left            =   6720
      Picture         =   "Form4.frx":1910
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "REFRESH"
      DownPicture     =   "Form4.frx":1EE2
      Height          =   735
      Left            =   120
      Picture         =   "Form4.frx":2516
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      DownPicture     =   "Form4.frx":27AE
      Height          =   855
      Left            =   960
      Picture         =   "Form4.frx":2C88
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DELETE"
      DownPicture     =   "Form4.frx":2EF5
      Height          =   855
      Left            =   4800
      Picture         =   "Form4.frx":33D9
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EDIT"
      DownPicture     =   "Form4.frx":35E4
      Height          =   855
      Index           =   0
      Left            =   2880
      Picture         =   "Form4.frx":3AF2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Options"
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
      Height          =   1215
      Left            =   720
      TabIndex        =   5
      Top             =   6360
      Width           =   5655
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Patron Number:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   0
      Left            =   1185
      TabIndex        =   22
      Top             =   1200
      Width           =   1110
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Patron Name:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   21
      Top             =   1755
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Address:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   20
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Telephone Number:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   19
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Email Address:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   525
      TabIndex        =   18
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      DataField       =   "Member Name"
      DataMember      =   "PatronDB"
      DataSource      =   "MLSDB"
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
      Height          =   255
      Left            =   2640
      TabIndex        =   17
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      DataField       =   "Telephone Number"
      DataMember      =   "PatronDB"
      DataSource      =   "MLSDB"
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
      Height          =   495
      Left            =   2685
      TabIndex        =   16
      Top             =   3000
      Width           =   5535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800000&
      DataField       =   "Email Address"
      DataMember      =   "PatronDB"
      DataSource      =   "MLSDB"
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
      Height          =   495
      Left            =   2685
      TabIndex        =   15
      Top             =   3600
      Width           =   5535
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Department:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   525
      TabIndex        =   14
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800000&
      DataField       =   "Department"
      DataMember      =   "PatronDB"
      DataSource      =   "MLSDB"
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
      Height          =   495
      Left            =   2685
      TabIndex        =   13
      Top             =   4200
      Width           =   5535
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Username:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   12
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Password:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   7
      Left            =   525
      TabIndex        =   11
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00800000&
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
      Height          =   495
      Left            =   2685
      TabIndex        =   10
      Top             =   4800
      Width           =   5535
   End
   Begin VB.Label Label7 
      BackColor       =   &H00800000&
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
      Height          =   375
      Left            =   2685
      TabIndex        =   9
      Top             =   5400
      Width           =   5535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      DataField       =   "Address"
      DataMember      =   "PatronDB"
      DataSource      =   "MLSDB"
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
      Height          =   615
      Left            =   2565
      TabIndex        =   8
      Top             =   2280
      Width           =   5655
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''THIS IS THE VIEW PATRON FORM''''''''''''''''''''''''''''''''

Private Sub Combo1_Click()  'Member Name Combobox
Dim found As Boolean
Command2.Enabled = True
Command3(0).Enabled = True
With MLSDB.rsPatronDB
.MoveFirst
found = False

While (Not .EOF) And (Not found)
      If Combo1.Text = MLSDB.rsPatronDB.Fields("Patron Number") Then
                found = True
                Label1.Caption = .Fields("Member Name")
                Label2.Caption = .Fields("Address")
                Label3.Caption = .Fields("Telephone Number")
                Label4.Caption = .Fields("Email Address")
                Label6.Caption = .Fields("Username")
                Label7.Caption = .Fields("Password")
                
     Else
                .MoveNext
            End If
        Wend
    End With
End Sub
Private Sub Command1_Click()    'Add Button
Form10.Show
Unload Me
End Sub
Private Sub Command2_Click()    'Delete Button
With MLSDB.rsPatronDB
Unload Form4

password.Show


End With
End Sub
Private Sub Command3_Click(Index As Integer)    'Edit Button
Form11.Show
Unload Me
End Sub
Private Sub Command4_Click()    'Exit Button
    Form2.Show
    Unload Me
End Sub
Private Sub Command5_Click()    'Refresh Button
Unload Me
Form4.Show
End Sub

Private Sub Command6_Click()
    helpclientprofile.Show
End Sub

Private Sub Form_Load()

    Command2.Enabled = False
    Command3(0).Enabled = False

If MLSDB.rsPatronDB.RecordCount = 0 Then
    MsgBox "The Record is currently empty", vbCritical
    
Else
    
With MLSDB.rsPatronDB

   .MoveFirst
    While (Not .EOF)
        Combo1.AddItem .Fields("Patron Number")
        .MoveNext
        Wend
        End With
End If

End Sub


