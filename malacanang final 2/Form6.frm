VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00800000&
   Caption         =   "Administrator Profile"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8490
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      DataMember      =   "Administrator"
      DataSource      =   "MLSDB"
      Height          =   315
      Left            =   2520
      TabIndex        =   7
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CommandButton Command5 
      DownPicture     =   "Form6.frx":0000
      Height          =   735
      Left            =   7680
      Picture         =   "Form6.frx":09B5
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MAIN MENU"
      DownPicture     =   "Form6.frx":1370
      Height          =   855
      Left            =   6720
      Picture         =   "Form6.frx":1910
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "REFRESH"
      DownPicture     =   "Form6.frx":1EE2
      Height          =   735
      Left            =   120
      Picture         =   "Form6.frx":2516
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DELETE"
      DownPicture     =   "Form6.frx":27AE
      Height          =   855
      Left            =   4560
      Picture         =   "Form6.frx":2C92
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ADD"
      DownPicture     =   "Form6.frx":2E9D
      Height          =   855
      Left            =   720
      Picture         =   "Form6.frx":3377
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EDIT"
      DownPicture     =   "Form6.frx":35E4
      Height          =   855
      Left            =   2640
      Picture         =   "Form6.frx":3AF2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Frame Frame2 
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
      Left            =   600
      TabIndex        =   6
      Top             =   4680
      Width           =   5535
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Administrator ID:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   16
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Administrator Name:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   15
      Top             =   1935
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Address:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   14
      Top             =   2520
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
      Left            =   360
      TabIndex        =   13
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
      Left            =   360
      TabIndex        =   12
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label Label3 
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
      Left            =   2520
      TabIndex        =   10
      Top             =   3000
      Width           =   5415
   End
   Begin VB.Label Label4 
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
      Left            =   2520
      TabIndex        =   9
      Top             =   3600
      Width           =   5415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      DataField       =   "Address"
      DataMember      =   "Administrator"
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
      Left            =   2520
      TabIndex        =   8
      Top             =   2400
      Width           =   5535
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''THIS IS THE ADMINISTRATOR PROFILE FORM'''''''''''''''''''''''''''

Private Sub Combo1_Click()  'Admin Name Combobox
Command1.Enabled = True
Command4.Enabled = True

With MLSDB.rsAdministrator

If MLSDB.rsAdministrator.RecordCount = 0 Then
    MsgBox "The Record is currently empty", vbCritical
    
Else
    
With MLSDB.rsAdministrator
.MoveFirst
found = False

While (Not .EOF) And (Not found)
        If Combo1.Text = MLSDB.rsAdministrator.Fields("Administrator ID") Then
                found = True
                Label1.Caption = .Fields("Administrator Name")
                Label2.Caption = .Fields("Address")
                Label3.Caption = .Fields("Telephone Number")
                Label4.Caption = .Fields("Email Address")
                
        Else
                .MoveNext
            End If
        Wend
    End With
End If
End With
End Sub
Private Sub Command1_Click()    'Edit Button
EditAdmin.Show
Unload Me
End Sub
Private Sub Command2_Click()    'Exit Button
    Form2.Show
    Unload Me
End Sub
Private Sub Command3_Click()    'Add Button
AddAdmin.Show
Unload Me
End Sub
Private Sub Command4_Click()    'Delete Button
With MLSDB.rsAdministrator
Unload Form6
If MsgBox("Delete Record?", vbInformation + vbYesNo) = vbYes Then
        MLSDB.rsAdministrator.Delete
        MLSDB.rsAdministrator.MoveFirst
    If MLSDB.rsAdministrator.RecordCount = 0 Then
        AdminReg.Show
        Unload Me
    Else
        Form6.Show
    End If
Else
        Form6.Show
End If
End With
End Sub
Private Sub Command5_Click()    'Refresh Button
helpadminprofile.Show
End Sub

Private Sub Command6_Click()
    Unload Me
    Form6.Show
End Sub

Private Sub Form_Load()
  Command1.Enabled = False
  Command4.Enabled = False
  

   
If MLSDB.rsAdministrator.RecordCount = 0 Then
    MsgBox "The Record is currently empty", vbCritical
    AdminReg.Show
    Unload Me
Else
    
With MLSDB.rsAdministrator

   .MoveFirst
    While (Not .EOF)
        Combo1.AddItem .Fields("Administrator ID")
        .MoveNext
    Wend
End With
End If
End Sub


