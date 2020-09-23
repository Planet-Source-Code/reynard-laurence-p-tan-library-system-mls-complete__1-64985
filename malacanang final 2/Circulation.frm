VERSION 5.00
Begin VB.Form Circulation 
   BackColor       =   &H00800000&
   Caption         =   "Circulation"
   ClientHeight    =   6030
   ClientLeft      =   6210
   ClientTop       =   3360
   ClientWidth     =   6660
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      DownPicture     =   "Circulation.frx":0000
      Height          =   735
      Left            =   5880
      Picture         =   "Circulation.frx":09B5
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      DownPicture     =   "Circulation.frx":1370
      Height          =   855
      Left            =   4080
      Picture         =   "Circulation.frx":1792
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RETURN"
      DownPicture     =   "Circulation.frx":18B0
      Height          =   855
      Left            =   1320
      Picture         =   "Circulation.frx":1D94
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      DataMember      =   "Circulation"
      DataSource      =   "MLSDB"
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label5 
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
      TabIndex        =   15
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Borrower:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   14
      Top             =   2520
      Width           =   1815
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
      TabIndex        =   13
      Top             =   3120
      Width           =   3135
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
      TabIndex        =   12
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   2520
      TabIndex        =   11
      Top             =   1200
      Width           =   3135
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
      Index           =   4
      Left            =   2520
      TabIndex        =   9
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      DataMember      =   "BookDbase"
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
      Index           =   1
      Left            =   2400
      TabIndex        =   8
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Issued By:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   7
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Date Issued:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   6
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Borrower ID:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   4
      Left            =   1530
      TabIndex        =   5
      Top             =   1920
      Width           =   885
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Book Title:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   1
      Left            =   1665
      TabIndex        =   4
      Top             =   1335
      Width           =   765
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Circulation Number:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   0
      Left            =   1035
      TabIndex        =   3
      Top             =   720
      Width           =   1380
   End
End
Attribute VB_Name = "Circulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''THIS IS THE CIRCULATION FORM''''''''''''''''''''''''''''''''

Private Sub Combo1_Click()  'BookID Combobox
Dim found As Boolean
Command1.Enabled = True
With MLSDB.rsCirculation
.MoveFirst
found = False

While (Not .EOF) And (Not found)
      If Combo1.Text = MLSDB.rsCirculation.Fields("BookID") Then
                found = True
                Label5.Caption = .Fields("Member Name")
                Label3.Caption = .Fields("Patron Number")
                Label1(4).Caption = .Fields("Date Issued")
                Label2.Caption = .Fields("BookTitle")
                Label4.Caption = .Fields("Administrator ID")
                
                
        Else
                .MoveNext
        End If
Wend
End With
End Sub

Private Sub Command1_Click()    'Return Button
Dim found As Boolean
With MLSDB.rsBookDbase
.MoveFirst
found = False

While (Not .EOF) And (Not found)
      If Label2.Caption = MLSDB.rsBookDbase.Fields("BookTitle") Then
        found = True
      Else
        .MoveNext
      End If
Wend
Unload Circulation

If MsgBox("Return Book?", vbInformation + vbYesNo) = vbYes Then
        .Fields("In / Out") = "in"
        MLSDB.rsCirculation.Delete
        MLSDB.rsCirculation.MoveFirst
        
        Circulation.Show
  Else
        Circulation.Show
End If
End With
End Sub

Private Sub Command2_Click()    'Exit Button
Form5.Show
Unload Me
End Sub


Private Sub Command6_Click()
helpcirculation.Show
End Sub

Private Sub Form_Load()
Command1.Enabled = False

If MLSDB.rsCirculation.RecordCount = 0 Then
    MsgBox "The Record is currently empty", vbCritical
    
Else
    
With MLSDB.rsCirculation
    .MoveFirst
    While (Not .EOF)
        Combo1.AddItem .Fields("BookID")
        .MoveNext
    Wend
End With
End If
End Sub
