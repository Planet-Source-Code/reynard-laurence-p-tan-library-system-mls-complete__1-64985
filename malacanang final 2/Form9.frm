VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00800000&
   Caption         =   "Barcode Search"
   ClientHeight    =   8205
   ClientLeft      =   4290
   ClientTop       =   3165
   ClientWidth     =   10635
   FillColor       =   &H00800000&
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   10635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      DownPicture     =   "Form9.frx":0000
      Height          =   735
      Left            =   9840
      Picture         =   "Form9.frx":09B5
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MAIN MENU"
      DownPicture     =   "Form9.frx":1370
      Height          =   855
      Left            =   6120
      Picture         =   "Form9.frx":1910
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ISSUE"
      DownPicture     =   "Form9.frx":1EE2
      Height          =   855
      Left            =   4560
      Picture         =   "Form9.frx":2300
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      DataMember      =   "BookDbase"
      DataSource      =   "MLSDB"
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label8 
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
      Left            =   4800
      TabIndex        =   18
      Top             =   5160
      Width           =   2775
   End
   Begin VB.Label Label12 
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
      Left            =   4800
      TabIndex        =   17
      Top             =   6240
      Width           =   2775
   End
   Begin VB.Label Label11 
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
      Left            =   4800
      TabIndex        =   16
      Top             =   5880
      Width           =   2775
   End
   Begin VB.Label Label10 
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
      Left            =   4800
      TabIndex        =   15
      Top             =   5520
      Width           =   2775
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
      Left            =   4800
      TabIndex        =   14
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   4440
      Width           =   4935
   End
   Begin VB.Label Label5 
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
      Height          =   735
      Left            =   4800
      TabIndex        =   12
      Top             =   3720
      Width           =   4935
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Status:"
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Index           =   8
      Left            =   4080
      TabIndex        =   11
      Top             =   6240
      Width           =   495
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Responsibility Center:"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   7
      Left            =   2760
      TabIndex        =   10
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Year of Acquisition:"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   6
      Left            =   2760
      TabIndex        =   9
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Price:"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   8
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Call Number:"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   7
      Top             =   4785
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Author:"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   6
      Top             =   4410
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Book Title:"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   5
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   "Information of the Books"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   2760
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   "Scan or enter a barcode number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      X1              =   240
      X2              =   10440
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Barcode Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''THIS IS THE BARCODE SEARCH FORM''''''''''''''''''''''''''

Private Sub Combo1_Change()
Command2.SetFocus
End Sub

Private Sub Combo1_Click()  'BookID Search Field
Dim found As Boolean
Command2.Enabled = True
With MLSDB.rsBookDbase
.MoveFirst
found = False

While (Not .EOF) And (Not found)
      If Combo1.Text = MLSDB.rsBookDbase.Fields("BookID").Value Then
                found = True
                Label5.Caption = .Fields("BookTitle")
                Label6.Caption = .Fields("Author")
                Label7.Caption = .Fields("CallNumber")
                Label8.Caption = .Fields("Price")
                Label10.Caption = .Fields("Year of Acquisition")
                Label11.Caption = .Fields("Responsibility Center")
                Label12.Caption = .Fields("In / Out")
                
            If Label12.Caption = "out" Then
             Command2.Enabled = False
            Else
             Command2.Enabled = True
            End If
     Else
                .MoveNext
    End If
Wend
End With
End Sub
Private Sub Command1_Click()    'Exit Button
Form2.Show
Unload Me
End Sub

Private Sub Command2_Click()    'Issue Button
Issue.Show
Unload Me
End Sub

Private Sub Command3_Click()
    Form2.Show
    Unload Me
End Sub

Private Sub Command6_Click()
    helpbcode.Show
End Sub

Private Sub Form_Load()
Command2.Enabled = False
If MLSDB.rsBookDbase.RecordCount = 0 Then
    MsgBox "The Record is currently empty", vbCritical
    
Else
    With MLSDB.rsBookDbase
        .MoveFirst
        While (Not .EOF)
            Combo1.AddItem .Fields("BookID")
            .MoveNext
        Wend
    End With
End If
End Sub

