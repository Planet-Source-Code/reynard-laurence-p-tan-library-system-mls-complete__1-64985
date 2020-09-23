VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00800000&
   Caption         =   "Book Information"
   ClientHeight    =   7545
   ClientLeft      =   5625
   ClientTop       =   3360
   ClientWidth     =   7785
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   7785
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      DownPicture     =   "Form5.frx":0000
      Height          =   735
      Left            =   6960
      Picture         =   "Form5.frx":09B5
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "CIRCULATION"
      DownPicture     =   "Form5.frx":1370
      Height          =   855
      Left            =   6240
      Picture         =   "Form5.frx":1783
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "REFRESH"
      DownPicture     =   "Form5.frx":187A
      Height          =   735
      Left            =   120
      Picture         =   "Form5.frx":1EAE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ISSUE"
      DownPicture     =   "Form5.frx":2146
      Height          =   855
      Left            =   4560
      Picture         =   "Form5.frx":2564
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      DataMember      =   "BookDbase"
      DataSource      =   "MLSDB"
      Height          =   315
      Left            =   3240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   3375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "EDIT"
      DownPicture     =   "Form5.frx":26BB
      Height          =   855
      Left            =   1680
      Picture         =   "Form5.frx":2BC9
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MAIN MENU"
      DownPicture     =   "Form5.frx":2F36
      Height          =   855
      Left            =   6240
      Picture         =   "Form5.frx":34D6
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DELETE"
      DownPicture     =   "Form5.frx":3AA8
      Height          =   855
      Left            =   3120
      Picture         =   "Form5.frx":3F8C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      DownPicture     =   "Form5.frx":4197
      Height          =   855
      Left            =   240
      Picture         =   "Form5.frx":4671
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   1335
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
      Left            =   120
      TabIndex        =   23
      Top             =   5760
      Width           =   5895
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00800000&
      Caption         =   "Go to"
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
      Height          =   2175
      Left            =   6120
      TabIndex        =   24
      Top             =   5280
      Width           =   1575
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
      Height          =   255
      Left            =   3240
      TabIndex        =   22
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Status:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   8
      Left            =   2640
      TabIndex        =   21
      Top             =   5100
      Width           =   495
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
      Left            =   3240
      TabIndex        =   20
      Top             =   4560
      Width           =   3375
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
      Height          =   375
      Left            =   3240
      TabIndex        =   19
      Top             =   3960
      Width           =   3375
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
      Left            =   3240
      TabIndex        =   18
      Top             =   3360
      Width           =   3375
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
      Left            =   3240
      TabIndex        =   17
      Top             =   2760
      Width           =   3375
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
      Left            =   3240
      TabIndex        =   16
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label Label2 
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
      Height          =   615
      Left            =   3240
      TabIndex        =   15
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Responsibility Center:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   7
      Left            =   1320
      TabIndex        =   14
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Year of Acquisition:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   1320
      TabIndex        =   13
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Price:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   1320
      TabIndex        =   12
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Book ID:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   4
      Left            =   2505
      TabIndex        =   11
      Top             =   2760
      Width           =   630
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Call Number:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   1320
      TabIndex        =   10
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Author:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   1260
      TabIndex        =   9
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Book Title:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   8
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''THIS IS THE BOOK INFORMATION FORM'''''''''''''''''''''''''''

Private Sub Combo2_Click()  'Book Title Combobox
    Command2.Enabled = True
    Command5.Enabled = True
   
    
    Dim found As Boolean
    With MLSDB.rsBookDbase
        .MoveFirst
        found = False
        While (Not .EOF) And (Not found)
            If Combo2.Text = MLSDB.rsBookDbase.Fields("BookTitle").Value Then
                found = True
                Label2.Caption = .Fields("Author")
                Label3.Caption = .Fields("CallNumber")
                Label4.Caption = .Fields("BookID")
                Label5.Caption = .Fields("Price")
                Label6.Caption = .Fields("Year of Acquisition")
                Label7.Caption = .Fields("Responsibility Center")
                Label8.Caption = .Fields("In / Out")
            If Label8.Caption = "out" Then
             Command3.Enabled = False
             Command2.Enabled = False
             Else
             Command3.Enabled = True
             End If
            Else
                .MoveNext
            End If
        Wend
    End With
End Sub
Private Sub Command1_Click()    'Add Button
Form13.Show
Unload Me
End Sub
Private Sub Command2_Click()    'Delete Button
With MLSDB.rsBookDbase
Unload Form5
If MsgBox("Delete Record?", vbInformation + vbYesNo) = vbYes Then
        MLSDB.rsBookDbase.Delete
        MLSDB.rsBookDbase.MoveFirst
        
        Form5.Show
  Else
        Form5.Show
End If
End With
End Sub
Private Sub Command3_Click()    'Issue Button
Issue.Show
Unload Me
End Sub
Private Sub Command4_Click()    'Exit Button
    Form2.Show
    Unload Me
End Sub
Private Sub Command5_Click()    'Edit Button
Form12.Show
Unload Me
End Sub
Private Sub Command6_Click()    'Refresh Button
Unload Me
Form5.Show
End Sub
Private Sub Command7_Click()    'Circulation Form Button
Circulation.Show
Unload Me
End Sub

Private Sub Command8_Click()
helpbookinfo.Show
End Sub

Private Sub Form_Load()

    Command2.Enabled = False
    Command5.Enabled = False
    Command3.Enabled = False
If MLSDB.rsBookDbase.RecordCount = 0 Then
    MsgBox "The Record is currently empty", vbCritical
    
Else
    With MLSDB.rsBookDbase
        .MoveFirst
        While (Not .EOF)
            Combo2.AddItem .Fields("BookTitle")
            .MoveNext
        Wend
    End With
End If
End Sub

