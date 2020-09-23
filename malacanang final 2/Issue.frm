VERSION 5.00
Begin VB.Form Issue 
   BackColor       =   &H00800000&
   Caption         =   "Issue Books"
   ClientHeight    =   6585
   ClientLeft      =   6015
   ClientTop       =   3165
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      DataMember      =   "PatronDB"
      DataSource      =   "MLSDB"
      Height          =   315
      Left            =   2760
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      DownPicture     =   "Issue.frx":0000
      Height          =   855
      Left            =   5640
      Picture         =   "Issue.frx":0422
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      DownPicture     =   "Issue.frx":0540
      Height          =   735
      Left            =   6360
      Picture         =   "Issue.frx":0EF5
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox Combo2 
      DataMember      =   "Administrator"
      DataSource      =   "MLSDB"
      Height          =   315
      Left            =   2760
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ISSUE"
      DownPicture     =   "Issue.frx":18B0
      Height          =   855
      Left            =   5040
      Picture         =   "Issue.frx":1CCE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox txtDateIssued 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3960
      Width           =   2040
   End
   Begin VB.Label Label2 
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
      Left            =   2760
      TabIndex        =   14
      Top             =   3120
      Width           =   4215
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Borrower Name:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   2
      Left            =   1515
      TabIndex        =   13
      Top             =   3120
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      DataField       =   "BookID"
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
      Left            =   2760
      TabIndex        =   12
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Book ID:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   0
      Left            =   2025
      TabIndex        =   11
      Top             =   720
      Width           =   630
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      DataField       =   "BookTitle"
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
      Height          =   615
      Left            =   2760
      TabIndex        =   9
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Issued By:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   840
      TabIndex        =   8
      Top             =   4800
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
      Left            =   840
      TabIndex        =   7
      Top             =   3990
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
      Left            =   1770
      TabIndex        =   6
      Top             =   2280
      Width           =   885
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Book Title:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
   End
End
Attribute VB_Name = "Issue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''THIS IS THE ISSUE FORM''''''''''''''''''''''''''''''''

Private Sub Combo1_Change()
Combo2.SetFocus
End Sub

Private Sub Combo1_Click()  'Member Name Combobox
Dim found As Boolean
With MLSDB.rsPatronDB
.MoveFirst
found = False

While (Not .EOF) And (Not found)
      If Combo1.Text = MLSDB.rsPatronDB.Fields("Patron Number") Then
            Label2.Caption = .Fields("Member Name")
            
                found = True
                
     Else
                .MoveNext
            End If
        Wend
    End With
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
Else
KeyAscii = 0
End If
End Sub

Private Sub Combo2_Change()
Command2.SetFocus
End Sub

Private Sub Combo2_Click()  'Admin Name Combobox
Dim found As Boolean
With MLSDB.rsAdministrator
.MoveFirst
found = False

While (Not .EOF) And (Not found)
      If Combo2.Text = MLSDB.rsAdministrator.Fields("Administrator ID") Then
                found = True
                
     Else
                .MoveNext
            End If
        Wend
    End With
End Sub


Private Sub Combo2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
Else
KeyAscii = 0
End If
End Sub

Private Sub Command1_Click()    'Issue Button
Dim found2 As Boolean
Dim found3 As Boolean
one = False
two = False
found2 = False
found3 = False

With MLSDB.rsPatronDB
.MoveFirst
    While (Not .EOF) And (Not found2)
    If MLSDB.rsPatronDB.Fields("Patron Number").Value = Combo1.Text Then
        found2 = True
    Else
        .MoveNext
    End If
    Wend
End With

If found2 = False Then
    MsgBox "Patron Number does not exist", vbCritical
Else


With MLSDB.rsAdministrator
.MoveFirst
    While (Not .EOF) And (Not found3)
    If MLSDB.rsAdministrator.Fields("Administrator ID").Value = Combo2.Text Then
        found3 = True
    Else
        .MoveNext
    End If
    Wend
End With

If found3 = False Then
    MsgBox "Administrator ID does not exist", vbCritical
Else

    If Combo1.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        Combo1.SetFocus
    ElseIf txtDateIssued.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtDateIssued.SetFocus
    ElseIf Combo2.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        Combo2.SetFocus

Else
        With MLSDB.rsBookDbase
        .Fields("In / Out") = "out"
        .Update
        End With
        
        With MLSDB.rsCirculation
        .AddNew
        .Fields("BookID") = Label1.Caption
        .Fields("BookTitle") = Label3.Caption
        .Fields("Patron Number") = Combo1.Text
        .Fields("Date Issued") = Trim(txtDateIssued.Text)
        .Fields("Administrator ID") = Combo2.Text
        .Fields("Member Name") = Label2.Caption
        .Update
        .MoveFirst
        End With

        
    MsgBox "Report Successfully added!", vbInformation
Form2.Show
Unload Me
End If
End If
End If
End Sub

Private Sub Command2_Click()
    Form2.Show
    Unload Me
End Sub

Private Sub Command3_Click()
With MLSDB.rsBookDbase
If MLSDB.rsBookDbase.RecordCount = 0 Then
    MsgBox "The Record is empty", vbCritical
    Form2.Refresh
Else
    Form8.Show
    Unload Me
End If
End With

End Sub

Private Sub Command4_Click()
    Form9.Show
    Unload Me
End Sub

Private Sub Command6_Click()
    helpissue.Show
End Sub

Private Sub Command7_Click()
    Form2.Show
    Unload Me
End Sub

Private Sub Form_Load()
Command1.Enabled = False
Dim x As Long

Appdate = Date
txtDateIssued.Text = Appdate

If MLSDB.rsPatronDB.RecordCount = 0 Then
    MsgBox "There are no Patrons", vbCritical
    
Else
    
With MLSDB.rsPatronDB
    .MoveFirst
    While (Not .EOF)
        Combo1.AddItem .Fields("Patron Number")
        .MoveNext
    Wend
End With

With MLSDB.rsAdministrator
    .MoveFirst
    While (Not .EOF)
        Combo2.AddItem .Fields("Administrator ID")
        .MoveNext
    Wend
End With
End If
End Sub

Private Sub Label2_Change()
Command1.Enabled = True
End Sub
