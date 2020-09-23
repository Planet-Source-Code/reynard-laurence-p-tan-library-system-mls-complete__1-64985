VERSION 5.00
Begin VB.Form PSearch 
   BackColor       =   &H00800000&
   Caption         =   "Simple Search - Search by Title"
   ClientHeight    =   8985
   ClientLeft      =   1995
   ClientTop       =   1575
   ClientWidth     =   10200
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      DownPicture     =   "PSearch.frx":0000
      Height          =   735
      Left            =   9360
      Picture         =   "PSearch.frx":09B5
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MAIN MENU"
      DownPicture     =   "PSearch.frx":1370
      Height          =   855
      Left            =   7440
      Picture         =   "PSearch.frx":1910
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8040
      Width           =   1335
   End
   Begin VB.ListBox lstdata 
      Height          =   3375
      ItemData        =   "PSearch.frx":1EE2
      Left            =   6240
      List            =   "PSearch.frx":1EE4
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   3840
      Width           =   3735
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6480
      MaxLength       =   50
      TabIndex        =   2
      Top             =   3240
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search by Author"
      DownPicture     =   "PSearch.frx":1EE6
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      Picture         =   "PSearch.frx":22F6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Search by Title"
      DownPicture     =   "PSearch.frx":2405
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      Picture         =   "PSearch.frx":2852
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Select a Search Domain"
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
      Left            =   1200
      TabIndex        =   21
      Top             =   2880
      Width           =   4095
   End
   Begin VB.Label Label4 
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
      Index           =   2
      Left            =   2280
      TabIndex        =   9
      Top             =   5400
      Width           =   3495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      X1              =   0
      X2              =   10200
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Welcome to Malacañang Library. We are glad to be of service."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   975
      Left            =   480
      TabIndex        =   24
      Top             =   1920
      Width           =   8775
   End
   Begin VB.Image Image1 
      Height          =   1755
      Left            =   240
      Picture         =   "PSearch.frx":298E
      Top             =   120
      Width           =   1875
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
      Height          =   255
      Left            =   2280
      TabIndex        =   20
      Top             =   8520
      Width           =   2655
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Status:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   8
      Left            =   1680
      TabIndex        =   19
      Top             =   8505
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   "Enter keyword(s) to search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6360
      TabIndex        =   18
      Top             =   2880
      Width           =   3855
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Book Title:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   17
      Top             =   4680
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
      Left            =   360
      TabIndex        =   16
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Call Number:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   15
      Top             =   6120
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
      Left            =   1545
      TabIndex        =   14
      Top             =   6600
      Width           =   630
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Price:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   13
      Top             =   7080
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
      Left            =   360
      TabIndex        =   12
      Top             =   7560
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Responsibility Center:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   11
      Top             =   8040
      Width           =   1935
   End
   Begin VB.Label Label4 
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
      Height          =   975
      Index           =   1
      Left            =   2280
      TabIndex        =   10
      Top             =   4560
      Width           =   3975
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
      Index           =   3
      Left            =   2280
      TabIndex        =   8
      Top             =   6120
      Width           =   3495
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
      Index           =   4
      Left            =   2280
      TabIndex        =   7
      Top             =   6600
      Width           =   3495
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
      Index           =   5
      Left            =   2280
      TabIndex        =   6
      Top             =   7080
      Width           =   3495
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
      Index           =   6
      Left            =   2280
      TabIndex        =   5
      Top             =   7560
      Width           =   3495
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
      Index           =   7
      Left            =   2280
      TabIndex        =   4
      Top             =   8040
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "MALACAÑANG LIBRARY SYSTEM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1155
      Left            =   1560
      TabIndex        =   25
      Top             =   360
      Width           =   8295
   End
End
Attribute VB_Name = "PSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''THIS IS THE PATRON SEARCH BOOKTITLE FORM'''''''''''''''''''''''''

Option Explicit
Dim db As Database
Dim rs As Recordset
Dim ws As Workspace
Dim max As Long
Dim i As Long
Dim errormsg
Private Sub Command1_Click()
PSearch.Refresh
End Sub
Private Sub Command2_Click()
Form7.Show
Unload Me
End Sub
Private Sub Exit_Click()
Form3.Show
Unload Me
End Sub

Private Sub Command4_Click()
    Unload Me
    Form3.Show
End Sub

Private Sub Command5_Click()
helpsearchtitle2.Show
End Sub

Private Sub Form_Load()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\db1.mdb")
Set rs = db.OpenRecordset("BookDatabase", dbOpenTable)
list

With MLSDB.rsBookDbase
.MoveFirst
End With
End Sub
Private Function list()
If rs.RecordCount = 0 Then
    errormsg = MsgBox("No Records Found", , "Error")
    If Len(txtSearch.Text) > 0 Then
        txtSearch.Text = Mid(txtSearch.Text, 1, Len(txtSearch.Text) - 1)
    Else
        Exit Function
    End If
End If
rs.MoveLast
rs.MoveFirst
max = rs.RecordCount
rs.MoveFirst
lstdata.Clear
For i = 1 To max
    lstdata.AddItem rs("BookTitle")
    rs.MoveNext
Next i

End Function

Private Sub lstdata_Click()
Call xListKillDupes(lstdata) 'calls sub from module
Set rs = db.OpenRecordset("Select * from BookDatabase where BookTitle = '" & Trim(lstdata.list(lstdata.ListIndex)) & "'")
 Dim found As Boolean
    With MLSDB.rsBookDbase
        .MoveFirst
        found = False
        While (Not .EOF) And (Not found)
            If lstdata.Text = MLSDB.rsBookDbase.Fields("BookTitle").Value Then
                found = True
                Label4(1).Caption = .Fields("BookTitle")
                Label4(2).Caption = .Fields("Author")
                Label4(3).Caption = .Fields("CallNumber")
                Label4(4).Caption = .Fields("BookID")
                Label4(5).Caption = .Fields("Price")
                Label4(6).Caption = .Fields("Year of Acquisition")
                Label4(7).Caption = .Fields("Responsibility Center")
                Label5.Caption = .Fields("In / Out")
            Else
                .MoveNext
            End If
        Wend
    End With

End Sub
Private Sub txtSearch_Change()
If txtSearch.Text = vbNullString Then
    Set rs = db.OpenRecordset("BookDatabase", dbOpenTable)
Else
    Set rs = db.OpenRecordset("SELECT * FROM BookDatabase WHERE BookTitle LIKE '" & txtSearch.Text & "'" & "& '*'")
End If
list

End Sub
Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 32 Or KeyAscii = 46) Then
Else
KeyAscii = 0
End If
End Sub
