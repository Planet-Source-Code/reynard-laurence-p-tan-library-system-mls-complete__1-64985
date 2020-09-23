VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   Caption         =   "Simple Search - Call Number"
   ClientHeight    =   8670
   ClientLeft      =   2805
   ClientTop       =   1725
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   8352.1
   ScaleMode       =   0  'User
   ScaleWidth      =   10170
   Begin VB.CommandButton Exit 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8640
      Picture         =   "PSearch3.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      Picture         =   "PSearch3.frx":011E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Call Number"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      Picture         =   "PSearch3.frx":025A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Author"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      Picture         =   "PSearch3.frx":03C6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
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
      TabIndex        =   0
      Top             =   3480
      Width           =   3135
   End
   Begin VB.ListBox lstdata 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   3180
      ItemData        =   "PSearch3.frx":04D5
      Left            =   6240
      List            =   "PSearch3.frx":04D7
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   4200
      Width           =   3735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      X1              =   480
      X2              =   9120
      Y1              =   1849.6
      Y2              =   1849.6
   End
   Begin VB.Image Image1 
      Height          =   1755
      Left            =   240
      Picture         =   "PSearch3.frx":04D9
      Top             =   120
      Width           =   1875
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
      Left            =   2160
      TabIndex        =   26
      Top             =   7680
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
      Left            =   2160
      TabIndex        =   25
      Top             =   7200
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
      Left            =   2160
      TabIndex        =   24
      Top             =   6720
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
      Left            =   2160
      TabIndex        =   23
      Top             =   6240
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
      Index           =   3
      Left            =   2160
      TabIndex        =   22
      Top             =   5760
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
      Index           =   2
      Left            =   2160
      TabIndex        =   21
      Top             =   5280
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
      Height          =   615
      Index           =   1
      Left            =   2160
      TabIndex        =   20
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
      Index           =   0
      Left            =   2160
      TabIndex        =   19
      Top             =   4200
      Width           =   3495
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Responsibility Center:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   18
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Year of Acquisition:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   17
      Top             =   7200
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
      Left            =   315
      TabIndex        =   16
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Barcode Number:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   315
      TabIndex        =   15
      Top             =   6270
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
      Left            =   315
      TabIndex        =   14
      Top             =   5760
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
      Left            =   315
      TabIndex        =   13
      Top             =   5280
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
      Left            =   315
      TabIndex        =   12
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Book ID:"
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Index           =   0
      Left            =   315
      TabIndex        =   11
      Top             =   4230
      Width           =   1815
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
      TabIndex        =   10
      Top             =   3120
      Width           =   3855
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
      TabIndex        =   9
      Top             =   1920
      Width           =   8775
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
      TabIndex        =   8
      Top             =   360
      Width           =   8295
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Status:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   8
      Left            =   1560
      TabIndex        =   7
      Top             =   8145
      Width           =   495
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
      Left            =   2160
      TabIndex        =   6
      Top             =   8040
      Width           =   2655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      X1              =   480
      X2              =   9120
      Y1              =   1849.6
      Y2              =   1849.6
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                         A Simple Database Example                       '''
'''                          Written By Darren Kurn                         '''
'''                                 28/09/01                                '''
'''                                   xdaz                                  '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Firstly, For those of you who have never used Databases with your applications
'before, please make a point of going to the project menu, and clicking on
'references. Here you will notice that one of the references is "Microsoft
'DAO Object Library". This is the component which is required to access the
'Database.

'Right,, Lets Start The Example.

'First, we need to dim some variables
'We will start with the DB Variable. This Variable simply tells VB Which Database
'we intend to work from
Dim db As Database
'Next, we have the rs variable. this tells VB which recordset, or table within
'the database that we intend to work from
Dim rs As Recordset
'The ws variable tells VB the workspace. Sorry, but im not entirely sure what this
'actually does, but i do know that it is essential for the project to work
Dim ws As Workspace
'The max variable will eventually hold the number of records in the database
'table, so we can use the variable as a loop control variable
Dim max As Long
'The i variable is just another loop control variable
Dim i As Long
'This variable is used to store any answers from message boxes
Dim errormsg
'these two variables are here so that we know whether we want to add new data to the
'database, or edit existing data
Dim dbadd As Boolean
Dim dbedit As Boolean

Private Sub Command1_Click()
PSearch.Show
Unload Me
End Sub

Private Sub Command2_Click()
Form7.Show
Unload Me
End Sub

Private Sub Command3_Click()
Form1.Refresh
End Sub

Private Sub Exit_Click()
Form3.Show
Unload Me
End Sub


Private Sub Form_Load()
'The first thing that we will need to do when the form loads
'is tell VB which table we are going to work from
Set ws = DBEngine.Workspaces(0)

'Here we set the database. This will tell VB to use the database called
'"Database" which can be found in the same directory as this project.
'Obviously, you do not always have to keep the database in the same directory,
'but i find that it helps!
Set db = ws.OpenDatabase(App.Path & "\db1.mdb")

'Here we set the table within the database which we intend to use.
'As we only have one table in our database, this is easy enough.
Set rs = db.OpenRecordset("BookDatabase", dbOpenTable)

'Next, we will call the List function which will get some of the
'Information out of the database, and place it into the list box
list


With MLSDB.rsBookDbase
.MoveFirst
End With
End Sub


Private Function list()
'This Function will extract Surnames from the database, and
'Place them into the list box for selection

'If there are no records in the database table, then we cannot extract
'any data, so we have to exit the function. If we do not do this, then
'The program will throw up an error, and crash
If rs.RecordCount = 0 Then
    errormsg = MsgBox("No Records Found", , "Error")
    'If no records have been found, then it is very likely that the user
    'is using the search field, so we will set the text box back to what
    'it was before the error came up
    If Len(txtSearch.Text) > 0 Then
        txtSearch.Text = Mid(txtSearch.Text, 1, Len(txtSearch.Text) - 1)
    Else
        Exit Function
    End If
End If
'Move to the first record in the database
rs.MoveLast
'Move to the last record in the database
rs.MoveFirst
'You're probably wondering why i have just moved to the first, and then
'the last record in the database. Well, I find that Access does not report
'the number of records in the table accurately. I have no idea Why. It's a
'microsoft thing. however, going to the last, and then the first record usually
'helps Access report the accurate number of records
'Next, we need to set variable "max" to the number of records in the database
max = rs.RecordCount
'Now we move back to the first record in the database
rs.MoveFirst
'Now we need to clear our list box, so that we do not have repeating data
lstdata.Clear
'Now we can start a loop, which will in turn extract data from each of the
'records in the database
For i = 1 To max
    'For each Entry, we want to put the surname into the list box
    'rs("Surname") simply tells VB to get the data from the surname field
    lstdata.AddItem rs("CallNumber")
    'Then we need to move to the next record in the table
    rs.MoveNext
'repeat the loop
Next i

End Function


Private Sub lstdata_Click()

'Once all our surnames are into the list box, we need to control what happens
'when the user click on one of the surnames
'The line below will send an SQL command to the database which tells access
'to create a new theoretical table which only holds the information on
'the selected surname. this works exactly the same way as making a query in the
'database to only show data on a person with a certain surname
Set rs = db.OpenRecordset("Select * from BookDatabase where CallNumber = '" & Trim(lstdata.list(lstdata.ListIndex)) & "'")
'move to the first record in the table
rs.MoveFirst
'Next, we need to extract the information out of the database, and put
'it in the relevent text boxes
Label4(0).Caption = rs("Book Id")
Label4(1).Caption = rs("BookTitle")
Label4(2).Caption = rs("Author")
Label4(3).Caption = rs("CallNumber")
Label4(4).Caption = rs("Barcode Number")
Label4(5).Caption = rs("Price")
Label4(6).Caption = rs("Year of Acquisition")
Label4(7).Caption = rs("Responsibility Center")
Label5.Caption = rs("In / Out")
'Now we need to enable some of the command buttons so that we can work on the
'data

End Sub


Private Sub txtSearch_Change()
'This is a nice little function which will allow the user to search for a surname
'in the database. Most modern databases are very large, and so allow the user
'to enter the first few letters of the desired surname in order to limit the
'selection.
'If the search filed contains no text, then we have to tell VB to read from
'the entire recordset

If txtSearch.Text = vbNullString Then
    Set rs = db.OpenRecordset("BookDatabase", dbOpenTable)
'if the search field contains text, then we need to tell VB to search for surnames
'beginning with the letters entered in the search field
Else
    Set rs = db.OpenRecordset("SELECT * FROM BookDatabase WHERE CallNumber LIKE '" & txtSearch.Text & "'" & "& '*'")
End If
'then we need to re-populate the list box with only those surnames beginning with
'the letters in the search field
list

End Sub



