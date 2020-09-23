VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6465
   ClientLeft      =   4635
   ClientTop       =   3555
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   9540
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3120
      TabIndex        =   20
      Text            =   "Combo1"
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   2055
   End
   Begin VB.ListBox lstdata 
      Height          =   2595
      Left            =   240
      TabIndex        =   0
      Top             =   2880
      Width           =   2055
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
      Left            =   4560
      TabIndex        =   19
      Top             =   5880
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
      Left            =   4560
      TabIndex        =   18
      Top             =   5400
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
      Left            =   4560
      TabIndex        =   17
      Top             =   4920
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
      Left            =   4560
      TabIndex        =   16
      Top             =   4440
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
      Left            =   4560
      TabIndex        =   15
      Top             =   3960
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
      Left            =   4560
      TabIndex        =   14
      Top             =   3480
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
      Height          =   735
      Index           =   1
      Left            =   4560
      TabIndex        =   13
      Top             =   2760
      Width           =   4695
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
      Left            =   4560
      TabIndex        =   12
      Top             =   2400
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
      Left            =   2520
      TabIndex        =   11
      Top             =   5880
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
      Left            =   2640
      TabIndex        =   10
      Top             =   5400
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
      Left            =   2715
      TabIndex        =   9
      Top             =   4920
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
      Left            =   2715
      TabIndex        =   8
      Top             =   4470
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
      Left            =   2715
      TabIndex        =   7
      Top             =   3960
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
      Left            =   2715
      TabIndex        =   6
      Top             =   3480
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
      Left            =   2715
      TabIndex        =   5
      Top             =   2880
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
      Left            =   2715
      TabIndex        =   4
      Top             =   2430
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Status:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   8
      Left            =   6120
      TabIndex        =   3
      Top             =   6585
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
      Left            =   6720
      TabIndex        =   2
      Top             =   6480
      Width           =   2655
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






'Right, now that we have declared our variables, we can start the actual
'coding of the project
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
    lstdata.AddItem rs("BookTitle")
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
Set rs = db.OpenRecordset("Select * from BookDatabase where BookTitle = '" & Trim(lstdata.list(lstdata.ListIndex)) & "'")
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
    Set rs = db.OpenRecordset("SELECT * FROM BookDatabase WHERE BookTitle LIKE '" & txtSearch.Text & "'" & "& '*'")
End If
'then we need to re-populate the list box with only those surnames beginning with
'the letters in the search field
list

End Sub



