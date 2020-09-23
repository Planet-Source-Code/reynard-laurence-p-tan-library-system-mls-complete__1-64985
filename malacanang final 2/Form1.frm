VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "End"
      Height          =   255
      Left            =   4200
      TabIndex        =   11
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add New"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtPhone 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtForename 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtSurname 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.ListBox lstdata 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Phone:"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Forename:"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Surname:"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   360
      Width           =   855
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
txtForename.Text = rs("Price")
txtSurname.Text = rs("Author")
txtPhone.Text = rs("CallNumber")
'Now we need to enable some of the command buttons so that we can work on the
'data
cmdEdit.Enabled = True
cmdDelete.Enabled = True

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


Private Sub cmdend_Click()
'The final code we need is the exit button to end the program
'close the database, just so that we can close everything cleanly
db.Close
End
End Sub

