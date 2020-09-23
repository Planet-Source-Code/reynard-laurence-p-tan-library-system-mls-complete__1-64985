VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00800000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form8"
   ClientHeight    =   8715
   ClientLeft      =   2280
   ClientTop       =   1440
   ClientWidth     =   10665
   DrawMode        =   1  'Blackness
   DrawStyle       =   1  'Dash
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   10665
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command6 
      Caption         =   "ISSUE"
      Height          =   735
      Left            =   4440
      Picture         =   "Form15.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Height          =   735
      Left            =   8880
      Picture         =   "Form15.frx":0157
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   7080
      Picture         =   "Form15.frx":12AD
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   24
      Top             =   3720
      Width           =   5295
   End
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
      Left            =   8880
      Picture         =   "Form15.frx":2229
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
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
      Height          =   735
      Left            =   7800
      Picture         =   "Form15.frx":2347
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
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
      Height          =   735
      Left            =   7800
      Picture         =   "Form15.frx":2483
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   1695
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
      Height          =   735
      Left            =   7800
      Picture         =   "Form15.frx":25EF
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1755
      Left            =   240
      Picture         =   "Form15.frx":26FE
      Top             =   120
      Width           =   1875
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2640
      TabIndex        =   27
      Top             =   8280
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
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
      Left            =   2640
      TabIndex        =   23
      Top             =   7800
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
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
      Left            =   2640
      TabIndex        =   22
      Top             =   7320
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
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
      Left            =   2640
      TabIndex        =   21
      Top             =   6840
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
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
      Left            =   2640
      TabIndex        =   20
      Top             =   6360
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
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
      Left            =   2640
      TabIndex        =   19
      Top             =   5880
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Index           =   2
      Left            =   2640
      TabIndex        =   18
      Top             =   5400
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
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
      Left            =   2640
      TabIndex        =   17
      Top             =   4680
      Width           =   4575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
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
      Left            =   2640
      TabIndex        =   16
      Top             =   4320
      Width           =   3495
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Status:"
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   8
      Left            =   1905
      TabIndex        =   15
      Top             =   8280
      Width           =   630
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Responsibility Center:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   7
      Left            =   600
      TabIndex        =   14
      Top             =   7800
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
      Left            =   720
      TabIndex        =   13
      Top             =   7320
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
      Left            =   795
      TabIndex        =   12
      Top             =   6840
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
      Left            =   795
      TabIndex        =   11
      Top             =   6390
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
      Left            =   795
      TabIndex        =   10
      Top             =   5880
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
      Left            =   795
      TabIndex        =   9
      Top             =   5400
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
      Left            =   795
      TabIndex        =   8
      Top             =   4800
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
      Left            =   795
      TabIndex        =   7
      Top             =   4350
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
      Left            =   840
      TabIndex        =   6
      Top             =   3000
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
      Left            =   840
      TabIndex        =   1
      Top             =   1920
      Width           =   8775
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      X1              =   960
      X2              =   9600
      Y1              =   1920
      Y2              =   1920
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
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   8295
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command4.Enabled = False
Command5.Enabled = False

With MLSDB.rsBookDbase
.MoveFirst
found = False

While (Not .EOF) And (Not found)
 If Text1.Text = MLSDB.rsBookDbase.Fields("Book Title") Then
                found = True
                Label4(0).Caption = .Fields("Book ID")
                Label4(1).Caption = .Fields("Book Title")
                Label4(2).Caption = .Fields("Author")
                Label4(3).Caption = .Fields("Call Number")
                Label4(4).Caption = .Fields("Barcode Number")
                Label4(5).Caption = .Fields("Price")
                Label4(6).Caption = .Fields("Year of Acquisition")
                Label4(7).Caption = .Fields("Responsibility Center")
                Label5.Caption = .Fields("In / Out")
                        
      Else
                .MoveNext
                End If
                
 Wend
 If (Not found) Then
  MsgBox "Invalid Entry", vbCritical
  End If
  
 End With
              
              
              
End Sub

Private Sub Command2_Click()

Command5.Enabled = True
With MLSDB.rsBookDbase
.MoveFirst
found = False

While (Not .EOF) And (Not found)
 If Text1.Text = MLSDB.rsBookDbase.Fields("Author") Then
                found = True
                Label4(0).Caption = .Fields("Book ID")
                Label4(1).Caption = .Fields("Book Title")
                Label4(2).Caption = .Fields("Author")
                Label4(3).Caption = .Fields("Call Number")
                Label4(4).Caption = .Fields("Barcode Number")
                Label4(5).Caption = .Fields("Price")
                Label4(6).Caption = .Fields("Year of Acquisition")
                Label4(7).Caption = .Fields("Responsibility Center")
                Label5.Caption = .Fields("In / Out")
                        
      Else
                .MoveNext
                End If
                
 Wend
 If (Not found) Then
  MsgBox "Invalid Entry", vbCritical
  Command5.Enabled = False
  End If
  
 End With
End Sub

Private Sub Command3_Click()
Command4.Enabled = False
Command5.Enabled = False
With MLSDB.rsBookDbase
.MoveFirst
found = False

While (Not .EOF) And (Not found)
 If Text1.Text = MLSDB.rsBookDbase.Fields("Call Number") Then
                found = True
                Label4(0).Caption = .Fields("Book ID")
                Label4(1).Caption = .Fields("Book Title")
                Label4(2).Caption = .Fields("Author")
                Label4(3).Caption = .Fields("Call Number")
                Label4(4).Caption = .Fields("Barcode Number")
                Label4(5).Caption = .Fields("Price")
                Label4(6).Caption = .Fields("Year of Acquisition")
                Label4(7).Caption = .Fields("Responsibility Center")
                Label5.Caption = .Fields("In / Out")
                        
      Else
                .MoveNext
                End If
                
 Wend
 If (Not found) Then
  MsgBox "Invalid Entry", vbCritical
  End If
  
 End With
End Sub

Private Sub Command4_Click()
Command5.Enabled = True
With MLSDB.rsBookDbase

.MovePrevious


found = False

While (Not .BOF) And (Not found)
 If Text1.Text = MLSDB.rsBookDbase.Fields("Author") Then
                found = True
                Label4(0).Caption = .Fields("Book ID")
                Label4(1).Caption = .Fields("Book Title")
                Label4(2).Caption = .Fields("Author")
                Label4(3).Caption = .Fields("Call Number")
                Label4(4).Caption = .Fields("Barcode Number")
                Label4(5).Caption = .Fields("Price")
                Label4(6).Caption = .Fields("Year of Acquisition")
                Label4(7).Caption = .Fields("Responsibility Center")
                Label5.Caption = .Fields("In / Out")
      Else
                .MovePrevious
                End If
 Wend
If (.BOF) And (Not found) Then
 Command4.Enabled = False
End If
End With
End Sub

Private Sub Command5_Click()
Command4.Enabled = True
With MLSDB.rsBookDbase

.MoveNext


found = False

While (Not .EOF) And (Not found)
 If Text1.Text = MLSDB.rsBookDbase.Fields("Author") Then
                found = True
                Label4(0).Caption = .Fields("Book ID")
                Label4(1).Caption = .Fields("Book Title")
                Label4(2).Caption = .Fields("Author")
                Label4(3).Caption = .Fields("Call Number")
                Label4(4).Caption = .Fields("Barcode Number")
                Label4(5).Caption = .Fields("Price")
                Label4(6).Caption = .Fields("Year of Acquisition")
                Label4(7).Caption = .Fields("Responsibility Center")
                Label5.Caption = .Fields("In / Out")
      Else
                .MoveNext
                End If
 Wend
If (.EOF) And (Not found) Then
 Command5.Enabled = False
End If
End With
End Sub

Private Sub Command6_Click()
Issue.Show
Unload Me
End Sub

Private Sub Exit_Click()
Form2.Show
Unload Me
End Sub


Private Sub Form_Load()
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
End Sub



Private Sub Text1_Change()
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command6.Enabled = True
End Sub
