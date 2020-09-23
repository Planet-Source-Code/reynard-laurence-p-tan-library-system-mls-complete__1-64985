VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00800000&
   Caption         =   "Edit Patron"
   ClientHeight    =   7275
   ClientLeft      =   6015
   ClientTop       =   2580
   ClientWidth     =   6780
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPatronNumber 
      DataField       =   "Patron Number"
      DataMember      =   "PatronDB"
      DataSource      =   "MLSDB"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   480
      Width           =   1020
   End
   Begin VB.TextBox txtTelephoneNumber 
      DataField       =   "Telephone Number"
      DataMember      =   "PatronDB"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   2145
      MaxLength       =   7
      TabIndex        =   2
      Top             =   2280
      Width           =   1380
   End
   Begin VB.TextBox txtEmailAddress 
      DataField       =   "Email Address"
      DataMember      =   "PatronDB"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   2145
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2835
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      DataField       =   "Member Name"
      DataMember      =   "PatronDB"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox txtDepartment 
      DataField       =   "Department"
      DataMember      =   "PatronDB"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   4
      Top             =   3420
      Width           =   3375
   End
   Begin VB.TextBox txtUsername 
      DataField       =   "Username"
      DataMember      =   "PatronDB"
      DataSource      =   "MLSDB"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   5
      Top             =   3975
      Width           =   2655
   End
   Begin VB.TextBox txtPassword 
      DataField       =   "Password"
      DataMember      =   "PatronDB"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   2145
      MaxLength       =   50
      TabIndex        =   6
      Top             =   4590
      Width           =   2655
   End
   Begin VB.TextBox txtAddress 
      DataField       =   "Address"
      DataMember      =   "PatronDB"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1680
      Width           =   4455
   End
   Begin VB.CommandButton Command6 
      DownPicture     =   "Form11.frx":0000
      Height          =   735
      Left            =   6000
      Picture         =   "Form11.frx":09B5
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      DownPicture     =   "Form11.frx":1370
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
      Left            =   4080
      Picture         =   "Form11.frx":1792
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      DownPicture     =   "Form11.frx":18B0
      Height          =   855
      Left            =   1560
      Picture         =   "Form11.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   375
      Left            =   3360
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   375
      Left            =   2760
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   375
      Left            =   2160
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   375
      Left            =   1560
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   3960
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   4560
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   5160
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "* Required field"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   19
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Patron Number:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   0
      Left            =   825
      TabIndex        =   18
      Top             =   480
      Width           =   1110
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Patron Name*:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   17
      Top             =   1140
      Width           =   1035
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Address*:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   2
      Left            =   1260
      TabIndex        =   16
      Top             =   1680
      Width           =   675
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Telephone Number:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   15
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Email Address:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   4
      Left            =   960
      TabIndex        =   14
      Top             =   2880
      Width           =   1035
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Department*:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   5
      Left            =   1005
      TabIndex        =   13
      Top             =   3465
      Width           =   930
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Username*:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   6
      Left            =   1110
      TabIndex        =   12
      Top             =   4020
      Width           =   825
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Password*:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   7
      Left            =   1140
      TabIndex        =   11
      Top             =   4635
      Width           =   795
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''THIS IS THE EDIT PATRON FORM''''''''''''''''''''''''''''''''

Private Sub Command1_Click()    'Save Button
Dim x As Long
x = 0

With MLSDB.rsPatronDB
    If Text1.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        Text1.SetFocus
    ElseIf txtAddress.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtAddress.SetFocus
    ElseIf (txtTelephoneNumber.Text = "" And txtEmailAddress.Text = "") Then
        txtTelephoneNumber.Text = x
        txtEmailAddress.Text = "n/a"
        MsgBox "Autofilling Email Address and Telephone Number", vbInformation
    ElseIf txtTelephoneNumber.Text = "" Then
        txtTelephoneNumber.Text = x
        MsgBox "Autofilling Telephone Number", vbInformation
    ElseIf txtEmailAddress.Text = "" Then
        txtEmailAddress.Text = "n/a"
        MsgBox "Autofilling Email Address ", vbInformation
    'ElseIf txtTelephoneNumber.Text = "" Then
    '    MsgBox "Kindly fill up all the Required  fields", vbCritical
    '    txtTelephoneNumber.SetFocus
    'ElseIf txtEmailAddress.Text = "" Then
    '     MsgBox "Kindly fill up all the Required  fields", vbCritical
    '    txtEmailAddress.SetFocus
    ElseIf txtDepartment.Text = "" Then
         MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtDepartment.SetFocus
    ElseIf txtUsername.Text = "" Then
         MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtUsername.SetFocus
    ElseIf txtPassword.Text = "" Then
         MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtPassword.SetFocus
    Else
    
        With MLSDB.rsPatronDB
        .Fields("Member Name").Value = Trim(Text1.Text)
        .Fields("Address") = Trim(txtAddress.Text)
        .Fields("Telephone Number").Value = Trim(txtTelephoneNumber.Text)
        .Fields("Email Address").Value = Trim(txtEmailAddress.Text)
        .Fields("Department").Value = Trim(txtDepartment.Text)
        .Fields("Username").Value = Trim(txtUsername.Text)
        .Fields("Password").Value = Trim(txtPassword.Text)
        
        .Update
               
        End With
    

     MsgBox "One record has been edited!", vbInformation
     Form4.Show
     Unload Me
     
End If
End With
End Sub

Private Sub Command2_Click()
Text1.Text = Label2.Caption
txtAddress.Text = Label3.Caption
txtTelephoneNumber.Text = Label4.Caption
txtEmailAddress.Text = Label5.Caption
txtDepartment.Text = Label6.Caption
txtUsername.Text = Label7.Caption
txtPassword.Text = Label8.Caption
    With MLSDB.rsPatronDB
        .Fields("Member Name").Value = Trim(Text1.Text)
        .Fields("Address") = Trim(txtAddress.Text)
        .Fields("Telephone Number").Value = Trim(txtTelephoneNumber.Text)
        .Fields("Email Address").Value = Trim(txtEmailAddress.Text)
        .Fields("Department").Value = Trim(txtDepartment.Text)
        .Fields("Username").Value = Trim(txtUsername.Text)
        .Fields("Password").Value = Trim(txtPassword.Text)
        
        .Update
               
        End With
    Unload Me
    Form4.Show
End Sub

Private Sub Command6_Click()
editclient.Show
End Sub

Private Sub txtPatronNumber_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
Else
KeyAscii = 0
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 32) Then
Else
KeyAscii = 0
End If
End Sub



Private Sub txtEmailAddress_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 95 Or KeyAscii = 64 Or KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 46) Then
Else
KeyAscii = 0
End If
End Sub


Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 32 Or KeyAscii = 44 Or KeyAscii >= 48 And KeyAscii <= 57) Then
Else
KeyAscii = 0
End If
End Sub

Private Sub txtTelephoneNumber_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
Else
KeyAscii = 0
End If
End Sub
 'Or KeyAscii = 45
Private Sub txtUsername_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 95) Then
Else
KeyAscii = 0
End If
End Sub

Private Sub txtDepartment_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32) Then
Else
KeyAscii = 0
End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 95) Then
Else
KeyAscii = 0
End If
End Sub

Private Sub Form_Load()
Label2.Caption = Text1.Text
Label3.Caption = txtAddress.Text
Label4.Caption = txtTelephoneNumber.Text
Label5.Caption = txtEmailAddress.Text
Label6.Caption = txtDepartment.Text
Label7.Caption = txtUsername.Text
Label8.Caption = txtPassword.Text

End Sub
