VERSION 5.00
Begin VB.Form EditAdmin 
   BackColor       =   &H00800000&
   Caption         =   "Edit Administrator"
   ClientHeight    =   6660
   ClientLeft      =   6585
   ClientTop       =   3165
   ClientWidth     =   6000
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text10 
      Height          =   855
      Left            =   5160
      TabIndex        =   31
      Text            =   "Text7"
      Top             =   4560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Default"
      Height          =   615
      Left            =   2520
      TabIndex        =   30
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text6 
      DataField       =   "Address"
      DataMember      =   "Administrator"
      Height          =   285
      Left            =   0
      MaxLength       =   50
      TabIndex        =   29
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text5 
      DataField       =   "Password"
      DataMember      =   "Administrator"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   0
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   28
      Top             =   6420
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text4 
      DataField       =   "Username"
      DataMember      =   "Administrator"
      Height          =   285
      Left            =   0
      MaxLength       =   50
      TabIndex        =   27
      Top             =   5925
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text3 
      DataField       =   "Email Address"
      DataMember      =   "Administrator"
      Height          =   285
      Left            =   0
      MaxLength       =   50
      TabIndex        =   26
      Top             =   5415
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      DataField       =   "Telephone Number"
      DataMember      =   "Administrator"
      Height          =   285
      Left            =   0
      MaxLength       =   7
      TabIndex        =   25
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      MaxLength       =   50
      TabIndex        =   24
      Top             =   3915
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox txtAdministratorName 
      DataField       =   "Administrator Name"
      DataMember      =   "Administrator"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1275
      Width           =   3210
   End
   Begin VB.TextBox txtTelephoneNumber 
      DataField       =   "Telephone Number"
      DataMember      =   "Administrator"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   2040
      MaxLength       =   7
      TabIndex        =   2
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtEmailAddress 
      DataField       =   "Email Address"
      DataMember      =   "Administrator"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2775
      Width           =   2535
   End
   Begin VB.TextBox txtUsername 
      DataField       =   "Username"
      DataMember      =   "Administrator"
      DataSource      =   "MLSDB"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   4
      Top             =   3285
      Width           =   3255
   End
   Begin VB.TextBox txtPassword 
      DataField       =   "Password"
      DataMember      =   "Administrator"
      DataSource      =   "MLSDB"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3780
      Width           =   3255
   End
   Begin VB.TextBox txtAddress 
      DataField       =   "Address"
      DataMember      =   "Administrator"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1800
      Width           =   3855
   End
   Begin VB.CommandButton Command6 
      DownPicture     =   "EditAdmin.frx":0000
      Height          =   735
      Left            =   5160
      Picture         =   "EditAdmin.frx":09B5
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      DownPicture     =   "EditAdmin.frx":1370
      Height          =   855
      Left            =   3600
      Picture         =   "EditAdmin.frx":1792
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      DownPicture     =   "EditAdmin.frx":18B0
      Height          =   855
      Left            =   960
      Picture         =   "EditAdmin.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   375
      Left            =   1560
      TabIndex        =   23
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   375
      Left            =   2160
      TabIndex        =   22
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   375
      Left            =   2760
      TabIndex        =   21
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   375
      Left            =   3360
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   3960
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
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
      Left            =   2280
      TabIndex        =   18
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Administrator ID:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Administrator Name*:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   1
      Left            =   495
      TabIndex        =   16
      Top             =   1320
      Width           =   1470
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
      TabIndex        =   15
      Top             =   1830
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
      Left            =   150
      TabIndex        =   14
      Top             =   2325
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Email Address:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   150
      TabIndex        =   13
      Top             =   2820
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Username*:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   5
      Left            =   1110
      TabIndex        =   12
      Top             =   3330
      Width           =   825
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Password*:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   6
      Left            =   1170
      TabIndex        =   11
      Top             =   3825
      Width           =   795
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "Administrator ID"
      DataMember      =   "Administrator"
      DataSource      =   "MLSDB"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "EditAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''THIS IS THE EDIT ADMINISTRATOR FORM'''''''''''''''''''''''''

Private Sub Command1_Click() 'Save Button
Dim x As Long
'Dim y As Long
With MLSDB.rsAdministrator
    If txtAdministratorName.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtAdministratorName.SetFocus
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
        'txtTelephoneNumber.SetFocus
    'ElseIf txtEmailAddress.Text = "" Then
        'y = 0
        'txtEmailAddress.Text = "n/a"
        'MsgBox "Kindly fill up all the Required  fields", vbCritical
        'txtEmailAddress.SetFocus
    ElseIf txtUsername.Text = "" Then
         MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtUsername.SetFocus
     ElseIf txtPassword.Text = "" Then
         MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtPassword.SetFocus
    Else
    
        With MLSDB.rsAdministrator
        .Fields("Administrator Name").Value = txtAdministratorName.Text
        .Fields("Address").Value = txtAddress.Text
        .Fields("Telephone Number").Value = txtTelephoneNumber.Text
        .Fields("Email Address").Value = txtEmailAddress.Text
        .Fields("Username").Value = txtUsername.Text
        .Fields("Password").Value = txtPassword.Text
        .Update
               
        End With
    

     MsgBox "One record has been edited!", vbInformation
     Form6.Show
     Unload Me
     
End If
End With
End Sub
Private Sub Command2_Click()
txtAdministratorName.Text = Label2.Caption
txtAddress.Text = Label4.Caption
txtTelephoneNumber.Text = Label5.Caption
txtEmailAddress.Text = Label6.Caption
txtUsername.Text = Label7.Caption
txtPassword.Text = Label8.Caption
With MLSDB.rsAdministrator
        .Fields("Administrator Name").Value = txtAdministratorName.Text
        .Fields("Address").Value = txtAddress.Text
        .Fields("Telephone Number").Value = txtTelephoneNumber.Text
        .Fields("Email Address").Value = txtEmailAddress.Text
        .Fields("Username").Value = txtUsername.Text
        .Fields("Password").Value = txtPassword.Text
        .Update
               
        End With
Form6.Show
Unload Me
End Sub

'Private Sub Command3_Click()
'Form6.Show
'Unload Me
'End Sub

Private Sub Command6_Click()
helpeditadmin.Show
End Sub

Private Sub Form_Load()
Label2.Caption = txtAdministratorName.Text
Label4.Caption = txtAddress.Text
Label5.Caption = txtTelephoneNumber.Text
Label6.Caption = txtEmailAddress.Text
Label7.Caption = txtUsername.Text
Label8.Caption = txtPassword.Text

End Sub
Private Sub txtAdministratorName_KeyPress(KeyAscii As Integer)
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

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 95) Then
Else
KeyAscii = 0
End If
End Sub



