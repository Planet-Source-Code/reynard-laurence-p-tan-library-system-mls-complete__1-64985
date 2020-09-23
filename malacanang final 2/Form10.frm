VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00800000&
   Caption         =   "Add Patron"
   ClientHeight    =   7380
   ClientLeft      =   6390
   ClientTop       =   2775
   ClientWidth     =   6270
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPatronNumber 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2055
      TabIndex        =   10
      Top             =   600
      Width           =   1140
   End
   Begin VB.TextBox txtMemberName 
      Height          =   285
      Left            =   2055
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1110
      Width           =   3255
   End
   Begin VB.TextBox txtTelephoneNumber 
      Height          =   285
      Left            =   2055
      MaxLength       =   7
      TabIndex        =   2
      Top             =   2100
      Width           =   1380
   End
   Begin VB.TextBox txtEmailAddress 
      Height          =   285
      Left            =   2055
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2595
      Width           =   3015
   End
   Begin VB.TextBox txtDepartment 
      Height          =   285
      Left            =   2055
      MaxLength       =   50
      TabIndex        =   4
      Top             =   3150
      Width           =   3375
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   2055
      MaxLength       =   50
      TabIndex        =   5
      Top             =   3750
      Width           =   2895
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   2055
      MaxLength       =   50
      TabIndex        =   6
      Top             =   4350
      Width           =   2895
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   2055
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1590
      Width           =   3975
   End
   Begin VB.CommandButton Command6 
      DownPicture     =   "Form10.frx":0000
      Height          =   735
      Left            =   5400
      Picture         =   "Form10.frx":09B5
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      DownPicture     =   "Form10.frx":1370
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
      Left            =   3600
      Picture         =   "Form10.frx":1792
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      DownPicture     =   "Form10.frx":18B0
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
      Left            =   1440
      Picture         =   "Form10.frx":1D8A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5520
      Width           =   1335
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
      Left            =   2520
      TabIndex        =   19
      Top             =   4800
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
      Left            =   840
      TabIndex        =   18
      Top             =   645
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
      Left            =   915
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
      Left            =   1275
      TabIndex        =   16
      Top             =   1650
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
      Left            =   135
      TabIndex        =   15
      Top             =   2145
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
      Left            =   135
      TabIndex        =   14
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Department*:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   5
      Left            =   1020
      TabIndex        =   13
      Top             =   3195
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
      Left            =   1125
      TabIndex        =   12
      Top             =   3750
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
      Top             =   4365
      Width           =   795
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
        ''''''''''''''''''''''''''''THIS IS THE ADD PATRON FORM'''''''''''''''''''''''''''''''''

Private Sub Command1_Click()    'Add Button
Dim found As Boolean
Dim found2 As Boolean
Dim x As Long
x = 0

With MLSDB.rsPatronDB

If MLSDB.rsPatronDB.RecordCount = 0 Then
   
With MLSDB.rsPatronDB

    found2 = False
    found = False
    'While (Not .EOF) And (Not found2)
     '   If MLSDB.rsPatronDB.Fields("Member Name").Value = txtMemberName.Text Then
    '        found2 = True
     '   Else
     '       .MoveNext
     '   End If
   ' Wend

'If found2 = True Then
    'MsgBox "Patron Name already exist!", vbCritical
'End If

    While (Not .EOF) And (Not found)
        If MLSDB.rsPatronDB.Fields("Username").Value = txtUsername.Text Then
            found = True
        Else
            .MoveNext
        End If
    Wend
    
End With
If found = True Then
    MsgBox "Username already exist!", vbCritical
    
ElseIf (found = False) Then

    If txtPatronNumber.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtPatronNumber.SetFocus
    ElseIf txtMemberName.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtMemberName.SetFocus
    ElseIf txtAddress.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtAddress.SetFocus
    'ElseIf (txtTelephoneNumber.Text = "" And txtEmailAddress.Text = "") Then
    '    txtTelephoneNumber.Text = x
    '    txtEmailAddress.Text = "n/a"
    '    MsgBox "Autofilling Email Address and Telephone Number", vbInformation
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
    
    Else
            With MLSDB.rsPatronDB
            
            .AddNew
            .Fields("Patron Number") = Trim(txtPatronNumber.Text)
            .Fields("Member Name") = Trim(txtMemberName.Text)
            .Fields("Address") = Trim(txtAddress.Text)
            .Fields("Telephone Number") = Trim(txtTelephoneNumber.Text)
            .Fields("Email Address") = Trim(txtEmailAddress.Text)
            .Fields("Department") = Trim(txtDepartment.Text)
            .Fields("Username") = Trim(txtUsername.Text)
            .Fields("Password") = Trim(txtPassword.Text)
            .Update
            End With
        MsgBox "One Patron has been Successfully added!", vbInformation
        Form4.Show
        Unload Me
        
End If
End If

Else        'Record Count is 0
        
 With MLSDB.rsPatronDB
'.MoveFirst
  '  found2 = False
  '  found = False
  '  While (Not .EOF) And (Not found2)
    '    If MLSDB.rsPatronDB.Fields("Member Name").Value = txtMemberName.Text Then
     '       found2 = True
     '   Else
     '       .MoveNext
     '   End If
   ' Wend

'If found2 = True Then
 '   MsgBox "Patron Name already exist!", vbCritical
'End If
.MoveFirst
    While (Not .EOF) And (Not found)
        If MLSDB.rsPatronDB.Fields("Username").Value = txtUsername.Text Then
            found = True
        Else
            .MoveNext
        End If
    Wend
    
End With
If found = True Then
    MsgBox "Username already exist!", vbCritical
    
ElseIf (found = False) Then
    If txtPatronNumber.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtPatronNumber.SetFocus
    ElseIf txtMemberName.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtMemberName.SetFocus
    ElseIf txtAddress.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtAddress.SetFocus
        'MsgBox "Autofilling Email Address ", vbInformation
    'ElseIf txtTelephoneNumber.Text = "" Then
    '    MsgBox "Kindly fill up all the Required  fields", vbCritical
    '    txtTelephoneNumber.SetFocus
    'ElseIf txtEmailAddress.Text = "" Then
    '    MsgBox "Kindly fill up all the Required  fields", vbCritical
    '    txtEmailAddress.SetFocus
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
    
        .AddNew
        .Fields("Patron Number") = Trim(txtPatronNumber.Text)
        .Fields("Member Name") = Trim(txtMemberName.Text)
        .Fields("Address") = Trim(txtAddress.Text)
        .Fields("Telephone Number") = Trim(txtTelephoneNumber.Text)
        .Fields("Email Address") = Trim(txtEmailAddress.Text)
        .Fields("Department") = Trim(txtDepartment.Text)
        .Fields("Username") = Trim(txtUsername.Text)
        .Fields("Password") = Trim(txtPassword.Text)
        .Update
        End With
        MsgBox "One record has been Successfully added!", vbInformation
        Form4.Show
        Unload Me
End If
End If
End If
End With
End Sub

Private Sub Command2_Click()    'Exit Button
Form4.Show
Unload Me
End Sub

Private Sub Command6_Click()
    addclient.Show
End Sub

Private Sub Form_Load()
Dim x As Long
With MLSDB.rsPatronDB
If MLSDB.rsPatronDB.RecordCount = 0 Then
    txtPatronNumber.Text = 1000001
Else
    .MoveLast
    x = MLSDB.rsPatronDB.Fields("Patron Number") + 1
    txtPatronNumber.Text = x
End If
End With
End Sub

Private Sub txtPatronNumber_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 45) Then
Else
KeyAscii = 0
End If
End Sub

Private Sub txtMemberName_KeyPress(KeyAscii As Integer)
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


