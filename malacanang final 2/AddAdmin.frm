VERSION 5.00
Begin VB.Form AddAdmin 
   BackColor       =   &H00800000&
   Caption         =   "Add Administrator"
   ClientHeight    =   6375
   ClientLeft      =   3780
   ClientTop       =   2850
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAdministratorID 
      Enabled         =   0   'False
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   720
      Width           =   660
   End
   Begin VB.TextBox txtTelephoneNumber 
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   2520
      MaxLength       =   7
      TabIndex        =   2
      Top             =   2250
      Width           =   855
   End
   Begin VB.TextBox txtEmailAddress 
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   2520
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2715
      Width           =   2055
   End
   Begin VB.TextBox txtUsername 
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   2520
      MaxLength       =   50
      TabIndex        =   4
      Top             =   3225
      Width           =   3015
   End
   Begin VB.TextBox txtPassword 
      ForeColor       =   &H80000007&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3690
      Width           =   3015
   End
   Begin VB.TextBox txtAdministratorName 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "__,__,__"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   2520
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1170
      Width           =   3015
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   2520
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1680
      Width           =   4095
   End
   Begin VB.CommandButton Command6 
      DownPicture     =   "AddAdmin.frx":0000
      Height          =   735
      Left            =   5880
      Picture         =   "AddAdmin.frx":09B5
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      DownPicture     =   "AddAdmin.frx":1370
      Height          =   855
      Left            =   4080
      Picture         =   "AddAdmin.frx":1792
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      DownPicture     =   "AddAdmin.frx":18B0
      Height          =   855
      Left            =   1320
      Picture         =   "AddAdmin.frx":1D8A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
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
      Left            =   2760
      TabIndex        =   17
      Top             =   4320
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
      Left            =   600
      TabIndex        =   16
      Top             =   765
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
      Left            =   945
      TabIndex        =   15
      Top             =   1170
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
      Left            =   1740
      TabIndex        =   14
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
      Left            =   600
      TabIndex        =   13
      Top             =   2265
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
      Left            =   600
      TabIndex        =   12
      Top             =   2760
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
      Left            =   1590
      TabIndex        =   11
      Top             =   3270
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
      Left            =   1620
      TabIndex        =   10
      Top             =   3765
      Width           =   795
   End
End
Attribute VB_Name = "AddAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''THIS IS THE ADD ADMINISTRATOR FORM''''''''''''''''''''''''''

Private Sub Command1_Click()    'Add Button
Dim found As Boolean
Dim found2 As Boolean
Dim x As Long
x = 0

With MLSDB.rsAdministrator

If MLSDB.rsAdministrator.RecordCount = 0 Then
   
With MLSDB.rsAdministrator

 '   found2 = False
   ' found = False
  '  While (Not .EOF) And (Not found2)
'        If MLSDB.rsAdministrator.Fields("Administrator Name").Value = txtAdministratorName.Text Then
'            found2 = True
'        Else
'            .MoveNext
'        End If
 '   Wend

'If found2 = True Then
 '   MsgBox "Administrator Name already exist!", vbCritical
'End If

    While (Not .EOF) And (Not found)
        If MLSDB.rsAdministrator.Fields("Username").Value = txtUsername.Text Then
            found = True
        Else
            .MoveNext
        End If
    Wend
   
End With
If found = True Then
    MsgBox "Username already exist!", vbCritical
    
ElseIf (found = False) Then

    If txtAdministratorID.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtAdministratorID.SetFocus
    ElseIf txtAdministratorName.Text = "" Then
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
    'ElseIf txtTelephoneNumber.Text = "" Then
    '    MsgBox "Kindly fill up all the Required  fields", vbCritical
    '    txtTelephoneNumber.SetFocus
    'ElseIf txtEmailAddress.Text = "" Then
    '    MsgBox "Kindly fill up all the Required  fields", vbCritical
    '    txtEmailAddress.SetFocus
    ElseIf txtUsername.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtUsername.SetFocus
    ElseIf txtPassword.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtPassword.SetFocus
    Else
        With MLSDB.rsAdministrator
    
        .AddNew
        .Fields("Administrator ID") = Trim(txtAdministratorID.Text)
        .Fields("Administrator Name") = Trim(txtAdministratorName.Text)
        .Fields("Address").Value = Trim(txtAddress.Text)
        .Fields("Telephone Number") = Trim(txtTelephoneNumber.Text)
        .Fields("Email Address") = Trim(txtEmailAddress.Text)
        .Fields("Username") = Trim(txtUsername.Text)
        .Fields("Password") = Trim(txtPassword.Text)
        .Update
        End With
        
    MsgBox "One record has been Successfully added!", vbInformation
    Form6.Show
    Unload Me
End If
End If

Else        'Record Count is not 0
        
With MLSDB.rsAdministrator
'.MoveFirst
 '   found2 = False
 '   found = False
 '   While (Not .EOF) And (Not found2)
 '       If MLSDB.rsAdministrator.Fields("Administrator Name").Value = txtAdministratorName.Text Then
  '          found2 = True
  '      Else
  '          .MoveNext
  '      End If
  '  Wend

'If found2 = True Then
 '   MsgBox "Administrator Name already exist!", vbCritical
'End If

.MoveFirst
    While (Not .EOF) And (Not found)
        If MLSDB.rsAdministrator.Fields("Username").Value = txtUsername.Text Then
            found = True
        Else
            .MoveNext
        End If
    Wend
    
End With
If found = True Then
    MsgBox "Username already exist!", vbCritical
    
ElseIf (found = False) Then

    If txtAdministratorID.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtAdministratorID.SetFocus
    ElseIf txtAdministratorName.Text = "" Then
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
    'ElseIf txtTelephoneNumber.Text = "" Then
    '    MsgBox "Kindly fill up all the Required  fields", vbCritical
    '    txtTelephoneNumber.SetFocus
    'ElseIf txtEmailAddress.Text = "" Then
    '    MsgBox "Kindly fill up all the Required  fields", vbCritical
    '    txtEmailAddress.SetFocus
    ElseIf txtUsername.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtUsername.SetFocus
    ElseIf txtPassword.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtPassword.SetFocus
        
    Else
        With MLSDB.rsAdministrator
    
        .AddNew
        .Fields("Administrator ID") = Trim(txtAdministratorID.Text)
        .Fields("Administrator Name") = Trim(txtAdministratorName.Text)
        .Fields("Address").Value = Trim(txtAddress.Text)
        .Fields("Telephone Number") = Trim(txtTelephoneNumber.Text)
        .Fields("Email Address") = Trim(txtEmailAddress.Text)
        .Fields("Username") = Trim(txtUsername.Text)
        .Fields("Password") = Trim(txtPassword.Text)
        .Update
        End With
        MsgBox "One record has been Successfully added!", vbInformation
        Form6.Show
        Unload Me
End If
End If
End If
End With
End Sub
Private Sub Command2_Click()    'Exit Button
Form6.Show
Unload Me
End Sub

Private Sub Command6_Click()
helpaddadmin.Show
End Sub

Private Sub Form_Load()
Dim x As Long
With MLSDB.rsAdministrator
If MLSDB.rsAdministrator.RecordCount = 0 Then
    Text1.Text = 10001
Else
    .MoveLast
    x = MLSDB.rsAdministrator.Fields("Administrator ID") + 1
    txtAdministratorID.Text = x
End If
End With

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


