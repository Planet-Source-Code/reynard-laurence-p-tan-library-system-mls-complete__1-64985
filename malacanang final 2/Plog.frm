VERSION 5.00
Begin VB.Form Plog 
   BackColor       =   &H00800000&
   Caption         =   "Patron Login"
   ClientHeight    =   2580
   ClientLeft      =   7170
   ClientTop       =   4905
   ClientWidth     =   4920
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   2880
      Picture         =   "Plog.frx":0000
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "*"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   915
      Width           =   2055
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   0
      Top             =   285
      Width           =   2040
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Password:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   -180
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Username:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   -150
      TabIndex        =   4
      Top             =   330
      Width           =   1815
   End
End
Attribute VB_Name = "Plog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''THIS IS THE PATRON LOGIN FORM''''''''''''''''''''''''''''

Private Sub Command1_Click()    'Login Button
Dim mark As Boolean

With MLSDB.rsPatronDB
    .MoveFirst
    mark = False
    While (Not mark) And (Not .EOF)
        If (MLSDB.rsPatronDB.Fields("Username").Value = txtUsername.Text) And (MLSDB.rsPatronDB.Fields("Password").Value = txtPassword.Text) Then
            mark = True
            Form3.Show
            Unload Me
        Else
            .MoveNext
            
            End If
       
    Wend
End With
    If mark = False Then
            MsgBox "Invalid username or password. Enter a valid username or password or select a different account type.", vbCritical
            MLSDB.rsCirculation.Close
            MLSDB.rsBookDbase.Close
            MLSDB.rsPatronDB.Close
            MLSDB.rsAdministrator.Close
            Unload Me
            MLSDB.rsCirculation.Open
            MLSDB.rsBookDbase.Open
            MLSDB.rsPatronDB.Open
            MLSDB.rsAdministrator.Open
            Plog.Show
            Exit Sub
    End If
End Sub
Private Sub Command2_Click()    'Exit Button
MLSDB.rsCirculation.Close
MLSDB.rsBookDbase.Close
MLSDB.rsPatronDB.Close
MLSDB.rsAdministrator.Close
Form0.Show
Unload Me
End Sub
Private Sub Form_Load()
With MLSDB.rsPatronDB
If MLSDB.rsPatronDB.RecordCount = 0 Then
    MsgBox "Please See an Administrator to register", vbOKOnly
    Form0.Show
    Unload Me
End If
End With
End Sub

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



