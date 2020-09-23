VERSION 5.00
Begin VB.Form Alog 
   BackColor       =   &H00800000&
   Caption         =   "Administrator Login"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4920
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "*"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmnd2 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton login 
      Caption         =   "LOGIN"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "Password :"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "Username :"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Alog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''THIS IS THE ADMINISTRATOR LOGIN FORM'''''''''''''''''''''''''

Private Sub login_Click()
Dim mark As Boolean
With MLSDB.rsAdministrator
If MLSDB.rsAdministrator.RecordCount = 0 Then
    AdminReg.Show
    Unload Me
    mark = False
    While (Not mark) And (Not .EOF)
        If (MLSDB.rsAdministrator.Fields("Username").Value = Text1.Text) And (MLSDB.rsAdministrator.Fields("Password").Value = Text2.Text) Then
            mark = True
            Form2.Show
            Unload Me
        Else
            .MoveNext
            
        End If
    Wend
        
Else

    mark = False
    While (Not mark) And (Not .EOF)
        If (MLSDB.rsAdministrator.Fields("Username").Value = Text1.Text) And (MLSDB.rsAdministrator.Fields("Password").Value = Text2.Text) Then
            mark = True
            Form2.Show
            Unload Me
        Else
            .MoveNext
            
        End If
    Wend
End If
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
            Alog.Show
            Exit Sub
    End If
End Sub

Private Sub cmnd2_Click()
    MLSDB.rsCirculation.Close
    MLSDB.rsBookDbase.Close
    MLSDB.rsPatronDB.Close
    MLSDB.rsAdministrator.Close
    Form0.Show
    Unload Me
End Sub


Private Sub txtText1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 95) Then
Else
KeyAscii = 0
End If
End Sub

Private Sub txtText2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 95) Then
Else
KeyAscii = 0
End If
End Sub



