VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H00800000&
   Caption         =   "Add Book"
   ClientHeight    =   6585
   ClientLeft      =   6210
   ClientTop       =   3360
   ClientWidth     =   6465
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   6465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      DownPicture     =   "Form13.frx":0000
      Height          =   735
      Left            =   5640
      Picture         =   "Form13.frx":09B5
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtResponsibilityCenter 
      Height          =   285
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   6
      Top             =   3840
      Width           =   3375
   End
   Begin VB.TextBox text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      DownPicture     =   "Form13.frx":1370
      Height          =   855
      Left            =   3840
      Picture         =   "Form13.frx":1792
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      DownPicture     =   "Form13.frx":18B0
      Height          =   855
      Left            =   1320
      Picture         =   "Form13.frx":1D8A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txtYearofAcquisition 
      Height          =   285
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   5
      Top             =   3390
      Width           =   1560
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   4
      Top             =   2895
      Width           =   1560
   End
   Begin VB.TextBox txtBarcodeNumber 
      Height          =   285
      Left            =   2160
      MaxLength       =   11
      TabIndex        =   3
      Top             =   2385
      Width           =   1560
   End
   Begin VB.TextBox txtCallNumber 
      Height          =   285
      Left            =   2160
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1890
      Width           =   1575
   End
   Begin VB.TextBox txtAuthor 
      Height          =   285
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1395
      Width           =   2175
   End
   Begin VB.TextBox txtBookTitle 
      Height          =   285
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   0
      Top             =   885
      Width           =   3135
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
      TabIndex        =   18
      Top             =   4365
      Width           =   495
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Responsibility Center:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   17
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Year of Acquisition:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   270
      TabIndex        =   16
      Top             =   3435
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
      Left            =   270
      TabIndex        =   15
      Top             =   2940
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Book ID:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   4
      Left            =   1455
      TabIndex        =   14
      Top             =   2430
      Width           =   630
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Call Number:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   270
      TabIndex        =   13
      Top             =   1935
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
      Left            =   270
      TabIndex        =   12
      Top             =   1440
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
      Left            =   270
      TabIndex        =   11
      Top             =   930
      Width           =   1815
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''THIS IS THE ADD BOOK FORM''''''''''''''''''''''''''''''''''''

Private Sub Command1_Click()    'Add Button
Dim found As Boolean
Dim found2 As Boolean
Dim found3 As Boolean
With MLSDB.rsBookDbase

If MLSDB.rsBookDbase.RecordCount = 0 Then
       found = False
       found2 = False
       found3 = False
       
With MLSDB.rsBookDbase

    found3 = False
    found2 = False
    found = False
    While (Not .EOF) And (Not found2)
        If MLSDB.rsBookDbase.Fields("BookTitle").Value = txtBookTitle.Text Then
            found2 = True
        Else
            .MoveNext
        End If
    Wend

If found2 = True Then
    MsgBox "Book Title already exist!", vbCritical
End If

    While (Not .EOF) And (Not found)
        If MLSDB.rsBookDbase.Fields("CallNumber").Value = txtCallNumber.Text Then
            found = True
        Else
            .MoveNext
        End If
    Wend
    

If found = True Then
    MsgBox "Call Number already exist!", vbCritical
End If

    While (Not .EOF) And (Not found3)
        If MLSDB.rsBookDbase.Fields("BookID").Value = txtBarcodeNumber.Text Then
            found3 = True
        Else
            .MoveNext
        End If
    Wend
    
End With
If found3 = True Then
    MsgBox "BookID already exist!", vbCritical
    
ElseIf (found = False) And (found2 = False) And (found3 = False) Then

    If txtBookTitle.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtBookTitle.SetFocus
    ElseIf txtAuthor.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtAuthor.SetFocus
    ElseIf txtCallNumber.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtCallNumber.SetFocus
    ElseIf txtBarcodeNumber.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtBarcodeNumber.SetFocus
    ElseIf txtPrice.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtPrice.SetFocus
    ElseIf txtYearofAcquisition.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtYearofAcquisition.SetFocus
    ElseIf txtResponsibilityCenter.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtResponsibilityCenter.SetFocus
    Else
        With MLSDB.rsBookDbase
    
        .AddNew
        .Fields("Book ID") = Trim(txtBookID.Text)
        .Fields("BookTitle") = Trim(txtBookTitle.Text)
        .Fields("Book Author") = Trim(txtAuthor.Text)
        .Fields("CallNumber") = Trim(txtCallNumber.Text)
        .Fields("BookID") = Trim(txtBarcodeNumber.Text)
        .Fields("Price") = Trim(txtPrice.Text)
        .Fields("Year of Acquisition") = Trim(txtYearofAcquisition.Text)
        .Fields("Responsibility Center") = Trim(txtResponsibilityCenter.Text)
        .Fields("In / Out") = Trim(Text1(1).Text)
        .Update
        End With
    MsgBox "One book record has been Successfully added!", vbInformation
    Form5.Show
    Unload Me
End If
End If

Else        'Record Count is not 0
With MLSDB.rsBookDbase
   .MoveFirst
    found3 = False
    found2 = False
    found = False
    While (Not .EOF) And (Not found2)
        If MLSDB.rsBookDbase.Fields("BookTitle").Value = txtBookTitle.Text Then
            found2 = True
        Else
            .MoveNext
        End If
    Wend

If found2 = True Then
    MsgBox "Book Title already exist!", vbCritical
End If
.MoveFirst
    While (Not .EOF) And (Not found)
        If MLSDB.rsBookDbase.Fields("CallNumber").Value = txtCallNumber.Text Then
            found = True
        Else
            .MoveNext
        End If
    Wend
    

If found = True Then
    MsgBox "Call Number already exist!", vbCritical
End If
    .MoveFirst
    While (Not .EOF) And (Not found3)
        If MLSDB.rsBookDbase.Fields("BookID").Value = txtBarcodeNumber.Text Then
            found3 = True
        Else
            .MoveNext
        End If
    Wend
    
End With
If found3 = True Then
    MsgBox "BookID already exist!", vbCritical
    
ElseIf (found = False) And (found2 = False) And (found3 = False) Then
    If txtBookTitle.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtBookTitle.SetFocus
    ElseIf txtAuthor.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtAuthor.SetFocus
    ElseIf txtCallNumber.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtCallNumber.SetFocus
    ElseIf txtBarcodeNumber.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtBarcodeNumber.SetFocus
    ElseIf txtPrice.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtPrice.SetFocus
    ElseIf txtYearofAcquisition.Text = "" Then
         MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtYearofAcquisition.SetFocus
    ElseIf txtResponsibilityCenter.Text = "" Then
         MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtResponsibilityCenter.SetFocus
    Else
        With MLSDB.rsBookDbase
    
        .AddNew
        .Fields("BookTitle") = Trim(txtBookTitle.Text)
        .Fields("Author") = Trim(txtAuthor.Text)
        .Fields("CallNumber") = Trim(txtCallNumber.Text)
        .Fields("BookID") = Trim(txtBarcodeNumber.Text)
        .Fields("Price") = Trim(txtPrice.Text)
        .Fields("Year of Acquisition") = Trim(txtYearofAcquisition.Text)
        .Fields("Responsibility Center") = Trim(txtResponsibilityCenter.Text)
        .Fields("In / Out") = Trim(Text1(1).Text)
        .Update
        End With
        
    MsgBox "One book record has been Successfully added!", vbInformation
    Form5.Show
    Unload Me
End If
End If
End If
End With
End Sub
Private Sub Command2_Click()    'Exit Button
Form5.Show
Unload Me
End Sub

Private Sub Command6_Click()
helpaddbook.Show
End Sub

Private Sub Form_Load()
Text1(1).Text = "in"
End Sub



Private Sub txtBarcodeNumber_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
Else
KeyAscii = 0
End If
End Sub
Private Sub txtBookID_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 45) Then
Else
KeyAscii = 0
End If
End Sub

Private Sub txtBookTitle_KeyPress(KeyAscii As Integer)
If (KeyAscii = 39) Then
KeyAscii = 0
Else
End If
End Sub

Private Sub txtCallNumber_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 46) Then
Else
KeyAscii = 0
'Or KeyAscii = 45
End If
End Sub
Private Sub txtPrice_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 45) Then
Else
KeyAscii = 0
End If
End Sub
Private Sub txtYearofAcquisition_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 45) Then
Else
KeyAscii = 0
End If
End Sub

Private Sub txtAuthor_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 32) Then
Else
KeyAscii = 0
End If
End Sub

