VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00800000&
   Caption         =   "Edit Book"
   ClientHeight    =   6120
   ClientLeft      =   6210
   ClientTop       =   3360
   ClientWidth     =   6630
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      DownPicture     =   "Form12.frx":0000
      Height          =   735
      Left            =   5760
      Picture         =   "Form12.frx":09B5
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      DownPicture     =   "Form12.frx":1370
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
      Left            =   3480
      Picture         =   "Form12.frx":1792
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      DownPicture     =   "Form12.frx":18B0
      Height          =   855
      Left            =   960
      Picture         =   "Form12.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txtResponsibilityCenter 
      DataField       =   "Responsibility Center"
      DataMember      =   "BookDbase"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   6
      Top             =   4035
      Width           =   2895
   End
   Begin VB.TextBox txtYearofAcquisition 
      DataField       =   "Year of Acquisition"
      DataMember      =   "BookDbase"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   5
      Top             =   3540
      Width           =   1545
   End
   Begin VB.TextBox txtPrice 
      DataField       =   "Price"
      DataMember      =   "BookDbase"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   4
      Top             =   3045
      Width           =   1515
   End
   Begin VB.TextBox txtBarcodeNumber 
      DataField       =   "BookID"
      DataMember      =   "BookDbase"
      DataSource      =   "MLSDB"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      MaxLength       =   11
      TabIndex        =   3
      Top             =   2535
      Width           =   1515
   End
   Begin VB.TextBox txtCallNumber 
      DataField       =   "CallNumber"
      DataMember      =   "BookDbase"
      DataSource      =   "MLSDB"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      MaxLength       =   7
      TabIndex        =   2
      Top             =   2040
      Width           =   1500
   End
   Begin VB.TextBox txtAuthor 
      DataField       =   "Author"
      DataMember      =   "BookDbase"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1545
      Width           =   3495
   End
   Begin VB.TextBox txtBookTitle 
      DataField       =   "BookTitle"
      DataMember      =   "BookDbase"
      DataSource      =   "MLSDB"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1035
      Width           =   3975
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   375
      Left            =   1560
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   375
      Left            =   2160
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   375
      Left            =   2760
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   3360
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   3960
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   4560
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   5160
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
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
      Left            =   180
      TabIndex        =   16
      Top             =   4080
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
      Left            =   180
      TabIndex        =   15
      Top             =   3585
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
      Left            =   180
      TabIndex        =   14
      Top             =   3090
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
      Left            =   1365
      TabIndex        =   13
      Top             =   2580
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
      Left            =   180
      TabIndex        =   12
      Top             =   2085
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
      Left            =   180
      TabIndex        =   11
      Top             =   1590
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
      Left            =   180
      TabIndex        =   10
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''THIS IS THE EDIT BOOK FORM'''''''''''''''''''''''''''''''''''

Private Sub Command1_Click()    'Save Button
With MLSDB.rsBookDbase
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
    ElseIf txtYearofAcquisition.Text = "" Then
         MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtYearofAcquisition.SetFocus
    ElseIf txtResponsibilityCenter.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtResponsibilityCenter.SetFocus
    ElseIf txtPrice.Text = "" Then
        MsgBox "Kindly fill up all the Required  fields", vbCritical
        txtResponsibilityCenter.SetFocus
    Else
    
        With MLSDB.rsBookDbase
        .Fields("BookTitle").Value = Trim(txtBookTitle.Text)
        .Fields("Author").Value = Trim(txtAuthor.Text)
        .Fields("CallNumber").Value = Trim(txtCallNumber.Text)
        .Fields("BookID").Value = Trim(txtBarcodeNumber.Text)
        .Fields("Year of Acquisition").Value = Trim(txtYearofAcquisition.Text)
        .Fields("Responsibility Center").Value = Trim(txtResponsibilityCenter.Text)
        .Fields("Price").Value = Trim(txtPrice.Text)
        .Update
               
        End With
    

     MsgBox "One record has been edited!", vbInformation
     Form5.Show
     Unload Me
     
End If
End With
End Sub

Private Sub Command2_Click()
    txtBookTitle.Text = Label1.Caption
    txtAuthor.Text = Label2.Caption
    txtCallNumber.Text = Label3.Caption
    txtBarcodeNumber.Text = Label4.Caption
    txtYearofAcquisition.Text = Label5.Caption
    txtResponsibilityCenter.Text = Label6.Caption
    txtPrice.Text = Label7.Caption
    
    With MLSDB.rsBookDbase
        .Fields("BookTitle").Value = Trim(txtBookTitle.Text)
        .Fields("Author").Value = Trim(txtAuthor.Text)
        .Fields("CallNumber").Value = Trim(txtCallNumber.Text)
        .Fields("BookID").Value = Trim(txtBarcodeNumber.Text)
        .Fields("Year of Acquisition").Value = Trim(txtYearofAcquisition.Text)
        .Fields("Responsibility Center").Value = Trim(txtResponsibilityCenter.Text)
        .Fields("Price").Value = Trim(txtPrice.Text)
        .Update
               
    End With

    Form5.Show
    Unload Me
End Sub

Private Sub Command6_Click()
helpeditbook.Show
End Sub

Private Sub txtBarcodeNumber_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 45) Then
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
Private Sub txtBookTitle_KeyPress(KeyAscii As Integer)
If (KeyAscii = 39) Then
KeyAscii = 0
Else
End If
End Sub

Private Sub Form_Load()
    Label1.Caption = txtBookTitle.Text
    Label2.Caption = txtAuthor.Text
    Label3.Caption = txtCallNumber.Text
    Label4.Caption = txtBarcodeNumber.Text
    Label5.Caption = txtYearofAcquisition.Text
    Label6.Caption = txtResponsibilityCenter.Text
    Label7.Caption = txtPrice.Text
End Sub

