VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   6705
   ClientLeft      =   3210
   ClientTop       =   2490
   ClientWidth     =   9390
   LinkTopic       =   "Form7"
   ScaleHeight     =   6705
   ScaleWidth      =   9390
   Begin VB.CommandButton Command4 
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   7800
      TabIndex        =   19
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "NEXT>>"
      Height          =   495
      Left            =   6360
      TabIndex        =   18
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<<PREV"
      Height          =   495
      Left            =   4920
      TabIndex        =   17
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox txtResponsibilityCenter 
      DataField       =   "Responsibility Center"
      DataMember      =   "BookDbase"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   5145
      TabIndex        =   16
      Top             =   3735
      Width           =   3375
   End
   Begin VB.TextBox txtYearofAcquisition 
      DataField       =   "Year of Acquisition"
      DataMember      =   "BookDbase"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   5145
      TabIndex        =   14
      Top             =   3360
      Width           =   1320
   End
   Begin VB.TextBox txtPrice 
      DataField       =   "Price"
      DataMember      =   "BookDbase"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   5145
      TabIndex        =   12
      Top             =   2985
      Width           =   1320
   End
   Begin VB.TextBox txtBarcodeNumber 
      DataField       =   "Barcode Number"
      DataMember      =   "BookDbase"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   5145
      TabIndex        =   10
      Top             =   2595
      Width           =   1320
   End
   Begin VB.TextBox txtCallNumber 
      DataField       =   "Call Number"
      DataMember      =   "BookDbase"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   5145
      TabIndex        =   8
      Top             =   2220
      Width           =   3375
   End
   Begin VB.TextBox txtAuthor 
      DataField       =   "Author"
      DataMember      =   "BookDbase"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   5145
      TabIndex        =   6
      Top             =   1845
      Width           =   3375
   End
   Begin VB.TextBox txtBookTitle 
      DataField       =   "Book Title"
      DataMember      =   "BookDbase"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   5145
      TabIndex        =   4
      Top             =   1455
      Width           =   3375
   End
   Begin VB.TextBox txtBookID 
      DataField       =   "Book ID"
      DataMember      =   "BookDbase"
      DataSource      =   "MLSDB"
      Height          =   285
      Left            =   5145
      TabIndex        =   2
      Top             =   1080
      Width           =   660
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Responsibility Center:"
      Height          =   255
      Index           =   7
      Left            =   3300
      TabIndex        =   15
      Top             =   3780
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Year of Acquisition:"
      Height          =   255
      Index           =   6
      Left            =   3300
      TabIndex        =   13
      Top             =   3405
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Price:"
      Height          =   255
      Index           =   5
      Left            =   3300
      TabIndex        =   11
      Top             =   3030
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Barcode Number:"
      Height          =   255
      Index           =   4
      Left            =   3300
      TabIndex        =   9
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Call Number:"
      Height          =   255
      Index           =   3
      Left            =   3300
      TabIndex        =   7
      Top             =   2265
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Author:"
      Height          =   255
      Index           =   2
      Left            =   3300
      TabIndex        =   5
      Top             =   1890
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Book Title:"
      Height          =   255
      Index           =   1
      Left            =   3300
      TabIndex        =   3
      Top             =   1500
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Book ID:"
      Height          =   255
      Index           =   0
      Left            =   3300
      TabIndex        =   1
      Top             =   1125
      Width           =   1815
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
