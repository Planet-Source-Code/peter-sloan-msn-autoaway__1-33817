VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find MSN User"
   ClientHeight    =   1620
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtState 
      Height          =   285
      Left            =   3120
      TabIndex        =   12
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   840
      TabIndex        =   10
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtCountry 
      Height          =   285
      Left            =   840
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtLastName 
      Height          =   285
      Left            =   3120
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtFirstName 
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "State:"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "City:"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Country:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Last Name:"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "First Name:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "REM LaunchFindContactUI(FirstName, LastName, City,State, Country);"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   1080
      Width           =   3735
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msnObj As IMsgrObject
Option Explicit

Private Sub Form_Load()
    Set msnObj = CreateObject("messenger.msgrobject")
End Sub

Private Sub OKButton_Click()
    Dim x As Variant
    x = msnObj.FindUser(txtFirstName.Text, txtLastName.Text, txtCity.Text, txtState.Text, txtCountry.Text)
    MsgBox x
End Sub
