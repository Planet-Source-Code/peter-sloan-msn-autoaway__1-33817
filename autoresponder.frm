VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRespond 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MSN AutoAway"
   ClientHeight    =   1995
   ClientLeft      =   2775
   ClientTop       =   4515
   ClientWidth     =   7095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "autoresponder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "autoresponder.frx":09CA
   ScaleHeight     =   1995
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   6240
      TabIndex        =   6
      Top             =   840
      Width           =   495
   End
   Begin MSComctlLib.ProgressBar gauUsers 
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   1080
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer tmrMsg 
      Interval        =   1000
      Left            =   6480
      Top             =   120
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Enable"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Text            =   "I'm not here at the moment. So please leave a message."
      Top             =   600
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Users Online"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   6855
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Type in what you want as your automated message."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Tools 
      Caption         =   "&Tools"
      Begin VB.Menu Find 
         Caption         =   "&Find"
      End
      Begin VB.Menu LogOffLine 
         Caption         =   "&LogIn as Offline"
      End
   End
End
Attribute VB_Name = "frmRespond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Autoresponder code - By Doggie
' Freely use this code
Dim WithEvents respond As MsgrObject
Attribute respond.VB_VarHelpID = -1
Dim flag As Integer
Dim MsnApp As IMessengerApp
Dim msnObj As IMsgrObject
Public MsnApi As MessengerAPI.IMessengerConversationWnd
Private Function AutoLogin()
    Dim Ans As Integer
    Ans = MsgBox("You are currently not logged in." & vbCrLf & "Do you want to login?", vbYesNo, "Critical")
    If Ans = 6 Then
        MsnApp.AutoLogon
        Do
            If msnObj.LocalState = MSTATE_ONLINE Then Exit Do
        Loop
    Else
        Unload Me
    End If
End Function

Private Sub Check2_Click()
If Check2.Value = 1 Then
msnObj.LocalState = MSTATE_INVISIBLE
Else
msnObj.LocalState = MSTATE_ONLINE
End If
End Sub

Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub Find_Click()
    Load frmFind
    frmFind.Show
End Sub

Private Sub Form_Load()
On Error GoTo errz
MsgBox App.Path
Set respond = New MsgrObject 'declaring the messenger object
Set MsnApp = CreateObject("messenger.messengerapp") 'declaring the messenger object
Set msnObj = CreateObject("messenger.msgrobject")
Set users = msnObj.List(MLIST_ALLOW)
If msnObj.LocalState = 1 Then AutoLogin
 'set progress up
If users.Count > 0 Then
    gauUsers.Max = users.Count
    gauUsers.Min = 0
Else
    gauUsers.Enabled = False
End If
 For i = 0 To users.Count - 1
                If users.Item(i).State <> MSTATE_OFFLINE Then
                gauUsers.Value = gauUsers.Value + 1
                End If
 Next i
Set users = Nothing
errz:
If Err.Number = 380 Then
    AutoLogin
        
        Err.Clear
        Resume
    

ElseIf Err.Number > 0 Then
    MsgBox Err.Number & vbCrLf & Err.Description
    Unload Me
End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Set respond = Nothing ' after use it will delete the declared method
End Sub

Private Sub lblMsg_Click()
On Error GoTo errz
    flag = flag + 1
    Select Case flag
        Case 1
            lblMsg.Caption = "You currently have " & MsnApp.IMWindows.Count & " msn windows open"
        Case 2
            lblMsg.Caption = "You currently have (" & msnObj.UnreadEmail(MFOLDER_INBOX) & ") unread messages."
        Case 3
            
            lblMsg.Caption = "Today is " & Date
        Case 4
            flag = 0
            Set users = msnObj.List(MLIST_ALLOW)
            Dim UsersOnline As Integer
            UsersOnline = 0
            'Loop trough the users and debug.print al the online buddies
            For i = 0 To users.Count - 1
                If users.Item(i).State <> MSTATE_OFFLINE Then
                UsersOnline = UsersOnline + 1
                End If
            Next i


            lblMsg.Caption = "You currently have (" & UsersOnline & " of " & users.Count & ") buddies online"
        End Select
errz:
If msnObj.LocalState = 1 Then AutoLogin
End Sub

Private Sub LogOffLine_Click()
On Error GoTo errz
ENTER = vbCrLf
If msnObj.LocalState = MSTATE_OFFLINE Then
    MsnApp.AutoLogon
    msnObj.LocalState = MSTATE_INVISIBLE
End If
End Sub

Private Sub respond_OnLogoff()
frmstop.Show ' detects if u log off and will shutdown the program
msnObj.UnreadEmail
End Sub

Private Sub respond_OnTextReceived(ByVal pIMSession As Messenger.IMsgrIMSession, ByVal pSourceUser As Messenger.IMsgrUser, ByVal bstrMsgHeader As String, ByVal bstrMsgText As String, pfEnableDefault As Boolean)
If Check1.Value = 1 Then
    pSourceUser.SendText bstrMsgHeader, Text1.Text, MMSGTYPE_ALL_RESULTS
End If
End Sub

Private Sub tmrMsg_Timer()
    lblMsg_Click
End Sub
