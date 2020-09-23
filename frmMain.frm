VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Pinger v1.0"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5775
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Reset Me"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H000000FF&
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   1095
   End
   Begin VB.ListBox listDisplay 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   3375
      ItemData        =   "frmMain.frx":0442
      Left            =   240
      List            =   "frmMain.frx":0449
      TabIndex        =   4
      Top             =   240
      Width           =   5295
   End
   Begin VB.CheckBox chkHide 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hide DOS window"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4560
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CommandButton cmdPing 
      BackColor       =   &H0000FF00&
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox txtPing 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   4080
      Width           =   5295
   End
   Begin VB.Label lblPrompt 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Enter IP address to ping, or command:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3840
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdStop_Click()
    Shell ("command.com /c Control-C"), vbHide
    listDisplay.AddItem ""
    listDisplay.AddItem "Ping Stopped"
    listDisplay.AddItem ""
    cmdPing.Enabled = True
    cmdStop.Enabled = False
    chkHide.Enabled = True
    txtPing.Text = ""
    Exit Sub
End Sub

Private Sub Command1_Click()
    cmdPing.Enabled = True
    cmdStop.Enabled = True
    chkHide.Enabled = True
    txtPing.SetFocus
End Sub

Private Sub Form_Load()
    Call help
End Sub

Private Sub help()
    listDisplay.AddItem ""
    listDisplay.AddItem "Commands:"
    listDisplay.AddItem "   /control-c     (Stops Ping)"
    listDisplay.AddItem "   /clear         (Clears the the display)"
    listDisplay.AddItem "   /help          (Displays this text)"
    listDisplay.AddItem ""
    listDisplay.AddItem "To ping a address just enter the IP (Internet Proticol) into the enter"
    listDisplay.AddItem "box below and click the go button or press the enter key."
End Sub

Private Sub cmdPing_Click()

    If txtPing.Text = "" Then
       MsgBox "No IP address was entered.", , "Ping Help"
       Exit Sub
    End If

    ' Is user type /clear then clear display
    If txtPing.Text = "/clear" Then
       listDisplay.Clear
       listDisplay.AddItem "Cleared"
       txtPing.Text = ""
       Exit Sub
    End If
    
    ' Is user type /control-c then stop ping
    If txtPing.Text = "/control-c" Then
       Shell ("command.com /c Control-C"), vbHide
       listDisplay.AddItem ""
       listDisplay.AddItem "Ping Stopped"
       txtPing.Text = ""
       Exit Sub
    End If
    
    ' Is user type /help then stop ping
    If txtPing.Text = "/help" Then
       Call help
       txtPing.Text = ""
       Exit Sub
    End If
    
    If chkHide.Value = 0 Then
       ' Show dos window if hide is not checked
       Shell ("command.com /c ping " & txtPing.Text), vbNormalFocus
       cmdStop.Enabled = False
       chkHide.Enabled = False  ' Disable hide check
       'listDisplay.Clear
       listDisplay.AddItem ""
       listDisplay.AddItem "Pinging: " & txtPing.Text
       listDisplay.AddItem ""
       listDisplay.AddItem "Look in Window"
       listDisplay.AddItem "Press Control-C to stop ping in the external window"
       txtPing.Text = ""
    Else
       ' hide dos window if hide is checked
       Shell ("command.com /c ping " & txtPing.Text), vbHide
       cmdStop.Enabled = True
       cmdPing.Enabled = False
       chkHide.Enabled = False  ' Disable hide check
       'listDisplay.Clear
       listDisplay.AddItem ""
       listDisplay.AddItem "Pinging: " & txtPing.Text
       listDisplay.AddItem ""
       listDisplay.AddItem "Window is hidden"
       listDisplay.AddItem "Type /control-c in input box to stop ping,"
       listDisplay.AddItem "or press the stop button."
       txtPing.Text = ""
    End If
       
End Sub

