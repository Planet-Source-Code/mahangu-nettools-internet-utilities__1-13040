VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmPort 
   Caption         =   "Port Flooding Session"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmPort.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3945
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   11
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop!"
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start!"
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   3240
      Width           =   975
   End
   Begin VB.Frame frmCopies 
      Caption         =   "No of Copies"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   1935
      Begin VB.TextBox txtCopies 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Text            =   "1"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame frameText 
      Caption         =   "Text to Send"
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   4335
      Begin VB.TextBox txtMessage 
         Height          =   1575
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame frameRemPort 
      Caption         =   "Remote Port"
      Height          =   615
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   1575
      Begin VB.TextBox txtRemPort 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame frameRemHost 
      Caption         =   "Remote Host"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
      Begin VB.TextBox txtRemHost 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock wsock 
      Left            =   3000
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConnect_Click()
If txtRemHost.Text = "" Then MsgBox "Hostname is empty!"
If txtRemPort.Text = "" Then MsgBox "Port name is empty!"

If txtRemPort.Text = "" Then GoTo 20
If txtRemHost.Text = "" Then GoTo 20
wsock.Connect txtRemHost, txtRemPort


20


End Sub

Private Sub cmdStart_Click()

If wsock.State <> sckConnected Then GoTo 20





While txtCopies.Text <> "0"

wsock.SendData (txtMessage.Text)


txtCopies.Text = txtCopies.Text - 1

frmLog.txtLog.Text = frmLog.txtLog.Text + "[Port Flooded " & txtRemHost.Text & " ," & txtRemPort.Text & " at " & Time$ & " ]"
Wend

20
MsgBox "Not connected to Server!"


End Sub

Private Sub Form_Load()
Me.Height = 4350
Me.Width = 4800

End Sub

Private Sub txtRemHost_Change()

Me.Caption = "New Port Flooding Session [Host - " & txtRemHost.Text & "] [Port - " & txtRemPort.Text & "]"
If txtRemHost.Text = "" Then Me.Caption = "New Port Flooding Session"

End Sub

Private Sub txtRemPort_Change()

Me.Caption = "New Port Flooding Session [Host - " & txtRemHost.Text & "] [Port - " & txtRemPort.Text & "]"
If txtRemHost.Text = "" Then Me.Caption = "New Port Flooding Session"

End Sub

Private Sub wsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "There has been an error while using winsock!"
End Sub
