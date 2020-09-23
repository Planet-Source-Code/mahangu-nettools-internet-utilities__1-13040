VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmScan 
   Caption         =   "Port Scanner"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   Icon            =   "frmScan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   7095
   Begin VB.TextBox FoundPorts 
      Height          =   2175
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   840
      Width           =   6735
   End
   Begin VB.TextBox txtPortEnd 
      Height          =   285
      Left            =   4080
      TabIndex        =   3
      Text            =   "65536"
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtPortStart 
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Text            =   "1"
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "Localhost"
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   285
      Left            =   5040
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock Sock 
      Index           =   0
      Left            =   6600
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waiting..."
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   6135
   End
End
Attribute VB_Name = "frmScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'************************************************
'* This is where it gets a bit more complicated *
'************************************************
Dim Socket As Variant ' for instances of the socket we will
                      ' use in the For loop

Dim CurrentPort As Integer ' Obvious

Const MaxSockets = 50 ' change this for Speed / Accuracy
                       ' between 1 - 200

' it's stable enough to use this
On Error Resume Next



' We need a way to Start / Stop, so we'll use
' the command button's caption as a reference
If Command1.Caption = "Start" Then

' to prevent errors, disable teh textboxes
txtHost.Enabled = False
txtPortStart.Enabled = False
txtPortEnd.Enabled = False


'see above
Command1.Caption = "Stop"
    ' Lets load some sockets to use
    For i = 1 To MaxSockets
        'Load new sock instance i
        Load Sock(i)
    Next i
    CurrentPort = txtPortStart.Text
    ' Again using the command1.caption as a reference
    ' to start / stop
    While Command1.Caption = "Stop"
        ' set up the ports to scan by referencing
        ' each instance of the socket in turn
        For Each Socket In Sock
            ' Definately Need this so the system doesn't freeze
            DoEvents
            ' check if the socket is still trying to connect
            ' or is connected
            If Socket.State <> sckClosed Then
                ' skip the increment of the port
                GoTo continue
            End If
            ' close the socket to make double sure
            Socket.Close
            ' if it got to here, it's ready to try
            ' the next port, only after checking
            ' if we've done all the ports and the user
            ' hasn't clicked on Stop
            
            If CurrentPort = Val(txtPortEnd.Text) + 1 _
            Then Exit For
            'set the host
            Socket.RemoteHost = txtHost.Text
            ' set the port
            Socket.RemotePort = CurrentPort
            ' inform the user of the port being scanned
            lblStatus.Caption = "Now Scanning Port " & CurrentPort
            ' attempt connect
            Socket.Connect
            ' fromhere, the socket will do one of two things
            ' 1) Raise a Connect therefore the port is open
            ' 2) Raise an Error therefore the port is closed
            
            ' increment the current port
            CurrentPort = CurrentPort + 1
' if the socketisn't ready to be incremented, go here
continue:
        
        ' goto the next socket instance
        Next Socket
    Wend
'set the command1.caption to Start so we can scan again
Command1.Caption = "Start"

' re-enable the textboxes
txtHost.Enabled = True
txtPortStart.Enabled = True
txtPortEnd.Enabled = True

Else ' command1.caption is "Stop"
    Command1.Caption = "Start"
End If

' close all the sockets to save memory
For i = 1 To MaxSockets
    Unload Sock(i)
Next i

End Sub

Private Sub FoundPorts_Change()
'****************************************************
'* So that out textbox scrolls down automatically   *
'* we use the SelStart property in the              *
'* FoundPorts_change Event.                         *
'****************************************************

' Pseudo code
'~~~~~~~~~~~~
' Selection start position = length of Text in Text control

FoundPorts.SelStart = Len(FoundPorts.Text)
End Sub


Private Function AddPort(Port As Integer)
'**************************************************
'* This is a function to add the port to the list *
'**************************************************

'Pseudo code
'~~~~~~~~~~~
' Text = current text + newtext + carriage return

FoundPorts.Text = FoundPorts.Text & "[Connected] Port " & Port & vbCrLf
End Function

Private Sub Sock_Connect(Index As Integer)
' the port is open so inform the user
AddPort (Sock(Index).RemotePort)
' close the socket so it can't be flooded by anti
' portscanner tools and it gets incremented
Sock(Index).Close
End Sub

Private Sub Sock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
' the port is closed so close the socket so it
' will be incremented
Sock(Index).Close
End Sub

