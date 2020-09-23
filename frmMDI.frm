VERSION 5.00
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "NetTools 1.0.0"
   ClientHeight    =   5130
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8385
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Begin VB.Menu mnuAnon 
            Caption         =   "&Anon Mailing Session"
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuNewMail 
            Caption         =   "&Mail Bombing Session"
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuNewIcq 
            Caption         =   "&ICQ Flooding Session"
            Shortcut        =   ^I
         End
         Begin VB.Menu mnuNewPing 
            Caption         =   "&Ping Session"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnuNewFlood 
            Caption         =   "&Port Flooding Session"
            Shortcut        =   ^F
         End
         Begin VB.Menu mnuNewScan 
            Caption         =   "&Port Scanning Session"
            Shortcut        =   ^S
         End
      End
      Begin VB.Menu mnuViewLog 
         Caption         =   "&ViewLog"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MailCount

Private Sub mnuAbout_Click()
frmAbout.Show


End Sub

Private Sub mnuAnon_Click()
Dim NewAnon As New frmAnon
NewAnon.Show

End Sub

Private Sub mnuExit_Click()
End

End Sub

Private Sub mnuNewSession_Click()
Dim NewSession As New frmMain
NewSession.Show

End Sub


Private Sub mnuNewFlood_Click()
If MailCount > 2 Then MsgBox "You can't open so many windows!"
If MailCount > 2 Then GoTo 40

Dim NewFlood As New frmPort
NewFlood.Show

MailCount = MailCount + 1

40


End Sub

Private Sub mnuNewIcq_Click()
If MailCount > 2 Then MsgBox "You can't open so many windows!"
If MailCount > 2 Then GoTo 20

Dim NewIcq As New frmIcq
NewIcq.Show

MailCount = MailCount + 1

20

End Sub

Private Sub mnuNewMail_Click()
If MailCount > 2 Then MsgBox "You can't open so many windows!"
If MailCount > 2 Then GoTo 10
Dim NewMail As New frmMain
NewMail.Show

MailCount = MailCount + 1

10
End Sub

Private Sub mnuNewPing_Click()
If MailCount > 2 Then MsgBox "You can't open so many windows!"
If MailCount > 2 Then GoTo 30

Dim NewPing As New frmPing
NewPing.Show

MailCount = MailCount + 1

30

End Sub

Private Sub mnuNewScan_Click()
If MailCount > 2 Then MsgBox "You can't open so many windows!"
If MailCount > 2 Then GoTo 50

Dim NewScan As New frmScan
NewScan.Show

MailCount = MailCount + 1

50

End Sub

Private Sub mnuViewLog_Click()
frmLog.Show



End Sub
