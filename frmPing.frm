VERSION 5.00
Begin VB.Form frmPing 
   Caption         =   "Ping Session"
   ClientHeight    =   2205
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmPing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2205
   ScaleWidth      =   4680
   Begin VB.Frame Frame1 
      Caption         =   "No of time to Ping"
      Height          =   855
      Left            =   1800
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
      Begin VB.TextBox txtCopies 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   2
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame frameIP 
      Caption         =   "IP Address"
      Height          =   855
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   2655
      Begin VB.TextBox txtIP 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Ping!"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Shape shpMain 
      Height          =   1335
      Left            =   120
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdStart_Click()
Dim Copies
Copies = txtCopies.Text



While txtCopies.Text <> "0"

Me.Cls
   
   Dim ECHO As ICMP_ECHO_REPLY
   Dim pos As Integer
   
  'ping an ip address, passing the
  'address and the ECHO structure
   Call Ping(txtIP, ECHO)
   
  'display the results from the ECHO structure
  Me.Print ""
   Me.Print "    " & GetStatusCode(ECHO.STATUS)
   Me.Print "    " & ECHO.Address
   Me.Print "    " & ECHO.RoundTripTime & " ms"
   Me.Print "    " & ECHO.DataSize & " bytes"



   If Left$(ECHO.DATA, 1) <> Chr$(0) Then
      pos = InStr(ECHO.DATA, Chr$(0))
      Me.Print "    " & Left$(ECHO.DATA, pos - 1)
   End If

   Me.Print "    " & ECHO.DataPointer

txtCopies.Text = txtCopies.Text - 1


Wend

frmLog.txtLog.Text = frmLog.txtLog.Text + "[Pinged " & txtIP.Text & ", " & Copies & " times and finished at " & Time$ & " ]"

End Sub

Private Sub mnuStop_Click(Index As Integer)
End

End Sub

Private Sub Form_Load()
Me.Height = 2610
Me.Width = 4800
End Sub

Private Sub txtCopies_Click()
txtCopies.Text = "0"
End Sub

Private Sub txtIP_Change()
Me.Caption = "Ping Session [To - " & txtIP.Text & " ]"
If txtIP.Text = "" Then Me.Caption = "Ping Session"
End Sub
