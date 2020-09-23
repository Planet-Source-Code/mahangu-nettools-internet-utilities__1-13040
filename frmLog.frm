VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLog 
   Caption         =   "NetTools LogFile"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   2160
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save Log File"
      FileName        =   "LogFile"
      Filter          =   "*.log"
   End
   Begin VB.TextBox txtLog 
      Height          =   3135
      Left            =   0
      ScrollBars      =   1  'Horizontal
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
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
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
txtLog.Height = ScaleHeight
txtLog.Width = ScaleWidth

End Sub

Private Sub mnuAbout_Click()
Call ShowAbout

End Sub

Private Sub mnuExit_Click()
Unload Me

End Sub

Private Sub mnuSave_Click()
dlgSave.ShowSave

If dlgSave.FileName <> "" Then
Open dlgSave.FileName For Output As #1
Print #1, txtLog.Text
Close 1
End If

End Sub
