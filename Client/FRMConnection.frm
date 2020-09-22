VERSION 5.00
Begin VB.Form FRMConnection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WinKontrol: Connection"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   Icon            =   "FRMConnection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   4320
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton ButConnect 
         Caption         =   "Connect"
         Height          =   315
         Left            =   4320
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Pass_Box 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2160
         TabIndex        =   6
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox User_Box 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox IP_Box 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Text            =   "127.0.0.1"
         Top             =   360
         Width           =   2055
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   240
         Picture         =   "FRMConnection.frx":08CA
         Stretch         =   -1  'True
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label3 
         Caption         =   "Password:"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Username:"
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "IP Address:"
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "FRMConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButConnect_Click()

'Disable buttons
IP_Box.Enabled = False
User_Box.Enabled = False
Pass_Box.Enabled = False
Command1.Enabled = False
ButConnect.Enabled = False
FRMMain.Connection_Connect.Enabled = False
FRMMain.Connection_Disconnect.Enabled = True


'Set Main Socket variables
FRMMain.Sock_Main.RemoteHost = IP_Box.Text
FRMMain.Sock_Main.RemotePort = 10020

'Store username/ Password
Mod_Variables.Username = User_Box.Text
Mod_Variables.Password = Pass_Box.Text

'Connect to Server
Replied = False
FRMMain.Sock_Main.Connect

DownTime = 0
While (Not Replied) And (DownTime < 100000)
    DoEvents
    DownTime = DownTime + 1
Wend

If DownTime >= 100000 Then
    'Didn't reply or timed out. close the connection
    'MsgBox "Unable to connect to WinKontrol server", vbCritical, "Connection Error"
    FRMMain.Sock_Main.Close
    IP_Box.Enabled = True
    User_Box.Enabled = True
    Pass_Box.Enabled = True
    Command1.Enabled = True
    ButConnect.Enabled = True
    FRMMain.Connection_Connect.Enabled = True
    FRMMain.Connection_Disconnect.Enabled = False
    Exit Sub
End If

'FRMMain.ScreenCaptureBox.Visible = True
FRMMain.ISPanel1.Visible = True
FRMMain.MouseTimer.Enabled = True

Mod_Variables.BytesSent = 0
Mod_Variables.BytesReceived = 0


End Sub

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()
IP_Box.Text = Mod_Variables.IPAddress

End Sub

Private Sub Form_Unload(Cancel As Integer)
Mod_Variables.IPAddress = IP_Box.Text
End Sub
