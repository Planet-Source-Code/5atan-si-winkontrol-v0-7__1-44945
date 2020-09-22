VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FRMTransfer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Transfer"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame sockDisplayFrame 
      Caption         =   "Socket Connections"
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4455
      Begin VB.Label Label5 
         Caption         =   "Screen Info Socket Status Unknown"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Image imgSockScreenInfo 
         Height          =   255
         Left            =   120
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Screen Socket Status Unknown"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   720
         Width           =   3855
      End
      Begin VB.Image imgSockScreen 
         Height          =   255
         Left            =   120
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Main Socket Status Unknown"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Width           =   3855
      End
      Begin VB.Image imgSockMain 
         Height          =   255
         Left            =   120
         Top             =   360
         Width           =   255
      End
   End
   Begin MSComctlLib.ImageList sockImageList 
      Left            =   2760
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMTransfer.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMTransfer.frx":0400
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3480
      Top             =   240
   End
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   2280
      Y1              =   120
      Y2              =   720
   End
   Begin VB.Label lblSent 
      Caption         =   "0"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblReceived 
      Caption         =   "0"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Sent:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Recieved:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FRMTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
' SetWindowPos Flags
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOZORDER = &H4
Const SWP_NOREDRAW = &H8
Const SWP_NOACTIVATE = &H10
Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Const SWP_SHOWWINDOW = &H40
Const SWP_HIDEWINDOW = &H80
Const SWP_NOCOPYBITS = &H100
Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering

Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
' SetWindowPos() hwndInsertAfter values
Const HWND_TOP = 0
Const HWND_BOTTOM = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Private Sub Form_Load()
Timer1.Enabled = True
'winFormOnTop me
'Me = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 1, 1, SWP_NOSIZE)
'SetWindowPos hwnd, conHwndTopmost, 45, 45, Me.Width / Screen.TwipsPerPixelY, Me.Height / Screen.TwipsPerPixelX, conSwpNoActivate Or conSwpShowWindow
Dim i
i = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

imgSockMain.Picture = sockImageList.ListImages(1).Picture 'Red Button
imgSockScreen.Picture = sockImageList.ListImages(1).Picture 'Red Button
imgSockScreenInfo.Picture = sockImageList.ListImages(1).Picture 'Red Button



End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False

End Sub

Private Sub Timer1_Timer()

tempreceive = Mod_Variables.BytesReceived
If tempreceive < 1048576 Then
    templabel = Format(tempreceive / 1024, "#0.0") & " KB"
Else
    templabel = Format(tempreceive / 1024 / 1024, "#0.0") & " MB"
End If
lblReceived.Caption = templabel



tempsent = Mod_Variables.BytesSent
If tempsent < 1024 Then
    templabel = Format(tempsent, "#0.0") & " Bytes"
ElseIf tempsent < 1048576 Then
    templabel = Format(tempsent / 1024, "#0.0") & " KB"
Else
    templabel = Format(tempsent / 1024 / 1024, "#0.0") & " MB"
End If
lblSent.Caption = templabel




Select Case FRMMain.Sock_Main.State
Case 0
    Label3.Caption = "Main Socket: " & "Not connected"
    imgSockMain.Picture = sockImageList.ListImages(1).Picture 'Red Button
    
Case 7
    Label3.Caption = "Main Socket: " & "Connected"
    imgSockMain.Picture = sockImageList.ListImages(2).Picture 'Red Button
End Select



Select Case FRMMain.Sock_Screen.State
Case 0
    Label4.Caption = "Screen Socket: " & "Not connected"
    imgSockScreen.Picture = sockImageList.ListImages(1).Picture 'Red Button
    
Case 7
    Label4.Caption = "Screen Socket: " & "Connected"
    imgSockScreen.Picture = sockImageList.ListImages(2).Picture 'Red Button
End Select



Select Case FRMMain.Sock_Screen_Info.State
Case 0
    Label5.Caption = "Screen Info Socket: " & "Not connected"
    imgSockScreenInfo.Picture = sockImageList.ListImages(1).Picture 'Red Button
    
Case 7
    Label5.Caption = "Screen Info Socket: " & "Connected"
    imgSockScreenInfo.Picture = sockImageList.ListImages(2).Picture 'Red Button
End Select

'Label3.Caption = "Main Socket: " & FRMMain.Sock_Main.State
'Label4.Caption = "Main Socket: " & FRMMain.Sock_Screen.State
'Label5.Caption = "Main Socket: " & FRMMain.Sock_Screen_Info.State


End Sub
