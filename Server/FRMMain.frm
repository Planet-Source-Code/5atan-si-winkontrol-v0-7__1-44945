VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form FRMMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WinKontrol: Server"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "FRMMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock sock_main 
      Left            =   2400
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   10020
   End
   Begin VB.Frame Frame2 
      Caption         =   "Status"
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   4455
      Begin MSWinsockLib.Winsock Sock_Screen_Info 
         Left            =   3240
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   10022
      End
      Begin MSWinsockLib.Winsock Sock_Screen 
         Left            =   2760
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   10021
      End
      Begin VB.TextBox StatusWindow 
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection Options"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.PictureBox Picture1 
         Height          =   1095
         Left            =   3000
         ScaleHeight     =   1035
         ScaleWidth      =   1275
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Server_Button 
         Caption         =   "Start Server"
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Mask_Button 
         Caption         =   "Mask Password"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox PassBox 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox UserBox 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "FRMMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

' Set Applications title bar / Version number
Me.Caption = "WinKontrol: Server Version v" & App.Major & "." & App.Minor & App.Revision
LoggedIn = False


End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Mask_Button_Click()

' Mask/ Unmask Password Box
Select Case Mask_Button.Caption
Case "Mask Password"
    PassBox.PasswordChar = "*"
    Mask_Button.Caption = "Unmask Password"
    Exit Sub
Case Else
    PassBox.PasswordChar = ""
    Mask_Button.Caption = "Mask Password"
    Exit Sub
End Select

End Sub

Private Sub Server_Button_Click()

Username = FRMMain.UserBox.Text
Password = FRMMain.PassBox.Text

' Code to determine to start/ end server connections
Select Case Server_Button.Caption
Case "Start Server"
    UserBox.Enabled = False
    PassBox.Enabled = False
    Mask_Button.Enabled = False
    Server_Button.Caption = "End Server"
    
    FRMMain.sock_main.Close
    FRMMain.Sock_Screen.Close
    FRMMain.Sock_Screen_Info.Close
    
    FRMMain.sock_main.Listen
    FRMMain.Sock_Screen.Listen
    FRMMain.Sock_Screen_Info.Listen
    FRMMain.StatusWindow.Text = ""
    UpdateStatus "Server Started"
    Exit Sub

Case Else
    UserBox.Enabled = True
    PassBox.Enabled = True
    Mask_Button.Enabled = True
    Server_Button.Caption = "Start Server"
    FRMMain.sock_main.Close
    FRMMain.Sock_Screen.Close
    FRMMain.Sock_Screen_Info.Close
    UpdateStatus "Server Closed"
    LoggedIn = False
    Exit Sub
    
End Select


End Sub

Private Sub sock_main_Close()
If FRMMain.sock_main.State <> sckClosed Then ConnectionError


End Sub

Private Sub sock_main_ConnectionRequest(ByVal requestID As Long)

'Accept Connection
FRMMain.sock_main.Close
FRMMain.sock_main.Accept requestID
FRMMain.sock_main.SendData "***OK:" & vbCrLf
UpdateStatus "Client Connected..."

TimeToLogin = 0
While (Not LoggedIn) And (TimeToLogin < 100000)
    DoEvents
    TimeToLogin = TimeToLogin + 1
Wend

If TimeToLogin >= 100000 Then
    LoggedIn = False
    FRMMain.sock_main.SendData "BADPS:" & vbCrLf
    UpdateStatus "Client supplied incorrect logon."
    Pause 1
    FRMMain.sock_main.Close
    FRMMain.sock_main.Listen
Else
    FRMMain.sock_main.SendData "**RES:" & Screen.Width / Screen.TwipsPerPixelX & "x" & Screen.Height / Screen.TwipsPerPixelY
    LoggedIn = True
End If





End Sub

Private Sub sock_main_DataArrival(ByVal bytesTotal As Long)
'Get incoming data and unsplit
Dim SplitData, strData As String
FRMMain.sock_main.GetData strData
SplitData = Split(strData, vbCrLf)


'On Error GoTo ErrKontrol

'Parse incoming data
For i = 0 To UBound(SplitData)

    Select Case Left$(SplitData(i), 6)

        Case "*PASS:"
            UserPass = Right(SplitData(i), Len(SplitData(i)) - 6)
            UserPass = Split(UserPass, ",")
            CheckUser = UserPass(0)
            CheckPass = UserPass(1)
            If (CheckUser = Username) And (CheckPass = Password) Then LoggedIn = True
            'Debug.Print "USER/PASS: '" & user & "' -- '" & pass & "'"
            'Debug.Print "USER/PASS: '" & CheckUser & "' -- '" & CheckPass & "'"
            
    End Select

Next i


End Sub

Private Sub Sock_Screen_ConnectionRequest(ByVal requestID As Long)
    Sock_Screen.Close
    Sock_Screen.Accept requestID

End Sub

Private Sub Sock_Screen_DataArrival(ByVal bytesTotal As Long)
    C_Response = True
End Sub

Private Sub Sock_Screen_Info_ConnectionRequest(ByVal requestID As Long)
    Sock_Screen_Info.Close
    Sock_Screen_Info.Accept requestID
    Mod_Capture.Capture 'PicColours, CRATE

End Sub

Private Sub Sock_Screen_Info_DataArrival(ByVal bytesTotal As Long)
    C_Set_Response = True
    
    Dim Command      As String
    Dim NewArrival   As String
    Dim Data         As String

    Sock_Screen_Info.GetData NewArrival ', vbString
    Debug.Print "ARRIVE=" & NewArrival
    Command$ = EvalData(NewArrival$, 1) 'Get Command from New Data (before the ,)
    Data$ = EvalData(NewArrival$, 2)   ' Get Data from new Data (After the ,)
    
    Select Case Command$
    
        'MOUSE MOVED/ BUTTON PRESSED ----------------------------------------------------
        Case "POS"
        'Split Mouse Positions into X, Y, & MouseButton
        'Dim SplitData As String
        SplitData = Split(Data$, ":")
        x = SplitData(0): y = SplitData(1): Button = SplitData(2)
        'Set MouseCursor Position
        SetCursorPos x, y
        'Read MouseButton Configs & Activate Button Accordingly
        'Debug.Print Button
        'Label1.Caption = LastMouseButton & " : " & Button
    
        Select Case Left$(Button, 1)
            Case "L" 'Left mouse button
                If LastMouseButton <> "L" Then Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                LastMouseButton = "L"
            Case "R" 'Right Mouse Button
                If LastMouseButton <> "R" Then Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
                LastMouseButton = "R"
            Case "0" 'No Mouse Buttons
                If LastMouseButton = "L" Then
                    Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                ElseIf LastMouseButton = "R" Then
                    Call mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
                End If
                LastMouseButton = "0"
            End Select
            
        
        'KEYBOARD BUTTON PRESSED ----------------------------------------------------
        Case "KEY"
            'KeyStroke = Data$
            'Debug.Print Data$
            keybd_event Data$, 0, 0, 0
            keybd_event Data$, 0, KEYEVENTF_KEYUP, 0
            DoEvents

          
    End Select
    
    
    
End Sub

