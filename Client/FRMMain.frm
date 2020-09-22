VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm FRMMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "WinKontrol: Client"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8565
   Icon            =   "FRMMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "FRMMain.frx":08CA
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5505
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Sock_Main 
      Left            =   120
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox BackgroundBox 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   11640
      Left            =   0
      Picture         =   "FRMMain.frx":AAD0
      ScaleHeight     =   11610
      ScaleWidth      =   8535
      TabIndex        =   2
      Top             =   420
      Width           =   8565
      Begin WinkClient.ISPanel ISPanel1 
         Height          =   4815
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   8493
         BackColor       =   0
         Begin VB.PictureBox ScreenCaptureBox 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1935
            Left            =   0
            ScaleHeight     =   127
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   223
            TabIndex        =   4
            Top             =   0
            Width           =   3375
            Begin VB.Timer MouseTimer 
               Enabled         =   0   'False
               Interval        =   10
               Left            =   120
               Top             =   840
            End
         End
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   240
         Top             =   1320
      End
      Begin MSWinsockLib.Winsock Sock_Screen_Info 
         Left            =   1080
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Sock_Screen 
         Left            =   600
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.Menu ConnectionMenu 
      Caption         =   "&Connection"
      Begin VB.Menu Connection_Connect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu Connection_Disconnect 
         Caption         =   "&Disconnect"
         Enabled         =   0   'False
      End
      Begin VB.Menu ConnectionBar1 
         Caption         =   "-"
      End
      Begin VB.Menu Connection_Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu WindowsMenu 
      Caption         =   "&Windows"
      Begin VB.Menu Windows_Transfer 
         Caption         =   "&Transfer Rate"
      End
   End
End
Attribute VB_Name = "FRMMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

Public RecDib As New cDIBSection
Dim ZLib As New clsZLib
Dim i As Long
Dim tmpPos() As String
Dim Pos(1) As Long
Const CRate = 5

Private a As New cFormSize







Private Sub Connection_Connect_Click()

FRMConnection.Show 1


End Sub

Private Sub Connection_Disconnect_Click()

reply = MsgBox("Terminate Current Connection?", vbYesNo + vbQuestion, "WinKontrol")
If reply = 6 Then TerminateConnection Else Exit Sub


End Sub

Private Sub Connection_Exit_Click()

'Check connection status. 7= Connected. 0=Not Connected
Debug.Print FRMMain.Sock_Main.State
If FRMMain.Sock_Main.State = 7 Then
    MsgBox "Terminate The Connection Before Quiting!"
    Exit Sub
End If

End





End Sub

Private Sub MDIForm_Load()



ISPanel1.Attatch ScreenCaptureBox
'ScreenCaptureBox.Visible = False
FRMMain.ISPanel1.Visible = False


'Resize Application -------------------------------
    ' Initialize the class
    a.Init Me.hwnd
    
    ' Set max. and min. sizes, in twips !
    a.MaxWidth = 9000
    a.MaxHeight = 7000
    
    a.MinWidth = 7000
    a.MinHeight = 5000
    
    ' Resize the form to its max size. This is needed to avoid
    ' starting with a wrong size...
    'a.ResizeToMax
'-------------------------------------------------

' Set Applications title bar / Version number
Me.Caption = "WinKontrol: Client Version v" & App.Major & "." & App.Minor & App.Revision
UpdateStatus "Ready..."




End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' If Connected Do Nothing
If FRMMain.Sock_Main.State = 7 Then
    MsgBox "Terminate The Connection Before Quiting!"
    Cancel = 1
End If

End Sub

Private Sub MDIForm_Resize()

If FRMMain.WindowState <> 1 Then

    BackgroundBox.Height = FRMMain.Height - 1500

    ISPanel1.Width = FRMMain.Width - 145
    ISPanel1.Height = FRMMain.Height - 1550

End If




End Sub


Private Sub MDIForm_Terminate()
End

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub ScreenCaptureBox_KeyDown(KeyCode As Integer, Shift As Integer)

On Error Resume Next
FRMMain.Sock_Screen_Info.SendData "KEY," & KeyCode & vbCrLf
Mod_Variables.BytesSent = Mod_Variables.BytesSent + 4 + Len(KeyCode)



End Sub

Private Sub ScreenCaptureBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

CheckMouse

End Sub

Private Sub ScreenCaptureBox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

CheckMouse

End Sub

Public Sub CheckMouse()

MouseCheck = MouseCheck + 1
If MouseCheck < 10 Then Exit Sub 'Do not respond to mouse actions


'On Error Resume Next

Dim pt As POINTAPI
Dim rv As Long

If DontCount < 100 Then DontCount = DontCount + 1
rv& = GetCursorPos(pt)
ReDim Preserve cp(idx)

MouseX = pt.x - (FRMMain.Left / Screen.TwipsPerPixelX) - 5
MouseY = pt.y - (FRMMain.Top / Screen.TwipsPerPixelY) - 80
    'FRMMain.Width / Screen.TwipsPerPixelX - 10
    ' ---  Width of viewable screen
    
'Debug.Print "FUK = " & FRMMain.Height / Screen.TwipsPerPixelY - 100
'Debug.Print "moX = " & Mousey
'Debug.Print "MAX = " & FRMMain.ScreenCaptureBox.Height / Screen.TwipsPerPixelY

If MouseX < 0 Then Exit Sub
If MouseY < 0 Then Exit Sub
If MouseX > (FRMMain.Width / Screen.TwipsPerPixelX - 10) Then Exit Sub
If MouseY > (FRMMain.Height / Screen.TwipsPerPixelY - 100) Then Exit Sub
'If MouseX > ScreenCaptureBox.Width / Screen.TwipsPerPixelX Then Exit Sub


'Debug.Print "X= "; MouseX
'Debug.Print "Y= "; MouseY
'Debug.Print DontCount
If DontCount > 50 Then 'WAIT! Dont count First Keystroke
    If GetAsyncKeyState(VK_LBUTTON) Then
        'Label1.Caption = "Left Button Down"
        ButtonDown = "L"
    ElseIf GetAsyncKeyState(VK_RBUTTON) Then
        'Label1.Caption = "Right Button Down"
        ButtonDown = "R"
    Else
        'Label1.Caption = ""
        ButtonDown = "0"
    End If
'MsgBox ButtonDown
End If
'Debug.Print ButtonDown


If MouseX = lastX And MouseY = LastY And ButtonDown = LastMouseButton Then
        'Dont Send Data to client
    Else
        'Image2.Visible = False
        'LightCounter.Enabled = True
        FRMMain.Sock_Screen_Info.SendData "POS," & Int(MouseX) & ":" & Int(MouseY) & ":" & ButtonDown & vbCrLf
        Mod_Variables.BytesSent = Mod_Variables.BytesSent + 6 + Len(Int(MouseX)) + Len(Int(MouseY))
End If

lastX = MouseX
LastY = MouseY
LastMouseButton = ButtonDown

End Sub

Private Sub Sock_Main_Close()

TerminateConnection
MsgBox "The Connection has been closed by the Server"
'FRMMain.Sock_Main.Close
'FRMMain.Sock_Screen.Close
'FRMMain.Sock_Screen_Info.Close



End Sub

Private Sub Sock_Main_DataArrival(ByVal bytesTotal As Long)

'Get incoming data and unsplit
Dim SplitData, strData As String
FRMMain.Sock_Main.GetData strData
Mod_Variables.BytesReceived = Mod_Variables.BytesReceived + bytesTotal

Debug.Print FRMMain.Sock_Main.BytesReceived

SplitData = Split(strData, vbCrLf)


'On Error GoTo ErrKontrol

'Parse incoming data
For i = 0 To UBound(SplitData)

    'Debug.Print Right$(SplitData(i), 6)
    
    Select Case Left$(SplitData(i), 6)

        Case "***OK:"
            Replied = True
            Unload FRMConnection
            UpdateStatus "Connected"
            FRMMain.Connection_Connect.Enabled = False
            FRMMain.Connection_Disconnect.Enabled = True
            FRMMain.Sock_Main.SendData "*PASS:" & Username & "," & Password & vbCrLf
            Mod_Variables.BytesSent = Mod_Variables.BytesSent + 7 + Len(Username) + Len(Password)
            
        Case "BADPS:"
            UpdateStatus "Disconnected"
            MsgBox "Bad Username/ Password"
            FRMMain.Sock_Main.Close
            FRMMain.Connection_Connect.Enabled = True
            FRMMain.Connection_Disconnect.Enabled = False
            
        Case "**RES:"
            'Get resolution
            Res = Right$(SplitData(i), Len(SplitData(i)) - 6)
            'Debug.Print "REs:  " & Res
            Ress = Split(Res, "x")
            ClientMod.dX = Ress(0)
            ClientMod.dY = Ress(1)
            RecDib.Colors = 16
            Call RecDib.Create(ClientMod.dX / CRate, ClientMod.dY / CRate)
            i = 0
            FRMMain.Sock_Screen.Connect Mod_Variables.IPAddress, 10021
            FRMMain.Sock_Screen_Info.Connect Mod_Variables.IPAddress, 10022
            While (FRMMain.Sock_Screen_Info.State <> 5) And (DownTimesock < 100000)
                'DoEvents
                DownTimesock = DownTimesock + 1
            Wend
            
            ScreenCaptureBox.Height = ClientMod.dY * Screen.TwipsPerPixelX
            ScreenCaptureBox.Width = ClientMod.dX * Screen.TwipsPerPixelY
            
            a.MaxWidth = ClientMod.dX * Screen.TwipsPerPixelY + 450
            a.MaxHeight = ClientMod.dY * Screen.TwipsPerPixelX + 1850
            
            MouseTimer.Enabled = True
            
            
    End Select

Next i

Exit Sub
ErrKontrol:
ConnectionError

End Sub

Private Sub Sock_Main_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
ConnectionError

End Sub

Private Sub Sock_Screen_DataArrival(ByVal bytesTotal As Long)
Dim Ret As Long
Dim ByteArray() As Byte
FRMMain.Sock_Screen.GetData ByteArray, vbByte

Mod_Variables.BytesReceived = Mod_Variables.BytesReceived + bytesTotal


'Debug.Print ByteArray

Call ZLib.DecompressByte(ByteArray)
DoEvents

Call RecDib.ParseByte(ByteArray)
'Debug.Print Pos(0); Pos(1); ClientMod.dx / CRate; ClientMod.dy / CRate; ClientMod.dx; ClientMod.dy
Ret = BitBlt(FRMMain.ScreenCaptureBox.hdc, Pos(0), Pos(1), ClientMod.dX / CRate, ClientMod.dY / CRate, RecDib.hdc, 0, 0, SRCCOPY)
FRMMain.ScreenCaptureBox.Refresh
FRMMain.Sock_Screen.SendData "a"
Mod_Variables.BytesSent = Mod_Variables.BytesSent + 1

'SavePicture FRMMain.ScreenCaptureBox.Picture, "c:\test.bmp"


End Sub

Private Sub Sock_Screen_Info_DataArrival(ByVal bytesTotal As Long)

    Dim tmp As String
    FRMMain.Sock_Screen_Info.GetData tmp, vbString
    Mod_Variables.BytesReceived = Mod_Variables.BytesReceived + bytesTotal
    tmpPos = Split(tmp, ";")
    Pos(0) = CLng(tmpPos(0))
    Pos(1) = CLng(tmpPos(1))
    Sock_Screen_Info.SendData "a"
    Mod_Variables.BytesSent = Mod_Variables.BytesSent + 1
    'Debug.Print "POS= " & Pos(0) & ";" & Pos(1)
    
End Sub

Private Sub Timer1_Timer()
FRMMain.ScreenCaptureBox.Refresh
End Sub

Private Sub VScroll_Change()
ScreenCaptureBox.Top = -VScroll
End Sub


Public Sub EnableScrollBars()

    
End Sub

Private Sub Windows_Transfer_Click()

If Windows_Transfer.Checked = True Then
    Windows_Transfer.Checked = False
    Unload FRMTransfer
    Exit Sub
Else
    Windows_Transfer.Checked = True
    FRMTransfer.Show
    Exit Sub
End If

    

End Sub
