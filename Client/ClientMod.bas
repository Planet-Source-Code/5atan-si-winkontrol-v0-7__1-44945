Attribute VB_Name = "ClientMod"
Option Explicit

Declare Function GetTickCount Lib "kernel32" () As Long
Public dX, dY As Long



Public Sub UpdateStatus(Status As String)
' Update status
FRMMain.StatusBar.SimpleText = Status
End Sub


Public Sub TerminateConnection()

FRMMain.Sock_Main.Close
FRMMain.Sock_Screen.Close
FRMMain.Sock_Screen_Info.Close
Replied = False
UpdateStatus "Connection Terminated"
FRMMain.Connection_Connect.Enabled = True
FRMMain.Connection_Disconnect.Enabled = False
'FRMMain.ScreenCaptureBox.Visible = False
FRMMain.ISPanel1.Visible = False
FRMMain.MouseTimer.Enabled = False

MouseCheck = 0



End Sub



Public Sub ConnectionError()

'Run this sub if a connection problem occurs
TerminateConnection
MsgBox "Connection Error! Connection Terminated"




End Sub

