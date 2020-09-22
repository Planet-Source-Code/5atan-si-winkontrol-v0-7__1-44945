Attribute VB_Name = "Mod_Capture"
Option Explicit

Global colour As Integer

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020            ' (DWORD) dest = source
Public Const WHITENESS = &HFF0062          ' (DWORD) dest = WHITE
Public DIB As New cDIBSection
Public RecDib As New cDIBSection
Private ZLib As New clsZLib

Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Public C_Response As Boolean
Public C_Set_Response As Boolean
Dim ENDE As Boolean
Public Q As Double


Public Sub Capture()

Dim xPos As Long
Dim yPos As Long
Dim Ret As Long
Dim DeskHwnd As Long
Dim DeskHdc As Long
Dim DeskRect As RECT
Dim dibHdc As Long
Dim ByteArray() As Byte
Dim sValue As String

Const CRATE = 5
Dim CS(CRATE * CRATE * CRATE) As Long
                                    
Dim CS_Tmp As Long
Dim K As Long

DeskHwnd = GetDesktopWindow()
DeskHdc = GetDC(DeskHwnd)
Ret = GetWindowRect(DeskHwnd, DeskRect)

'Set Colour Depth - 16 colours
DIB.Colors = 16
RecDib.Colors = 16
    
Call DIB.Create(DeskRect.Right / CRATE, DeskRect.Bottom / CRATE)
Call RecDib.Create(DeskRect.Right / CRATE, DeskRect.Bottom / CRATE)
ENDE = False
K = 0
C_Response = False
C_Set_Response = False

Do Until ENDE
    Ret = BitBlt(DIB.hdc, 0, 0, DeskRect.Right / CRATE, DeskRect.Bottom / CRATE, DeskHdc, xPos, yPos, SRCCOPY)

    For yPos = 0 To DeskRect.Bottom Step (DeskRect.Bottom / CRATE)
        For xPos = 0 To DeskRect.Right Step (DeskRect.Right / CRATE)
            Ret = BitBlt(DIB.hdc, 0, 0, DeskRect.Right / CRATE, DeskRect.Bottom / CRATE, DeskHdc, xPos, yPos, SRCCOPY)
            Call DIB.ToByte(ByteArray) 'Save Desktop
            Call ZLib.CompressByte(ByteArray) 'Compress
            CS_Tmp = UBound(ByteArray) ' Save Checksum
      
            If CS_Tmp <> CS(K) Then
                CS(K) = CS_Tmp
                On Error GoTo NoConn
                FRMMain.Sock_Screen_Info.SendData (CStr(xPos) & ";" & CStr(yPos))
                Do Until C_Set_Response
                    DoEvents
                Loop
                C_Set_Response = False
                
                FRMMain.Sock_Screen.SendData ByteArray
                Do Until C_Response
                    DoEvents
                Loop
                C_Response = False
                On Error GoTo 0
            End If
        
            K = K + 1
            DoEvents
        Next xPos
    Next yPos

    xPos = 0
    yPos = 0
    K = 0
    Q = Q + 1
Loop

Exit Sub


NoConn:
    'Error Kontrol
    FRMMain.Sock_Screen.Close
    FRMMain.Sock_Screen_Info.Close
    FRMMain.Sock_Screen.Listen
    FRMMain.Sock_Screen_Info.Listen
End Sub
