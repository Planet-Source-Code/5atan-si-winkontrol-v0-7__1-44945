VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFormSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Message definition
Private Const WM_GETMINMAXINFO = &H24

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
    
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Private m_lHwnd As Long
Private m_lMinHeigth As Long
Private m_lMinWidth As Long
Private m_lMaxHeigth As Long
Private m_lMaxWidth As Long

Implements ISubclass

Private m_emr As EMsgResponse


'********************************************************************************
' Procedure:      Sub Init
' Author:         M. Novaro
' Creation Date:  09/09/2002
' Description:    Initialize the class with the hwnd of the form that shall be resized
' Notes:
'
'
'********************************************************************************
Public Sub Init(hwnd As Long)
   
    AttachMessage Me, hwnd, WM_GETMINMAXINFO
    m_lHwnd = hwnd
    
End Sub

'********************************************************************************
' Procedure:      Sub ResizeToMax
' Author:         M. Novaro
' Creation Date:  09/09/2002
' Description:    Resize the form to the maximum size MaxWidth, MaxHeigth
' Notes:
'
'
'********************************************************************************
Public Sub ResizeToMax()
    Dim winRect As RECT
    
    GetWindowRect m_lHwnd, winRect
    MoveWindow m_lHwnd, winRect.Left + 1, winRect.Top, m_lMaxWidth, m_lMaxHeigth, 0
End Sub

'********************************************************************************
' Procedure:      Sub Class_Initialize
' Author:         M. Novaro
' Creation Date:  09/09/2002
' Description:
' Notes:
'
'
'********************************************************************************
Private Sub Class_Initialize()
    m_lHwnd = 0
    
    m_lMinHeigth = 0
    m_lMinWidth = 0
    
    m_lMaxHeigth = Screen.Height / Screen.TwipsPerPixelY
    m_lMaxWidth = Screen.Width / Screen.TwipsPerPixelX
End Sub

'********************************************************************************
' Procedure:      Sub Class_Terminate
' Author:         M. Novaro
' Creation Date:  09/09/2002
' Description:
' Notes:
'
'
'********************************************************************************
Private Sub Class_Terminate()
    If m_lHwnd <> 0 Then
        DetachMessage Me, m_lHwnd, WM_GETMINMAXINFO
    End If
End Sub

'********************************************************************************
' Procedure:      Property ISubclass_MsgResponse
' Author:         M. Novaro
' Creation Date:  09/09/2002
' Description:
' Notes:
'
'
'********************************************************************************
Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
    m_emr = RHS
End Property

'********************************************************************************
' Procedure:      Property ISubclass_MsgResponse
' Author:         M. Novaro
' Creation Date:  09/09/2002
' Description:
' Notes:
'
'
'********************************************************************************
Private Property Get ISubclass_MsgResponse() As SSubTimer6.EMsgResponse
'    Debug.Print CurrentMessage
    m_emr = emrConsume
    ISubclass_MsgResponse = m_emr
End Property

'********************************************************************************
' Procedure:      Function ISubclass_WindowProc
' Author:         M. Novaro
' Creation Date:  09/09/2002
' Description:
' Notes:
'
'
'********************************************************************************
Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim mmiT As MINMAXINFO

    ' Copy parameter to local variable for processing
    CopyMemory mmiT, ByVal lParam, LenB(mmiT)

    ' Minimium width and height for sizing
    mmiT.ptMinTrackSize.x = m_lMinWidth
    mmiT.ptMinTrackSize.y = m_lMinHeigth
    
    mmiT.ptMaxTrackSize.x = m_lMaxWidth
    mmiT.ptMaxTrackSize.y = m_lMaxHeigth
    
    ' Copy modified results back to parameter
    CopyMemory ByVal lParam, mmiT, LenB(mmiT)

End Function

'********************************************************************************
' Procedure:      Property MinWidth
' Author:         M. Novaro
' Creation Date:  09/09/2002
' Description:    Set minimum width of the form
' Notes:
'
'
'********************************************************************************
Property Let MinWidth(value As Long)
    m_lMinWidth = value / Screen.TwipsPerPixelX
End Property

'********************************************************************************
' Procedure:      Property MinHeigth
' Author:         M. Novaro
' Creation Date:  09/09/2002
' Description:    Set minimum height of the form
' Notes:
'
'
'********************************************************************************
Property Let MinHeight(value As Long)
    m_lMinHeigth = value / Screen.TwipsPerPixelY
End Property

'********************************************************************************
' Procedure:      Property MaxWidth
' Author:         M. Novaro
' Creation Date:  09/09/2002
' Description:    Set maximum height of the form
' Notes:
'
'
'********************************************************************************
Property Let MaxWidth(value As Long)
    m_lMaxWidth = value / Screen.TwipsPerPixelX
End Property

'********************************************************************************
' Procedure:      Property MaxHeigth
' Author:         M. Novaro
' Creation Date:  09/09/2002
' Description:    Set maximum height of the form
' Notes:
'
'
'********************************************************************************
Property Let MaxHeight(value As Long)
    m_lMaxHeigth = value / Screen.TwipsPerPixelY
End Property

