Attribute VB_Name = "Mod_Server"
Option Explicit

Declare Function GetTickCount Lib "kernel32" () As Long

Public Username As String
Public Password As String
Public LoggedIn As Boolean


Public Sub UpdateStatus(Status As String)

FRMMain.StatusWindow.Text = FRMMain.StatusWindow.Text + Status + Chr$(13) + Chr(10)
FRMMain.StatusWindow.SelLength = Len(FRMMain.StatusWindow.Text)

End Sub


Public Sub ConnectionError()

FRMMain.sock_main.Close
FRMMain.Sock_Screen.Close
FRMMain.Sock_Screen_Info.Close

FRMMain.sock_main.Listen
FRMMain.Sock_Screen.Listen
FRMMain.Sock_Screen_Info.Listen
UpdateStatus "Connection Error. Server Reset..."
LoggedIn = False


End Sub



Public Function EvalData(sIncoming As String, iRtLt As Integer, _
                  Optional sDivider As String) As String
   Dim i As Integer
   Dim tempStr As String
   ' Storage for the current Divider
   Dim sSplit As String
   
   ' the current character used to divide the data
   If sDivider = "" Then
      sSplit = ","
   Else
      sSplit = sDivider
   End If
   
   ' getting the right or left?
   Select Case iRtLt
        
      Case 1
          ' remove the data to the Left of the Current Divider
          For i = 0 To Len(sIncoming)
            tempStr = Left(sIncoming, i)
            
            If Right(tempStr, 1) = sSplit Then
              EvalData = Left(tempStr, Len(tempStr) - 1)
              Exit Function
            End If
          Next
          
      Case 2
          ' remove the data to the Right of the Current Divider
          For i = 0 To Len(sIncoming)
            tempStr = Right(sIncoming, i)
            
            If Left(tempStr, 1) = sSplit Then
              EvalData = Right(tempStr, Len(tempStr) - 1)
              Exit Function
            End If
          Next
   End Select
   
End Function

Sub Pause(HowLong As Long)
    Dim u%, tick As Long
    tick = GetTickCount()
    Do
        u% = DoEvents
    Loop Until tick + HowLong < GetTickCount
End Sub
