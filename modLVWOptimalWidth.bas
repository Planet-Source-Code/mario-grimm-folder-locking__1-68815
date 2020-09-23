Attribute VB_Name = "modLVWOptimalWidth"
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
  (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, _
  ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
  (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
  (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) _
  As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
  ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
  ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Public Sub optimalWidth(lv As ListView)
  Dim t%
  
  On Error Resume Next
  For t = 0 To lv.ColumnHeaders.Count - 1
    If t = lv.ColumnHeaders.Count - 1 Then
        SendMessageLong lv.hwnd, &H101E, t, -2
    Else
        SendMessageLong lv.hwnd, &H101E, t, -1
    End If
  Next t

End Sub









