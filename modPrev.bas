Attribute VB_Name = "modPrev"
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.
'
'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source !
'

'**************************************
'* modPrev.bas                        *
'* Module for window communication of *
'* CInstanceMonitor.cls               *
'* Programmed: Achim Neubauer         *
'* Last Change: 07.03.2004 16:07      *
'* Version: 1.0.1                     *
'**************************************

Option Explicit

Private Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type COPYDATASTRUCT
  dwData As Long
  cbData As Long
  lpData As Long
End Type

Private Const WM_COPYDATA As Long = &H4A

Public modPrev_EventTarget As CInstanceMonitor


Public Function modPrev_WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim result As Long
  
  Select Case uMsg
    Case WM_COPYDATA
      Dim CopyData As COPYDATASTRUCT
      Dim B() As Byte
      
      Dim DataType As Long
      Dim Data As String
      
      Call CopyMemory(CopyData, ByVal lParam, LenB(CopyData))
      
      DataType = CopyData.dwData
      If CopyData.cbData > 0 Then
        ReDim B(CopyData.cbData - 1)
        Call CopyMemory(B(0), ByVal CopyData.lpData, CopyData.cbData)
      
        Data = StrConv(B, vbUnicode)
      End If
      
      If Not modPrev_EventTarget Is Nothing Then Call modPrev_EventTarget.InternalEventRaiser(DataType, Data)
  End Select

  modPrev_WindowProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
End Function

Public Function modPrev_Address2Long(ByVal Address As Long) As Long
  modPrev_Address2Long = Address
End Function
