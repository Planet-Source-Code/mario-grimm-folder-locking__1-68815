VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInstanceMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.
'
'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source !
'

'****************************************
'* CInstanceMonitor.cls                 *
'* Enables unidirectional communication *
'* to the first instance of the program *
'* Programmed: Achim Neubauer           *
'* Last Change: 07.03.2004 16:07        *
'* Version: 1.0.1                       *
'****************************************

Option Explicit

'*****************************************************
'* Edit this constant to a unique ID of your program *
   Private Const ProgramID As String = "LockingFolderApplication"
'*****************************************************

'API Deklarationen

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function RegisterClass Lib "user32.dll" Alias "RegisterClassA" (lpWndClass As WNDCLASS) As Long
Private Declare Function UnregisterClass Lib "user32.dll" Alias "UnregisterClassA" (ByVal lpClassName As Any, ByVal hInstance As Long) As Long
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type COPYDATASTRUCT
  dwData As Long
  cbData As Long
  lpData As Long
End Type

Private Type WNDCLASS
  style As Long
  lpfnWndProc As Long
  cbClsExtra As Long
  cbWndExtra As Long
  hInstance As Long
  hIcon As Long
  hCursor As Long
  hbrBackground As Long
  lpszMenuName As String
  lpszClassName As String
End Type

Private Const WS_EX_APPWINDOW = &H40000
Private Const WM_CLOSE As Long = &H10
Private Const WM_COPYDATA As Long = &H4A

'***

Event DataArrival(DataType As Long, Data As String)

Private m_hWnd As Long
Private m_hAtom As Long
Private m_PrevInstance As Boolean
Private m_PrevHandle As Long

Private Function MakeClassName(ProgramID As String) As String
  MakeClassName = ProgramID & " PrevInstance Monitor"
End Function

Private Function FindMonitor() As Long
  If Len(ProgramID) > 0 Then
    FindMonitor = FindWindow(MakeClassName(ProgramID), ProgramID)
  End If
End Function

Private Function InstallMonitor() As Boolean
  If Len(ProgramID) = 0 Then Exit Function 'Keine Programm-Idendifikation angegeben.
  
  Dim MonitorClass As WNDCLASS
  Dim result As Long
  
  If m_hAtom = 0 Then
    With MonitorClass
      .style = 0
      .lpfnWndProc = modPrev_Address2Long(AddressOf modPrev_WindowProc)
      .hInstance = App.hInstance
      .lpszClassName = MakeClassName(ProgramID)
    End With
  
    m_hAtom = RegisterClass(MonitorClass)
    If m_hAtom = 0 Then Exit Function 'Fensterklasse konnte nicht registriert werden.
  End If
  
  If m_hWnd = 0 Then
    m_hWnd = CreateWindowEx(WS_EX_APPWINDOW, MakeClassName(ProgramID), ProgramID, _
                            0&, 200&, 200&, 320&, 200&, _
                            0&, 0&, App.hInstance, ByVal 0&)
    If m_hWnd = 0 Then 'Fenster konnte nicht erstellt werden.
      result = UnregisterClass(MakeClassName(ProgramID), App.hInstance)
      If result <> 0 Then 'Fensterklasse konnte deregistriert werden.
        m_hAtom = 0
      End If
    
      Exit Function
    End If
  Else 'Fenster existiert bereits
    Exit Function
  End If
  
  Set modPrev_EventTarget = Me
  
  InstallMonitor = True
End Function

Private Sub UninstallMonitor()
  Dim result As Long
  
  If m_hWnd <> 0 Then 'Monitor-Fenster vorhanden.
    Call SendMessage(m_hWnd, WM_CLOSE, ByVal 0&, ByVal 0&)
    m_hWnd = 0
  End If
    
  result = UnregisterClass(MakeClassName(ProgramID), App.hInstance)
  If result <> 0 Then 'Fensterklasse konnte deregistriert werden.
    m_hAtom = 0
  End If
End Sub

Public Sub SendData(DataType As Long, Data As String)
  Dim CopyData As COPYDATASTRUCT
  Dim B() As Byte
  
  With CopyData
    .dwData = DataType
    .cbData = Len(Data)
    If Len(Data) > 0 Then
      B = StrConv(Data, vbFromUnicode)
      .lpData = VarPtr(B(0))
    Else
      .lpData = 0&
    End If
  End With
  
  If m_PrevHandle <> 0 Then Call SendMessage(m_PrevHandle, WM_COPYDATA, 0&, CopyData)
End Sub

Public Property Get PrevInstance() As Boolean
  PrevInstance = m_PrevInstance
End Property

Public Sub InternalEventRaiser(DataType As Long, Data As String)
Attribute InternalEventRaiser.VB_MemberFlags = "40"
  RaiseEvent DataArrival(DataType, Data)
End Sub

Private Sub Class_Initialize()
  m_PrevHandle = FindMonitor
  
  If m_PrevHandle = 0 Then
    InstallMonitor
  Else
    m_PrevInstance = True
  End If
End Sub

Private Sub Class_Terminate()
  UninstallMonitor
End Sub
