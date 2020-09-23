VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDeviceDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ********************************
' *            Events            *
' ********************************
Event Error(ByVal m_Error As e_Errors)


' ********************************
' *            Types             *
' ********************************
Private Type SERVICE_STATUS
    dwServiceType As Long
    dwCurrentState As Long
    dwControlsAccepted As Long
    dwWin32ExitCode As Long
    dwServiceSpecificExitCode As Long
    dwCheckPoint As Long
    dwWaitHint As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type DRIVER_STACK
    dsOSMajor       As Long
    dsOSMinor       As Long
    dsResourceID    As Long
    dsResourceType  As String
End Type


' ********************************
' *            Enums             *
' ********************************
Public Enum e_Errors
    DD_ERROR_SUCCESS
    DD_ERROR_SYS_DOES_NOT_EXIST
    DD_ERROR_COPY_FAIL
    DD_ERROR_SERVICE_UNKNOWN
    DD_ERROR_INSUFFICIENT_RIGHTS
    DD_ERROR_INVALID_NAME
    DD_ERROR_DRIVER_INACCESSIBLE
    DD_ERROR_INVALID_RESOURCE_DRIVER
    DD_ERROR_OS_NOT_SUPPORTED
End Enum

Public Enum e_ErrorNotification
    EVENT_BASED
    ERROR_RAISED
    MESSAGE_BOX_BASED
    ERR_NOTIFY_INVALID_TYPE
End Enum


' ********************************
' *          Constants           *
' ********************************
Private Const MAX_PATH                      As Long = 260
Private Const DEVICE_DRIVER_EXTENTION       As String = ".sys"
Private Const DD_ERROR_BASE                 As Long = &H1000
Private Const DEFAULT_RESOURCE_TYPE         As String = "custom"

Private Const VER_PLATFORM_WIN32_NT         As Long = 2

Private Const SC_MANAGER_ALL_ACCESS         As Long = &HF003F
Private Const SERVICE_ALL_ACCESS            As Long = &HF01FF
Private Const SERVICE_KERNEL_DRIVER         As Long = 1
Private Const SERVICE_DEMAND_START          As Long = 3
Private Const SERVICE_ERROR_NORMAL          As Long = 1
Private Const SERVICE_CONTROL_STOP          As Long = 1

Private Const ERROR_SERVICE_EXISTS          As Long = &H431
Private Const ERROR_ACCESS_DENIED           As Long = 5
Private Const ERROR_INVALID_NAME            As Long = &H7B
Private Const ERROR_SERVICE_DOES_NOT_EXIST  As Long = &H424
Private Const ERROR_SERVICE_ALREADY_RUNNING As Long = &H420

Private Const OPEN_EXISTING                 As Long = 3
Private Const GENERIC_READ                  As Long = &H80000000
Private Const GENERIC_WRITE                 As Long = &H40000000
Private Const FILE_ATTRIBUTE_SYSTEM         As Long = &H4
Private Const INVALID_HANDLE_VALUE          As Long = -1


' ********************************
' *      Private variables       *
' ********************************
Private m_DeviceDriver_FilePath             As String           ' sys full file path
Private m_DeviceDriver_FileName             As String           ' sys file name
Private m_DeviceDriver_NameOfService        As String           ' service name
Private m_DeviceDriver_NameOfDisplay        As String           ' display service name
Private m_DeviceDriver_SysFullFileName      As String           ' sys in system32\drivers\...

Private m_DeviceDriver_Handle               As Long
Private m_OS_Major                          As Long
Private m_OS_Minor                          As Long
Private b_IsWinNT                           As Boolean
Private m_ExtendedFlag                      As Boolean
Private m_DriverStack()                     As DRIVER_STACK
Private m_DriverStackC                      As Long
Private m_ErrorNotificationMethod           As e_ErrorNotification


' ********************************
' *      API declarations        *
' ********************************
Private Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function GetLastError Lib "kernel32.dll" () As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFilename As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function CreateService Lib "advapi32.dll" Alias "CreateServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal lpDisplayName As String, ByVal dwDesiredAccess As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, ByRef lpdwTagId As Long, ByVal lpDependencies As String, ByVal lp As String, ByVal lpPassword As String) As Long
Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
Private Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
Private Declare Function ControlService Lib "advapi32.dll" (ByVal hService As Long, ByVal dwControl As Long, ByRef lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function DeleteService Lib "advapi32.dll" (ByVal hService As Long) As Long


Private Function InstallDriver() As Boolean

    Dim SchSCManager            As Long, _
        schService              As Long
    
    
    If pvKillFile(m_DeviceDriver_SysFullFileName) Then
        If CopyFile(m_DeviceDriver_FilePath, m_DeviceDriver_SysFullFileName, False) = 0 Then
            pError DD_ERROR_COPY_FAIL
            Exit Function
        End If
    End If
    
    If Not pvFileExists(m_DeviceDriver_SysFullFileName) Then
        pError DD_ERROR_COPY_FAIL
        Exit Function
    End If
    
    SchSCManager = OpenSCManager(vbNullString, _
                                 vbNullString, _
                                 SC_MANAGER_ALL_ACCESS)
    
    schService = CreateService(SchSCManager, _
                               m_DeviceDriver_NameOfService, _
                               m_DeviceDriver_NameOfDisplay, _
                               SERVICE_ALL_ACCESS, _
                               SERVICE_KERNEL_DRIVER, _
                               SERVICE_DEMAND_START, _
                               SERVICE_ERROR_NORMAL, _
                               m_DeviceDriver_SysFullFileName, _
                               vbNullString, _
                               ByVal 0, _
                               vbNullString, _
                               vbNullString, _
                               vbNullString)
    
    If (schService = 0) Then
    
        Select Case (GetLastError())
        
            Case ERROR_SERVICE_EXISTS
                GoTo con:
                
            Case Else
                pError DD_ERROR_SERVICE_UNKNOWN
                
        End Select
        
        Call CloseServiceHandle(SchSCManager)
        
        Exit Function
        
    End If
    
con:
    Call CloseServiceHandle(schService)
    Call CloseServiceHandle(SchSCManager)
    
    InstallDriver = True

End Function

Private Function StartDriver() As Boolean

    Dim SchSCManager            As Long, _
        schService              As Long, _
        ret                     As Long

    
    SchSCManager = OpenSCManager(vbNullString, _
                                 vbNullString, _
                                 SC_MANAGER_ALL_ACCESS)
    
    If (SchSCManager = 0) Then
        pError DD_ERROR_INSUFFICIENT_RIGHTS
        Exit Function
    End If
    
    Do
        schService = OpenService(SchSCManager, _
                                 m_DeviceDriver_NameOfService, _
                                 SERVICE_ALL_ACCESS)
        
        If (schService = 0) Then
        
            Select Case (GetLastError())
            
                Case ERROR_ACCESS_DENIED
                    pError DD_ERROR_INSUFFICIENT_RIGHTS
                
                Case ERROR_INVALID_NAME
                    pError DD_ERROR_INVALID_NAME
                
                Case ERROR_SERVICE_DOES_NOT_EXIST, 0
                    If InstallDriver() Then
                        GoTo con_loop:
                    End If
                
                Case Else
                    pError DD_ERROR_SERVICE_UNKNOWN
                    
            End Select
            
            Call CloseServiceHandle(SchSCManager)
            Exit Function
            
        End If
con_loop:
    Loop While (schService = 0)
    
    ret = StartService(schService, _
                       0, _
                       0)
    
    If (ret = 0) Then
        If (GetLastError() <> ERROR_SERVICE_ALREADY_RUNNING) Then
            
            pError DD_ERROR_SERVICE_UNKNOWN
            GoTo last_end:
            
        End If
    End If
    
    StartDriver = True
    
last_end:
    Call CloseServiceHandle(schService)
    Call CloseServiceHandle(SchSCManager)

End Function

Private Function pvGetSysDir() As String
    
    Dim lpos As Long
    
    
    pvGetSysDir = Space(MAX_PATH)
    
    lpos = GetSystemDirectory(pvGetSysDir, _
                              MAX_PATH)
    
    If lpos Then pvGetSysDir = Mid(pvGetSysDir, 1, lpos)
    
End Function

Private Function pvSlash(ByVal m_Path As String) As String

    If (m_Path Like "*\") Then
        pvSlash = m_Path
    Else
        pvSlash = m_Path & "\"
    End If
    
End Function

Private Sub pError(ByVal m_Error As e_Errors)
    
    Select Case ErrorNotificationMethod
        
        Case e_ErrorNotification.EVENT_BASED
            RaiseEvent Error(m_Error)
        
        Case e_ErrorNotification.ERROR_RAISED
            Err.Raise DD_ERROR_BASE + m_Error, _
                      "cDeviceDriver", _
                      GetErrorMessage(m_Error)
    
        Case e_ErrorNotification.MESSAGE_BOX_BASED
            MsgBox "Source module:  Device Driver Support Class Module" & vbCrLf & _
                   "Error number:   " & m_Error & vbCrLf & _
                   "Error message:  " & GetErrorMessage(m_Error), _
                   vbExclamation, _
                   "An error has occured"
    
    End Select
    
End Sub



Private Function pvFileExists(ByVal m_Path As String) As Boolean

    On Error GoTo er_pvFileExists:

    Call FileLen(m_Path)
    pvFileExists = True

er_pvFileExists:
End Function

Private Function pvKillFile(ByVal m_Path As String) As Boolean
    
    On Error GoTo er_pvKillFile:
    
    If (Not pvFileExists(m_Path)) Then
        
        pvKillFile = True
        Exit Function
        
    End If
    
    SetAttr m_Path, vbNormal
    Kill m_Path
    
    pvKillFile = True
    
er_pvKillFile:
End Function

Public Function InstallDeviceDriver( _
            ByVal m_SysFileName As String, _
            ByVal m_ServiceName As String, _
            ByRef m_Driver_Handle As Long, _
            Optional ByVal m_DisplayName As String) As Boolean

    
    If (Not b_IsWinNT) Then
        
        pError DD_ERROR_OS_NOT_SUPPORTED
        Exit Function
        
    End If
    
    If Not pvFileExists(m_SysFileName) Then
        pError DD_ERROR_SYS_DOES_NOT_EXIST
        Exit Function
    End If
    
    m_DeviceDriver_FilePath = m_SysFileName
    m_DeviceDriver_FileName = pvGetFileTitle(m_SysFileName)
    m_DeviceDriver_NameOfService = m_ServiceName
    m_DeviceDriver_SysFullFileName = pvSlash(pvGetSysDir()) & "Drivers\" & _
                                     m_DeviceDriver_FileName & DEVICE_DRIVER_EXTENTION
    
    If Len(m_DisplayName) = 0 Then
        m_DeviceDriver_NameOfDisplay = m_ServiceName
    End If
    
    m_DeviceDriver_Handle = CreateFile("\\.\" & m_DeviceDriver_FileName, _
                                       GENERIC_READ Or GENERIC_WRITE, _
                                       0, _
                                       0, _
                                       OPEN_EXISTING, _
                                       FILE_ATTRIBUTE_SYSTEM, _
                                       0)
    
    If (m_DeviceDriver_Handle = INVALID_HANDLE_VALUE) Then
    
        If (Not StartDriver()) Then
            Exit Function
        End If
        
        m_DeviceDriver_Handle = CreateFile("\\.\" & m_DeviceDriver_FileName, _
                                           GENERIC_READ Or GENERIC_WRITE, _
                                           0, _
                                           0, _
                                           OPEN_EXISTING, _
                                           FILE_ATTRIBUTE_SYSTEM, _
                                           0)
        
        If (m_DeviceDriver_Handle = INVALID_HANDLE_VALUE) Then
        
            pError DD_ERROR_DRIVER_INACCESSIBLE
            Exit Function
            
        End If
        
    End If
    
    m_Driver_Handle = m_DeviceDriver_Handle
    InstallDeviceDriver = True
    
End Function

Public Function UninstallDeviceDriver() As Boolean
    
    Dim SchSCManager            As Long, _
        schService              As Long, _
        ret                     As Long
    
    Dim serviceStatus           As SERVICE_STATUS

    
    If (Not b_IsWinNT) Then
        
        pError DD_ERROR_OS_NOT_SUPPORTED
        Exit Function
        
    End If
    
    If (m_DeviceDriver_Handle) Then
        CloseHandle m_DeviceDriver_Handle
    End If
        
    SchSCManager = OpenSCManager(vbNullString, _
                                 vbNullString, _
                                 SC_MANAGER_ALL_ACCESS)
        
    schService = OpenService(SchSCManager, _
                             m_DeviceDriver_NameOfService, _
                             SERVICE_ALL_ACCESS)
        
    If (schService = 0) Then
        GoTo contin:
    End If
        
    ret = ControlService(schService, _
                         SERVICE_CONTROL_STOP, _
                         serviceStatus)
        
    ret = DeleteService(schService)
        
    Call CloseServiceHandle(schService)
    
contin:

    Call CloseServiceHandle(SchSCManager)
    pvKillFile m_DeviceDriver_SysFullFileName
    pvReset
    
done:
    UninstallDeviceDriver = True

End Function

Private Sub pvReset()

    m_DeviceDriver_FilePath = ""
    m_DeviceDriver_FileName = ""
    m_DeviceDriver_NameOfService = ""
    m_DeviceDriver_NameOfDisplay = ""
    m_DeviceDriver_Handle = 0
    
End Sub

Private Sub Class_Initialize()

    m_ErrorNotificationMethod = ERROR_RAISED
    m_ExtendedFlag = False
    Call pvGetOsVersion

End Sub

Private Sub Class_Terminate()

    Call UninstallDeviceDriver

End Sub



Private Function pvGetFileTitle(ByVal m_Path As String) As String
    
    Dim sext    As String
    Dim sfp     As String
    
    sext = Right(m_Path, Len(m_Path) - InStrRev(m_Path, "."))
    sfp = Mid(m_Path, 1, InStrRev(m_Path, "\"))
    
    pvGetFileTitle = Replace(m_Path, "." & sext, "")
    pvGetFileTitle = Replace(pvGetFileTitle, sfp, "")
    
End Function

Public Function InstallDriverResource( _
            ByVal m_ResourceID As Long, _
            ByVal m_TempFileTitle As String, _
            ByVal m_ServiceName As String, _
            ByRef m_Driver_Handle As Long, _
            Optional ByVal m_DisplayName As String, _
            Optional ByVal m_ResourceType As String = DEFAULT_RESOURCE_TYPE) As Boolean
    
    
    On Error GoTo er_InstallDriverResource:
    
    Dim sTempTarget         As String
    Dim DriverIndex         As Long
    
    
    If (Not b_IsWinNT) Then
        
        pError DD_ERROR_OS_NOT_SUPPORTED
        Exit Function
        
    End If
    
    If (ExtendedHandle) Then
        
        DriverIndex = pvFindDriverInformation(m_OS_Major, m_OS_Minor)
        
        If (Not (DriverIndex = INVALID_HANDLE_VALUE)) Then
        
            m_ResourceID = m_DriverStack(DriverIndex).dsResourceID
            m_ResourceType = m_DriverStack(DriverIndex).dsResourceType
            
        End If
        
    End If
    
    If (m_ResourceID = 0) Then
        Exit Function
    End If
    
    sTempTarget = pvSlash(pvGetTempDir()) & m_TempFileTitle
    
    If Not (pvExtractResBinary(sTempTarget, _
                               m_ResourceID, _
                               m_ResourceType)) Then
                               
        pError DD_ERROR_INVALID_RESOURCE_DRIVER
        Exit Function
        
    End If
    
    If Not pvFileExists(sTempTarget) Then
        pError DD_ERROR_INVALID_RESOURCE_DRIVER
        Exit Function
    End If
    
    If Not InstallDeviceDriver(sTempTarget, _
                               m_ServiceName, _
                               m_Driver_Handle, _
                               m_DisplayName) Then
                               
        GoTo er_InstallDriverResource:
        
    End If
    
    InstallDriverResource = True
    
er_InstallDriverResource:
    pvKillFile sTempTarget
    
End Function
Private Function pvExtractResBinary( _
            ByVal m_Destination As String, _
            ByVal m_ResID As Long, _
            ByVal m_ResType As String) As Boolean
    
    On Error GoTo er_pvExtractResBinary:
    
    Dim i           As Integer
    Dim Data()      As Byte
    Dim bOpen       As Boolean
    
    
    i = FreeFile
    
    Data = LoadResData(m_ResID, m_ResType)

    Open m_Destination For Binary Access Write As #i
        bOpen = True
        Put #i, , Data
    Close #i
    
    bOpen = False
    
    Erase Data()
    
    pvExtractResBinary = True
    
er_pvExtractResBinary:
    If bOpen Then Close #i
    
End Function

Private Function pvGetTempDir() As String
    
    Dim lpos As Long
    
    
    pvGetTempDir = Space(260)
    lpos = GetTempPath(260, pvGetTempDir)
    If lpos Then pvGetTempDir = Mid(pvGetTempDir, 1, lpos)
    
End Function

Public Property Get ErrorNotificationMethod() As e_ErrorNotification

    ErrorNotificationMethod = m_ErrorNotificationMethod
    
End Property

Public Property Let ErrorNotificationMethod(ByVal vNewValue As e_ErrorNotification)

    If vNewValue < 0 Or vNewValue >= ERR_NOTIFY_INVALID_TYPE Then
        Exit Property
    Else
        m_ErrorNotificationMethod = vNewValue
    End If
    
End Property

Public Function GetErrorMessage(ByVal m_Error_Number As Long)
    
    If m_Error_Number >= DD_ERROR_BASE Then
        m_Error_Number = m_Error_Number - DD_ERROR_BASE
    End If
    
    Select Case m_Error_Number
    
        Case e_Errors.DD_ERROR_COPY_FAIL
            GetErrorMessage = "Cannot copy the device driver file to your system32\drivers folder."
    
        Case e_Errors.DD_ERROR_DRIVER_INACCESSIBLE
            GetErrorMessage = "The device driver is not accessible (invalid service name?)"
    
        Case e_Errors.DD_ERROR_INSUFFICIENT_RIGHTS
            GetErrorMessage = "Insufficient rights. You must be an administrator."
    
        Case e_Errors.DD_ERROR_INVALID_NAME
            GetErrorMessage = "Invalid device driver service name."
    
        Case e_Errors.DD_ERROR_INVALID_RESOURCE_DRIVER
            GetErrorMessage = "Specified resource data could not be loaded."
    
        Case e_Errors.DD_ERROR_SERVICE_UNKNOWN
            GetErrorMessage = "Please check if the device driver is in your system32\drivers folder."
    
        Case e_Errors.DD_ERROR_SUCCESS
            GetErrorMessage = "Operation successfully completed."
    
        Case e_Errors.DD_ERROR_SYS_DOES_NOT_EXIST
            GetErrorMessage = "Source device driver file could not be found."
        
        Case e_Errors.DD_ERROR_OS_NOT_SUPPORTED
            GetErrorMessage = "Unsupported windows version. Only Windows NT and higher are supported."
    
        Case Else
            GetErrorMessage = "Unknown: (" & Hex(m_Error_Number) & ")"
        
    End Select
    
End Function

Private Sub pvGetOsVersion()

    Dim osvi                As OSVERSIONINFO
    Dim bOsVersionInfo      As Boolean
    
    
    osvi.dwOSVersionInfoSize = Len(osvi)
    bOsVersionInfo = GetVersionEx(osvi)
    
    If (Not bOsVersionInfo) Then
    
        pError DD_ERROR_OS_NOT_SUPPORTED
        Exit Sub
        
    End If
    
    Select Case (osvi.dwPlatformId)
        
        Case VER_PLATFORM_WIN32_NT
    '   {
            b_IsWinNT = True
            m_OS_Major = osvi.dwMajorVersion
            m_OS_Minor = osvi.dwMinorVersion
    '   }
        Case Else
        
            pError DD_ERROR_OS_NOT_SUPPORTED
        
    End Select

End Sub


Public Property Get ExtendedHandle() As Boolean
    
    ExtendedHandle = m_ExtendedFlag
    
End Property

Public Property Let ExtendedHandle(ByVal vNewValue As Boolean)

    m_ExtendedFlag = vNewValue

End Property

Public Function AddDriverInformation( _
                    ByVal dwOS_Major As Long, _
                    ByVal dwOS_Minor As Long, _
                    ByVal d_ResourceID As Long, _
                    Optional d_ResourceType As String = DEFAULT_RESOURCE_TYPE) As Boolean

    
    ReDim Preserve m_DriverStack(m_DriverStackC) As DRIVER_STACK

    With m_DriverStack(m_DriverStackC)
        .dsOSMajor = dwOS_Major
        .dsOSMinor = dwOS_Minor
        .dsResourceID = d_ResourceID
        .dsResourceType = d_ResourceType
    End With
    
    m_DriverStackC = m_DriverStackC + 1
    
End Function

Private Function pvFindDriverInformation( _
                    ByVal dwMajor As Long, _
                    ByVal dwMinor As Long) As Long
    
    
    Dim i   As Long
    
    
    pvFindDriverInformation = INVALID_HANDLE_VALUE
    
    For i = 0 To m_DriverStackC - 1
        
        With m_DriverStack(i)
            If (.dsOSMajor = dwMajor And .dsOSMinor = dwMinor) Then
                
                pvFindDriverInformation = i
                Exit Function
                
            End If
        End With
        
    Next i
    
End Function
