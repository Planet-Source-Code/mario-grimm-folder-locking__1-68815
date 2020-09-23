Attribute VB_Name = "modService"
Option Explicit
Private Const SERVICE_WIN32_OWN_PROCESS = &H10&
Private Const SERVICE_WIN32_SHARE_PROCESS = &H20&
Private Const SERVICE_WIN32 = SERVICE_WIN32_OWN_PROCESS + SERVICE_WIN32_SHARE_PROCESS
Private Const SERVICE_ACCEPT_STOP = &H1
Private Const SERVICE_ACCEPT_PAUSE_CONTINUE = &H2
Private Const SERVICE_ACCEPT_SHUTDOWN = &H4
Private Const SC_MANAGER_CONNECT = &H1
Private Const SC_MANAGER_CREATE_SERVICE = &H2
Private Const SC_MANAGER_ENUMERATE_SERVICE = &H4
Private Const SC_MANAGER_LOCK = &H8
Private Const SC_MANAGER_QUERY_LOCK_STATUS = &H10
Private Const SC_MANAGER_MODIFY_BOOT_CONFIG = &H20
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SERVICE_QUERY_CONFIG = &H1
Private Const SERVICE_CHANGE_CONFIG = &H2
Private Const SERVICE_QUERY_STATUS = &H4
Private Const SERVICE_ENUMERATE_DEPENDENTS = &H8
Private Const SERVICE_START = &H10
Private Const SERVICE_STOP = &H20
Private Const SERVICE_PAUSE_CONTINUE = &H40
Private Const SERVICE_INTERROGATE = &H80
Private Const SERVICE_USER_DEFINED_CONTROL = &H100
Private Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)
Private Const SERVICE_DISABLED As Long = &H4
Private Const SERVICE_DEMAND_START As Long = &H3
Private Const SERVICE_AUTO_START As Long = &H2
Private Const SERVICE_SYSTEM_START As Long = &H1
Private Const SERVICE_BOOT_START As Long = &H0


Public Enum e_ServiceType
    e_ServiceType_Disabled = 4
    e_ServiceType_Manual = 3
    e_ServiceType_Automatic = 2
    e_ServiceType_SystemStart = 1
    e_ServiceType_BootTime = 0
End Enum
Private Const SERVICE_ERROR_NORMAL As Long = &H1


Private Enum SERVICE_CONTROL
    SERVICE_CONTROL_STOP = &H1
    SERVICE_CONTROL_PAUSE = &H2
    SERVICE_CONTROL_CONTINUE = &H3
    SERVICE_CONTROL_INTERROGATE = &H4
    SERVICE_CONTROL_SHUTDOWN = &H5
End Enum


Private Enum SERVICE_STATE
    SERVICE_STOPPED = &H1
    SERVICE_START_PENDING = &H2
    SERVICE_STOP_PENDING = &H3
    SERVICE_RUNNING = &H4
    SERVICE_CONTINUE_PENDING = &H5
    SERVICE_PAUSE_PENDING = &H6
    SERVICE_PAUSED = &H7
End Enum


Private Type SERVICE_TABLE_ENTRY
    lpServiceName As String
    lpServiceProc As Long
    lpServiceNameNull As Long
    lpServiceProcNull As Long
    End Type


Private Type SERVICE_STATUS
    dwServiceType As Long
    dwCurrentState As Long
    dwControlsAccepted As Long
    dwWin32ExitCode As Long
    dwServiceSpecificExitCode As Long
    dwCheckPoint As Long
    dwWaitHint As Long
    End Type


Private Declare Function StartServiceCtrlDispatcher Lib "advapi32.dll" Alias "StartServiceCtrlDispatcherA" (lpServiceStartTable As SERVICE_TABLE_ENTRY) As Long


Private Declare Function RegisterServiceCtrlHandler Lib "advapi32.dll" Alias "RegisterServiceCtrlHandlerA" (ByVal lpServiceName As String, ByVal lpHandlerProc As Long) As Long


Private Declare Function SetServiceStatus Lib "advapi32.dll" (ByVal hServiceStatus As Long, lpServiceStatus As SERVICE_STATUS) As Long


Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long


Private Declare Function CreateService Lib "advapi32.dll" Alias "CreateServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal lpDisplayName As String, ByVal dwDesiredAccess As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, ByVal lpdwTagId As String, ByVal lpDependencies As String, ByVal lp As String, ByVal lpPassword As String) As Long


Private Declare Function DeleteService Lib "advapi32.dll" (ByVal hService As Long) As Long


Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long


Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
    Private hServiceStatus As Long
    Private ServiceStatus As SERVICE_STATUS
    Dim SERVICE_NAME As String


Public Sub InstallService(serviceName As String, serviceType As e_ServiceType)
    Dim hSCManager As Long
    Dim hService As Long
    Dim cmd As String
    Dim lServiceType As Long
    
    
    


    Select Case serviceType
        Case e_ServiceType_Automatic
        lServiceType = SERVICE_AUTO_START
        Case e_ServiceType_BootTime
        lServiceType = SERVICE_BOOT_START
        Case e_ServiceType_Disabled
        lServiceType = SERVICE_DISABLED
        Case e_ServiceType_Manual
        lServiceType = SERVICE_DEMAND_START
        Case e_ServiceType_SystemStart
        lServiceType = SERVICE_SYSTEM_START
    End Select

hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CREATE_SERVICE)
hService = CreateService(hSCManager, serviceName, serviceName, SERVICE_ALL_ACCESS, SERVICE_WIN32_OWN_PROCESS, lServiceType, SERVICE_ERROR_NORMAL, App.Path & "\" & App.EXEName, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
CloseServiceHandle hService
CloseServiceHandle hSCManager
End Sub


Public Sub RemoveService(serviceName As String)
    Dim hSCManager As Long
    Dim hService As Long
    Dim cmd As String
    hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CREATE_SERVICE)
    hService = OpenService(hSCManager, serviceName, SERVICE_ALL_ACCESS)
    DeleteService hService
    CloseServiceHandle hService
    CloseServiceHandle hSCManager
End Sub


Public Function RunService(serviceName As String) As Boolean
    Dim ServiceTableEntry As SERVICE_TABLE_ENTRY
    Dim b As Boolean
    ServiceTableEntry.lpServiceName = serviceName
    SERVICE_NAME = serviceName
    ServiceTableEntry.lpServiceProc = FncPtr(AddressOf ServiceMain)
    b = StartServiceCtrlDispatcher(ServiceTableEntry)
    RunService = b
End Function


Private Sub ServiceMain(ByVal dwArgc As Long, ByVal lpszArgv As Long)
    Dim b As Boolean
    'Set initial state
    ServiceStatus.dwServiceType = SERVICE_WIN32_OWN_PROCESS
    ServiceStatus.dwCurrentState = SERVICE_START_PENDING
    ServiceStatus.dwControlsAccepted = SERVICE_ACCEPT_STOP Or SERVICE_ACCEPT_PAUSE_CONTINUE Or SERVICE_ACCEPT_SHUTDOWN
    ServiceStatus.dwWin32ExitCode = 0
    ServiceStatus.dwServiceSpecificExitCode = 0
    ServiceStatus.dwCheckPoint = 0
    ServiceStatus.dwWaitHint = 0
    hServiceStatus = RegisterServiceCtrlHandler(SERVICE_NAME, AddressOf Handler)
    ServiceStatus.dwCurrentState = SERVICE_START_PENDING
    b = SetServiceStatus(hServiceStatus, ServiceStatus)
    ServiceStatus.dwCurrentState = SERVICE_RUNNING
    b = SetServiceStatus(hServiceStatus, ServiceStatus)
End Sub


Private Sub Handler(ByVal fdwControl As Long)
    Dim b As Boolean
    


    Select Case fdwControl
        Case SERVICE_CONTROL_PAUSE
        ServiceStatus.dwCurrentState = SERVICE_PAUSED
        Case SERVICE_CONTROL_CONTINUE
        ServiceStatus.dwCurrentState = SERVICE_RUNNING
        Case SERVICE_CONTROL_STOP
        ServiceStatus.dwWin32ExitCode = 0
        ServiceStatus.dwCurrentState = SERVICE_STOP_PENDING
        ServiceStatus.dwCheckPoint = 0
        ServiceStatus.dwWaitHint = 0
        b = SetServiceStatus(hServiceStatus, ServiceStatus)
        ServiceStatus.dwCurrentState = SERVICE_STOPPED
        Case SERVICE_CONTROL_INTERROGATE
        Case Else
    End Select
b = SetServiceStatus(hServiceStatus, ServiceStatus)
End Sub


Function FncPtr(ByVal fnp As Long) As Long
    FncPtr = fnp
End Function

