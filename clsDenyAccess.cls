VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDenyAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Const OFS_MAXPATHNAME = 128
Const OF_CREATE = &H1000
Const OF_READ = &H0
Const OF_WRITE = &H1
Private Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type
Dim m_Hwnd As Long
Public Sub DenyAccess(sFolder As String)
Dim OF As OFSTRUCT
m_Hwnd = OpenFile(sFolder & "\Protect.dat", OF, OF_READ)
End Sub
Public Sub AllowAccess()
CloseHandle m_Hwnd
m_Hwnd = 0
End Sub
