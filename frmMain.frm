VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   " Folder locking"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7965
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   7965
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComctlLib.ListView lvwFolder 
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Folder"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Lock/ Unlock"
      Height          =   735
      Index           =   2
      Left            =   3000
      Picture         =   "frmMain.frx":030A
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Remove"
      Height          =   735
      Index           =   1
      Left            =   1560
      Picture         =   "frmMain.frx":0614
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "ADD"
      Height          =   735
      Index           =   0
      Left            =   120
      Picture         =   "frmMain.frx":091E
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cHide            As cHideProccess
Attribute cHide.VB_VarHelpID = -1
Dim cFolder As New clsFolderLock
Dim iMax_ID As Integer
Dim cAccess() As New clsDenyAccess

Dim WithEvents IM As CInstanceMonitor
Attribute IM.VB_VarHelpID = -1


Private Sub IM_DataArrival(DataType As Long, Data As String)
If Data = "" Then
    frmMain.Show
Else
    If (GetAttr(Command$) And vbDirectory) = vbDirectory Then
        LockUnlock Command$, lvwFolder.ListItems.Count
    End If
End If
End Sub
Private Sub cHide_Error(ByVal m_Error As Long, ByVal s_Error As String)

    MsgBox s_Error, vbExclamation, "Error occured"

End Sub


Private Function HidingProcess() As Boolean
    Set cHide = New cHideProccess
    
    With cHide
        If (.Init()) Then
            If (.HideProcess()) Then
                HidingProcess = True
            End If
        End If
    End With
    
    Set cHide = Nothing

End Function

Private Sub cmdChoice_Click(Index As Integer)
Dim sName As String
Dim strFolder As String
Dim i As Integer
Dim lv As ListItem
Select Case Index
    Case 0
        strFolder = ShowSelection(Me.hWnd, "Choose folder", _
        BrowseForLocalFolder)
        If strFolder <> vbNullString Then
            Set lv = lvwFolder.ListItems.Add(, , strFolder)
        End If
        AddFolderToReg strFolder
    Case 1
        For i = lvwFolder.ListItems.Count To 1 Step -1
            If lvwFolder.ListItems(i).Checked = True Then
                lvwFolder.ListItems.Remove i
            End If
        Next
        DeleteRegEntry "Folder"
        For i = 1 To lvwFolder.ListItems.Count
            CreateRegEntry "Folder" & CStr(i), lvwFolder.ListItems(i).Text
        Next
    Case 2
        For i = lvwFolder.ListItems.Count To 1 Step -1
            If lvwFolder.ListItems(i).Checked = True Then
                sName = lvwFolder.ListItems(i).Text
                CheckFolderName sName
                If sName <> "" Then LockUnlock sName, i
            End If
        Next
End Select

End Sub

Private Sub Form_Load()
Dim bVal As Boolean
Dim bEncrypt As Boolean
Dim sUser As String
Dim sPWD As String

ReDim cAccess(100)

Set IM = New CInstanceMonitor

If IM.PrevInstance Then
  IM.SendData Me.hWnd, Command$
  Unload Me
  Exit Sub
  End
End If


bVal = HidingProcess

If Command$ = "-install" Then
    InstallService "Folderlocking", e_ServiceType_Automatic
    WriteContextMenuEntry
    RunService "Folderlocking"
    iMax_ID = ReceiveAllFolder
    Me.Hide
ElseIf Command$ = "" Then
    iMax_ID = ReceiveAllFolder
    Me.Hide
ElseIf Command$ = "-uninstall" Then
    RemoveService "Folderlocking"
ElseIf (GetAttr(Command$) And vbDirectory) = vbDirectory Then
    LockUnlock Command$, lvwFolder.ListItems.Count
    Me.Hide
End If
optimalWidth lvwFolder

End Sub
Public Sub AddFolderToReg(sFolder As String)

Dim i As Integer

CreateRegEntry "Folder" & CStr(iMax_ID), sFolder
iMax_ID = iMax_ID + 1
End Sub
Public Function ReceiveAllFolder() As Integer
Dim i As Integer
Dim sRes As String
Dim lv As ListItem
Do
i = i + 1
sRes = GetRegEntry("Folder" & CStr(i))


If Not sRes = "Error" Then
    Set lv = lvwFolder.ListItems.Add(, , sRes)
    CheckFolderName sRes
    cAccess(lv.Index).DenyAccess sRes
End If
Loop Until sRes = "Error"
ReceiveAllFolder = i
End Function

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form
Dim i As Integer
For i = LBound(cAccess) To UBound(cAccess)
    cAccess(i).AllowAccess
Next
For Each frm In Forms
    Unload frm
Next
End
End Sub


Private Sub LockUnlock(sFolder As String, Index As Integer)
Dim bEncrypt As Boolean
Dim sUser As String
Dim sPWD As String
Dim bRes As Boolean
If InStr(1, sFolder, AlterExtn) Then
    frmPassword.Caption = " Unlocking folder"
    bEncrypt = True
Else
    frmPassword.Caption = " Locking folder"
End If
frmPassword.lblFolderLocation.Caption = sFolder
frmPassword.lblFolderLocation.ToolTipText = sFolder
frmPassword.ShowMe sUser, sPWD
If sUser <> "" And sPWD <> "" Then
    If bEncrypt = False Then
        cFolder.LockFolder sFolder, sUser, sPWD
        cAccess(Index).DenyAccess sFolder & AlterExtn
    Else
        cAccess(Index).AllowAccess
        bRes = cFolder.UnlockFolder(sFolder, sUser, sPWD)
        If bRes = False Then cAccess(Index).DenyAccess sFolder
        
    End If
End If

End Sub
Private Sub CheckFolderName(ByRef sName As String)
Dim i As Integer
If Dir$(sName, vbDirectory) = "" Then
    If Dir$(sName & AlterExtn, vbDirectory) = "" Then
        MsgBox "Folder isn´t available. Probably it has been removed!", vbOKOnly + vbInformation, "Information"
        lvwFolder.ListItems.Remove i
        sName = ""
    Else
        sName = sName & AlterExtn
    End If
End If

End Sub

Private Sub lvwFolder_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim bEncrypt As Boolean
Dim i As Integer
If Data.GetFormat(vbCFFiles) = True Then
    For i = 1 To Data.Files.Count
        If (GetAttr(Data.Files(i)) And vbDirectory) = vbDirectory Then
            lvwFolder.ListItems.Add , , Data.Files(i)
            AddFolderToReg Data.Files(i)
        End If
    Next
End If
End Sub
