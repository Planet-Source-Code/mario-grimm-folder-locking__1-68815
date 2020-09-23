VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFolderLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Sub LockFolder(sFolder As String, sUser As String, sPWD As String)
Dim F As Integer
On Error GoTo Err_Rep

If InStr(1, sFolder, ".", vbTextCompare) = 0 Then

    sFolder = IIf(Right$(sFolder, 1) = "\", Right$(sFolder, Len(sFolder) - 1), sFolder)
    F = FreeFile
    
    Open sFolder & "\Protect.dat" For Output As #F
        Print #F, sUser
        Print #F, sPWD
    Close #F
    
    Name sFolder As sFolder & AlterExtn
    
    Call MsgBox("The folder has been locked." & vbCrLf & vbCrLf & "You can Unlock it anytime by providing the same Name and PassWord.", vbInformation, "Folder Locked")
    
Else
    Call MsgBox("This folder is already Locked by other users.", vbExclamation, "Access Denied")
End If

Err_Rep:
If Err Then
    Call MsgBox(Err.Description, vbCritical, "Error !")
End If
End Sub


Public Function UnlockFolder(sFolder As String, sUser As String, sPWD As String) As Boolean
On Error GoTo Err_Rep
Dim F As Integer

If InStr(1, sFolder, ".", vbTextCompare) > 0 Then

    Dim TempNm, TempPw As String
    
    sFolder = IIf(Right$(sFolder, 1) = "\", Right$(sFolder, Len(sFolder) - 1), sFolder)
    
    F = FreeFile
    
    Open sFolder & "\Protect.dat" For Input As #F
        Line Input #F, TempNm
        Line Input #F, TempPw
    Close #F
    
    If TempNm <> sUser Or TempPw <> sPWD Then
        Call MsgBox("You are not authorised to Unlock this folder." & vbCrLf & vbCrLf & "This folder has been Locked by other users and can be Unlocked exclusively by the Owner user.", vbExclamation, "Unauthorised intrusion not allowed.")
        Exit Function
    End If
    
    Kill sFolder & "\Protect.dat"
    
    Name sFolder As Mid(sFolder, 1, InStr(1, sFolder, ".", vbTextCompare) - 1)
    
    Call MsgBox("The folder has been unlocked." & vbCrLf & vbCrLf & "You can now access it.", vbInformation, "Folder Unlocked")
    
Else
    Call MsgBox("This folder isn't Locked at all. You are free to access it.", vbExclamation, "Access Denied")
End If
UnlockFolder = True
Err_Rep:
If Err Then
    Call MsgBox(Err.Description, vbCritical, "Error !")
End If

End Function

