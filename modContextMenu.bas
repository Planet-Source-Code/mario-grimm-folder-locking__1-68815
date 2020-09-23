Attribute VB_Name = "modContextMenu"
Option Explicit
Public Sub WriteContextMenuEntry()
Dim EXEpath As String
Dim Path As String
CreateKey "HKEY_CLASSES_ROOT\CLSID\" & AlterExtn & ""
CreateKey "HKEY_CLASSES_ROOT\CLSID\" & AlterExtn & "\DefaultIcon"
CreateKey "HKEY_CLASSES_ROOT\CLSID\" & AlterExtn & "\InProcServer32"
CreateKey "HKEY_CLASSES_ROOT\CLSID\" & AlterExtn & "\PersistentAddinsRegistered"
CreateKey "HKEY_CLASSES_ROOT\CLSID\" & AlterExtn & "\PersistentAddinsRegistered\{89BCB740-6119-101A-BCB7-00DD010655AF}"
CreateKey "HKEY_CLASSES_ROOT\CLSID\" & AlterExtn & "\PersistentHandler"
CreateKey "HKEY_CLASSES_ROOT\CLSID\" & AlterExtn & "\ProgID"
CreateKey "HKEY_CLASSES_ROOT\CLSID\" & AlterExtn & "\shellex"
CreateKey "HKEY_CLASSES_ROOT\CLSID\" & AlterExtn & "\shellex\MayChangeDefaultMenu"
EXEpath = "" + Chr(34) + "" + App.Path + "\" + App.EXEName + ".exe" + Chr(34) + " %1"
SetStringValue "HKEY_CLASSES_ROOT\CLSID\" & AlterExtn & "\DefaultIcon", "", "" + App.Path + "\" + App.EXEName + ".exe ,0"
SetStringValue "HKEY_CLASSES_ROOT\CLSID\" & AlterExtn & "\InProcServer32", "", "shell32.dll"
SetStringValue "HKEY_CLASSES_ROOT\CLSID\" & AlterExtn & "\InProcServer32", "ThreadingModel", "Apartment"
SetStringValue "HKEY_CLASSES_ROOT\CLSID\" & AlterExtn & "\PersistentAddinsRegistered\{89BCB740-6119-101A-BCB7-00DD010655AF}", "", "" & AlterExtn & ""
SetStringValue "HKEY_CLASSES_ROOT\CLSID\" & AlterExtn & "\PersistentHandler", "", "" & AlterExtn & ""
SetStringValue "HKEY_CLASSES_ROOT\CLSID\" & AlterExtn & "\ProgID", "", "LockedFolder"

CreateKey "HKEY_CLASSES_ROOT\LockedFolder"
CreateKey "HKEY_CLASSES_ROOT\LockedFolder\DefaultIcon"
CreateKey "HKEY_CLASSES_ROOT\LockedFolder\Shell"
CreateKey "HKEY_CLASSES_ROOT\LockedFolder\Shell\Unlock"
CreateKey "HKEY_CLASSES_ROOT\LockedFolder\Shell\Unlock\command"
CreateKey "HKEY_CLASSES_ROOT\LockedFolder\shellex"
CreateKey "HKEY_CLASSES_ROOT\LockedFolder\shellex\ContextMenuHandlers"
CreateKey "HKEY_CLASSES_ROOT\LockedFolder\shellex\ContextMenuHandlers\" & AlterExtn & ""

SetStringValue "HKEY_CLASSES_ROOT\LockedFolder\DefaultIcon", "", "" + App.Path + "\" + App.EXEName + ".exe ,0"
SetStringValue "HKEY_CLASSES_ROOT\LockedFolder\Shell\Unlock\command", "", "" & EXEpath

CreateKey "HKEY_CLASSES_ROOT\Directory\Shell\Locking Folder"
CreateKey "HKEY_CLASSES_ROOT\Directory\Shell\Locking Folder\Command"
SetStringValue "HKEY_CLASSES_ROOT\Directory\Shell\Locking Folder\Command", "", "" & EXEpath

End Sub


Public Sub DeleteRegEntry(sKey As String)
DeleteKey "HKEY_CURRENT_USER\Software\ATS\Folder"

End Sub
Public Sub CreateRegEntry(sKey As String, sVal As String)
CreateKey "HKEY_CURRENT_USER\Software\ATS"
CreateKey "HKEY_CURRENT_USER\Software\ATS\Folder"
SetStringValue "HKEY_CURRENT_USER\Software\ATS\Folder", sKey, sVal
End Sub
Public Function GetRegEntry(sKey As String) As String
GetRegEntry = GetStringValue("HKEY_CURRENT_USER\Software\ATS\Folder", sKey)
End Function

