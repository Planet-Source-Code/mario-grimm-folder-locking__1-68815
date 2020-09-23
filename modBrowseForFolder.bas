Attribute VB_Name = "modBrowseForFolder"
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias _
"SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias _
"SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Enum Flag
    '//Nach Drucker im Netzwerk suchen lassen
    BrowseForPrinter = BIF_BROWSEFORPRINTER
    '//Nach Computer im Netzwerk suchen lassen
    BrowseForComputer = BIF_BROWSEFORCOMPUTER
    '//Nach lokalem oder Netzwerkordner suchen lassen
    BrowseForFolder = BIF_RETURNONLYFSDIRS
    '//Nur lokale Ordner zur Auswahl zulassen
    BrowseForLocalFolder = BIF_RETURNFSANCESTORS
End Enum

Public Function ShowSelection(ByVal hWnd As Long, Title As String, _
Optional dwFlag As Flag = BIF_RETURNONLYFSDIRS) As String
    
    Dim bi As BROWSEINFO
    Dim pidl As Long
    Dim strFolder As String
    
    strFolder = String$(255, vbNullChar)
    With bi
        .hOwner = hWnd
        .ulFlags = dwFlag
        .pidlRoot = 0
        .lpszTitle = Title
    End With
    
    pidl = SHBrowseForFolder(bi)
    ShowSelection = IIf(SHGetPathFromIDList(ByVal pidl, ByVal strFolder), _
                       Left$(strFolder, InStr(strFolder, vbNullChar) _
                       - 1), _
                       vbNullString)
    
End Function


