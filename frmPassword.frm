VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   " Folderlocking"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.TextBox txtInfo 
      Appearance      =   0  '2D
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  '2D
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin prjFolderLocking.isButton isOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Icon            =   "frmPassword.frx":0000
      Style           =   8
      Caption         =   "&OK"
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin prjFolderLocking.isButton isCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Icon            =   "frmPassword.frx":001C
      Style           =   8
      Caption         =   "&Abbrechen"
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin VB.Label lblFolderLocation 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3015
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   690
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   420
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s1 As String
Dim s2 As String
Public Function ShowMe(ByRef sName As String, ByRef sPWD As String) As Boolean
On Error GoTo ErrHandling
Me.Show vbModal
sName = s1
sPWD = s2
ShowMe = True

ErrHandling:
End Function

Private Sub Form_GotFocus()
txtInfo(0).SetFocus
End Sub

Private Sub isCancel_Click()
s1 = ""
s1 = ""
Unload Me
End Sub

Private Sub isOK_Click()
If txtInfo(0).Text = "" Or txtInfo(1).Text = "" Then
    MsgBox " Type in username AND password", vbOKOnly + vbInformation
    Exit Sub
End If
s1 = txtInfo(0).Text
s2 = txtInfo(1).Text
Unload Me
End Sub

