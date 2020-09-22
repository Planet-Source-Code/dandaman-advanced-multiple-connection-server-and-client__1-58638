VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Connect"
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   3105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Join"
      Height          =   615
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   450
      Width           =   1575
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   720
      MaxLength       =   30
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "IP:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub InitCommonControls Lib "comctl32" ()    'for the manifestfile

Const strDefaultIP As String = "127.0.0.1"
Const iRemotePort As Integer = 8000

Private Sub cmdConnect_Click()
If Trim(txtName.text) = "" Then Msgbox2 "Name cannot be blank.": Exit Sub
'check for inappropiate symbols in the name
If InStr(1, txtName.text, "*") <> 0 Then MsgBox "Name cannot have *'s in it.": Exit Sub
If InStr(1, txtName.text, ":") <> 0 Then MsgBox "Name cannot have :'s in it.": Exit Sub

If Trim(txtIP.text) = "" Then Msgbox2 "IP cannot be blank.": Exit Sub

Call SaveSetting(AppName, "main", "name", txtName.text)
Call SaveSetting(AppName, "main", "ip", txtIP.text)


frmMain.sockMain.Close
frmMain.sockMain.RemotePort = iRemotePort
Client.Name = txtName.text
frmMain.sockMain.Connect txtIP.text
End Sub

Private Sub cmdConnect_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then 'they hit enter, either in one of the textboxes, or here
    Call cmdConnect_Click
    'this will only be called once, even if
    'you are selected on this and hit enter
ElseIf KeyAscii = 27 Then 'they hit esc
    Call Form_Unload(0)
End If
End Sub

Private Sub Form_Load()
'load the winsock into memory
Load frmMain    'we are loading the controls on the form, just not making it visible

Dim strName As String
Dim strIP As String

strName = GetSetting(AppName, "main", "name", "Client")
strIP = GetSetting(AppName, "main", "ip", strDefaultIP)

txtName.text = strName
txtIP.text = strIP

Options.Color.Mine = GetSetting(AppName, "main", "mycolor", CLng(vbRed))
Options.Color.Others = GetSetting(AppName, "main", "othercolor", CLng(vbBlue))
Options.Color.Connect = GetSetting(AppName, "main", "connectcolor", CLng(vbGreen))
Options.Color.Disconnect = GetSetting(AppName, "main", "disconnectcolor", CLng(vbRed))
'the server will supply this!
'Options.MaxSend = GetSetting(AppName, "main", "maxsend", 2000)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'close the program
End
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub txtIP_GotFocus()
txtIP.SelStart = 0
txtIP.SelLength = Len(txtIP.text)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
Call cmdConnect_KeyPress(KeyAscii)
End Sub

Private Sub txtIP_KeyPress(KeyAscii As Integer)
Call cmdConnect_KeyPress(KeyAscii)
End Sub

Private Sub txtName_GotFocus()
txtName.SelStart = 0
txtName.SelLength = Len(txtName.text)
End Sub
