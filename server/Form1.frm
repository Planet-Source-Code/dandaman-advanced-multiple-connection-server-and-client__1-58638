VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server"
   ClientHeight    =   2625
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   2415
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   4260
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
   End
   Begin MSWinsockLib.Winsock sockTemp 
      Index           =   0
      Left            =   3120
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sockMain 
      Index           =   0
      Left            =   2640
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sockListen 
      Left            =   2160
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtLog 
      Height          =   2415
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1

Const iMaxConnections As Byte = 50     'we will never get about 255 people connecting at once
Const iMaxTempConnections As Byte = 3
Const iMaxMessage = 100
Const iLocalPort As Integer = 8000      'the port the server will use

Private Type User
    inUse As Boolean
    Verified As Boolean
    Name As String
    Away As Byte    'it is going to be 0 or 1
End Type

Private Users(iMaxConnections) As User
Private Type temp
    inUse As Boolean
    Verified As Boolean
End Type
Private Temps(iMaxTempConnections) As temp   'allow iMaxTempConnections people to attempt to connect at once
Private sendProgress As Boolean

Private Sub Form_Load()
Me.Show
txtLog.Text = "Server loaded " & Now & " on " & sockListen.LocalIP & ":" & sockListen.LocalPort & "."

'get the maxsendsize
Options.MaxSend = GetSetting(Me.Caption, "main", "maxsend", 2000)
AddToLog "Maximum send size: " & Options.MaxSend

'initialize the sockListen
AddToLog "Initializing listen socket."
sockListen.LocalPort = iLocalPort
AddToLog "Socket set to listen mode."
sockListen.Listen

End Sub

Public Sub SendData(strText As String, Optional Index As Byte, Optional temp As SendDataOptions)
If Index = 0 Then   'we are sending to everyone
    Dim iTemp As Byte
    For iTemp = 1 To iMaxConnections
        iTemp = FindSocket(iTemp)   'TODO check whether this should be +1 or not
        If iTemp = 0 Then
            'no sockets left, or all are closed
            'should we alert that it sent to no one? nah just exit the loop
            Exit For
        Else
            'a connected socket has been found
            'now check whether they have verified
            If Users(iTemp).Verified = True Then
                sockMain(iTemp).SendData strText
                DoEvents
            End If
        End If
    Next iTemp
Else
    If temp = 1 Or (temp = 0) Then '0 for default
        sockMain(Index).SendData strText
    ElseIf temp = 2 Then
        sockTemp(Index).SendData strText
    End If
    DoEvents
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Me.Visible = False
SendData "ms2Server shutting down."    'inform everyone
End
End Sub

Private Sub mnuOptions_Click()
Call LoadOption(1)
End Sub

Private Sub sockListen_ConnectionRequest(ByVal requestID As Long)
'we are going to forward the connection reuest to sockmain
'first we loop to find an open socket

Dim iOpenSocket As Byte
iOpenSocket = FindSocket

'check if we are full
If iOpenSocket = 0 Then
    'alert the server
    AddToLog "User connecting but slots full. Making temporary connection."
    'create a temporary connection
    iOpenSocket = FindTempSocket
    If iOpenSocket = 0 Then 'they are 0 available temp sockets
        AddToLog "No available temporary sockets. "
    Else
        Temps(iOpenSocket).inUse = True
        Load sockTemp(iOpenSocket)
        sockTemp(iOpenSocket).Accept requestID
    End If
Else
    'we now have an available socket so lets connect it
    'we use sockmain, instead of socklisten and the index
    Users(iOpenSocket).inUse = True
    'the socket has not yet been loaded
    Load sockMain(iOpenSocket)
    sockMain(iOpenSocket).Accept requestID
End If

'we are still in listen mode, listening for other connections
End Sub

Private Function FindSocket(Optional currentSocket As Byte) As Byte
Dim iTemp As Byte
'currentsocket is so that we can send to all clients
'by default is is optional and set to 0 or 1
If currentSocket = 0 Then
    iTemp = 1
Else
    iTemp = currentSocket
End If

For iTemp = iTemp To iMaxConnections
    'this loop is quick because we are not checking the sockets which
    'are objects and would cause the program to lag
    If currentSocket = 0 Then
        If Users(iTemp).inUse = False Then
            FindSocket = iTemp
            Exit For
        End If
    Else
        If Users(iTemp).inUse = True Then
            FindSocket = iTemp
            Exit For
        End If
    End If
Next iTemp

'if nothing applies above, then it will be 0
End Function

Private Function FindTempSocket() As Byte
Dim iTemp As Byte

For iTemp = 1 To iMaxTempConnections
    If Temps(iTemp).inUse = False Then
        FindTempSocket = iTemp
        Exit For
    End If
Next iTemp

'if nothing applies above, then it will be 0
End Function

Private Sub sockListen_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "Socket listen error: " & Description & ". " & Number
End Sub

Private Sub sockMain_Close(Index As Integer)
AddToLog "User " & Users(Index).Name & "(" & Index & ")" & " has disconnected."
Call DisconnectUsers(CByte(Index))
End Sub

Private Sub sockMain_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strHeader As String
Dim strData As String

sockMain(Index).GetData strData
strHeader = Left(strData, 3)
strData = Mid(strData, 4)

Select Case strHeader
    Case Is = "nam"     'the server version of nam has strdata as long as strheader with it, strdata being the name
        AddToLog "User index " & Index & " has verified name as " & strData & "."
        
        'check to see if anyone else is using this username
        Dim iTemp As Byte
        Do
            'loop through the connected sockets
            iTemp = iTemp + 1
            iTemp = FindSocket(iTemp)
            If iTemp <> 0 Then
                If LCase(Users(iTemp).Name) = LCase(strData) Then 'they have the same username
                    'add to log that someone attempted to connect with used username
                    AddToLog "User " & sockMain(Index).RemoteHostIP & " attempted to connect with used username " & strData & "."
                    'tell the client to disconnect
                    SendData "usd", CByte(Index) 'iTemp
                    'now that we sent this information we will exit this sub
                    'for we do not want it to set the users information
                    Exit Sub
                End If
            Else
                Exit Do
            End If
        Loop
        
        'send confirmation
        'SendData "ver", CInt(Index)
        'us1 for user list
        Dim strTemp As String
        strTemp = SendUsers
        'set the users name
        Users(Index).Name = strData
        'the user is connected
        Users(Index).Verified = True

        SendData "us1" & strTemp, CInt(Index)
        'update all the clients
        
    Case Is = "co1" '    'the last two sends for the connection to take place
        'send all our options
        SendData "opt" & "max" & Options.MaxSend
        
    Case Is = "co2"      'let everyone know he connected ,connected2
        'now we call update users
        Call UpdateUsers(Users(Index).Name)    'update with the username
        
    Case Is = "msg"     'they are sending a message
        AddToLog "Message from " & Index & " (" & Users(Index).Name & ") arrived."
        'add it to the current viewer
        AddToChat Users(Index).Name & ":" & strData, True
        'now send it to all
        SendData "msg" & Users(Index).Name & ":" & strData    'no index because we are sending it to everyone that is connected
    
    Case Is = "awa"
        'someone changed their away state
        Users(Index).Away = Left(strData, 1)
        'addtolog
        AddToLog "(" & Users(Index).Name & ") " & Index & " changed his away state to " & Left(strData, 1) & "."
        'inform everyone else
        SendData "awa" & strData & Users(Index).Name
    
    Case Else
        MsgBox "unknown prefix"
End Select
End Sub

Private Function FindNewMessage() As Integer
Dim iTemp As Integer
For iTemp = 1 To UBound(Message)
    If Message(iTemp).inUse = False Then
        FindNewMessage = iTemp
        Exit For
    End If
Next iTemp
End Function

Private Sub sockMain_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
AddToLog "User " & Users(Index).Name & "(" & Index & ")" & " had winsock error. {" & Description & "}"
'clear them
Call DisconnectUsers(CByte(Index))
End Sub

Private Sub sockTemp_Close(Index As Integer)
    Temps(Index).inUse = False
    Temps(Index).Verified = False
    
    Unload sockTemp(Index)
End Sub

Private Sub sockTemp_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strHeader As String
Dim strData As String

sockTemp(Index).GetData strData
strHeader = Left(strData, 3)
strData = Mid(strData, 4)       'TODO find if this is 3 or four?

Select Case strHeader
    Case Is = "nam"     'the server version of nam has strdata as long as strheader with it, strdata being the name
        AddToLog "User temp index " & Index & " has verified name as " & strData & " but max connections reached."
        'the user is connected
        Temps(Index).Verified = True
        'send information
        'ms2 is for a messagebox
        SendData "ms2" & "The server has reached maximum connections.", CByte(Index), sendTemp
        
    Case Is = "ver"
        'the client has verified the closing of the connection
        'now we close it
        sockTemp(Index).Close
        
    Case Else
        MsgBox "unknown prefix"
End Select
End Sub

Private Sub sockTemp_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "Socket temp error: " & Description & ". " & Number
sockTemp(Index).Close
Temps(Index).inUse = False
Temps(Index).Verified = False
End Sub

Private Sub DisconnectUsers(Index As Byte)
'first we clear the information that the users.things have
Users(Index).Verified = False
Users(Index).inUse = False

'then we reupdate all the client's lists
Call RemoveUsers(Users(Index).Name)

'then we remove the name, because we need it
Users(Index).Name = ""

Unload sockMain(Index)  'unload the winsock
End Sub

Private Function UpdateUsers(newName As String)
    SendData "us2" & "0" & newName '0 for their state
End Function

Private Function RemoveUsers(newName As String)
    SendData "us3" & newName
End Function

Private Function SendUsers()
'make a list of the names
Dim strList As String

Dim iTemp As Byte
Do
    iTemp = iTemp + 1
    iTemp = FindSocket(iTemp)
    If iTemp <> 0 Then
        If Users(iTemp).Name <> "" Then
            strList = strList & IIf(Right(strList, 1) <> "", "," & Users(iTemp).Away & Users(iTemp).Name, Users(iTemp).Away & Users(iTemp).Name)
        End If
    Else
        Exit Do
    End If
Loop

SendUsers = strList
End Function

