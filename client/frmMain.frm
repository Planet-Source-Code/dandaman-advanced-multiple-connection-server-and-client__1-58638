VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Chat"
   ClientHeight    =   3840
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sockMessage 
      Index           =   0
      Left            =   1680
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imglstAway 
      Left            =   600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "here"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0352
            Key             =   "away"
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock sockMain 
      Left            =   1200
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imglstFontBar 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   16
      MaskColor       =   14215660
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":06A4
            Key             =   "Emoticon"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AF6
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F08
            Key             =   "Normal"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":131A
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":172C
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AFE
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F10
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2322
            Key             =   "Color"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2734
            Key             =   "Face"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2BBE
            Key             =   "here"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2FD0
            Key             =   "away"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView treeUsers 
      Height          =   3615
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   6376
      _Version        =   393217
      Sorted          =   -1  'True
      Style           =   1
      FullRowSelect   =   -1  'True
      ImageList       =   "imglstAway"
      Appearance      =   1
   End
   Begin RichTextLib.RichTextBox txtSend 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1508
      _Version        =   393217
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   2
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":33E2
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4260
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":3464
   End
   Begin MSComctlLib.Toolbar barFont 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   2535
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   582
      ButtonWidth     =   714
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imglstFontBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Emoticon"
            Object.ToolTipText     =   "Insert Emoticon"
            ImageKey        =   "Emoticon"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Down"
            Object.ToolTipText     =   "Smaller Font Size"
            ImageKey        =   "Down"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Normal"
            Object.ToolTipText     =   "Normal Size Font"
            ImageKey        =   "Normal"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Up"
            Object.ToolTipText     =   "Bigger Font Size"
            ImageKey        =   "Up"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold Font"
            ImageKey        =   "Bold"
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic Font"
            ImageKey        =   "Italic"
            Style           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline Font"
            ImageKey        =   "Underline"
            Style           =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Color"
            Object.ToolTipText     =   "Font Color"
            ImageKey        =   "Color"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Face"
            Object.ToolTipText     =   "Font Face"
            ImageKey        =   "Face"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "away"
            ImageKey        =   "here"
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   660
      Left            =   3855
      TabIndex        =   2
      Top             =   2865
      Width           =   855
   End
   Begin VB.Label lblLength 
      Alignment       =   2  'Center
      Height          =   180
      Left            =   3840
      TabIndex        =   5
      Top             =   3540
      Width           =   870
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuBar1 
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
Option Explicit
'Dim strTemp As String

Private Sub barFont_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case Is = "Emoticon"
        Load frmEmoticon
        With frmEmoticon
            .Top = frmMain.Top + Me.Height - Me.ScaleHeight + barFont.Top - frmEmoticon.Height - 50
            .Left = frmMain.Left + Me.Width - Me.ScaleWidth + barFont.Left + Button.Left - 50
             Call .textboxSet(txtSend)
            .Show
            .SetFocus
        End With
    'the font things need 11 spaces in the RTF
    Case Is = "Down"
        If Options.TextLeft > 11 Then
        Else
            Beep
            Exit Sub
        End If
        
        If txtSend.SelFontSize < 3 Then Exit Sub
        If txtSend.SelLength = 0 Then txtSend.Font.Size = txtSend.Font.Size - 2
        If txtSend.SelFontSize = Null Then
            'set it to default (8) - 2
            txtSend.SelFontSize = 6
        Else
            txtSend.SelFontSize = txtSend.SelFontSize - 2
        End If
    Case Is = "Normal"
        If Options.TextLeft > 11 Then
        Else
            Beep
            Exit Sub
        End If
        
        txtSend.SelFontSize = 8
        If txtSend.SelLength = 0 Then txtSend.Font.Size = 8
    Case Is = "Up"
        If Options.TextLeft > 11 Then
        Else
            Beep
            Exit Sub
        End If
        
        If txtSend.SelFontSize > 21 Then Exit Sub
        If txtSend.SelLength = 0 Then txtSend.Font.Size = txtSend.Font.Size + 2
        txtSend.SelFontSize = txtSend.SelFontSize + 2
    Case Is = "Bold"
        If Options.TextLeft > 7 Then
        Else
            Beep
            Exit Sub
        End If
        
        txtSend.SelBold = CBool(Button.Value)
'        mnuEditBold.Checked = CBool(Button.Value)
    Case Is = "Italic"
        If Options.TextLeft > 7 Then
        Else
            Beep
            Exit Sub
        End If
        
        txtSend.SelItalic = CBool(Button.Value)
'        mnuEditItalic.Checked = CBool(Button.Value)
    Case Is = "Underline"
        If Options.TextLeft > 7 Then
        Else
            Beep
            Exit Sub
        End If
        
        txtSend.SelUnderline = CBool(Button.Value)
'        mnuEditUnderline.Checked = CBool(Button.Value)
    Case Is = "Color"
        If Options.TextLeft > 47 Then
        Else
            Beep
            Exit Sub
        End If
        
        CommonDialog.ShowColor
        If CommonDialog.Color <> -1 Then txtSend.SelColor = CommonDialog.Color
        
    Case Is = "Face"
        With CommonDialog
            .FontFlags = CF_LIMITSIZE 'CF_WYSIWYG + CF_EFFECTS + CF_LIMITSIZE
            .FontMinSize = 3
            .FontMaxSize = 21
            .ShowFont
            
            If Options.TextLeft > 84 + Len(.FontName) Then
            Else
                Beep
                Exit Sub
            End If

            
            If .FontName <> "" Then txtSend.SelFontName = .FontName
            If .FontColor <> -1 Then txtSend.SelColor = .FontColor
            If .FontBold <> Empty Then txtSend.SelBold = .FontBold
            If .FontItalic <> Empty Then txtSend.SelItalic = .FontItalic
            If .FontSize <> Empty Then txtSend.SelFontSize = .FontSize
            
        End With
        
    Case Is = "away"
        'they changed the button state
        'this will run when we set the chkaway value in form_load
        'so we will add a sub to check if we have connected
        'If sockMain.State = 7 Then
        
        'find the index
        Dim iIndex As Byte
        For iIndex = 1 To treeUsers.Nodes.Count
            If treeUsers.Nodes(iIndex).text = Client.Name Then
                Exit For
            End If
        Next iIndex
        If treeUsers.Nodes(iIndex).Image <> Button.Image Then Debug.Print "SAME": Button.Value = IIf(Button.Value = 1, 0, 1): Exit Sub
        'check if the server has previously responded yet
        Button.Image = IIf(Button.Value = 1, "away", "here")
        cmdSend.Enabled = Not CBool(Button.Value)
        SendData "awa" & Button.Value
        'End If
End Select
End Sub

Private Sub cmdSend_Click()
If txtSend.text = "" Then Exit Sub
'If Len(txtSend.TextRTF) > 4096 Then
'    'the information is too large to send
'    'this will case two sends, we dont want that
'    Msgbox2 "The data is too large to send.", vbInformation, "Error"
'    Exit Sub
'End If
'first check options.textleft
If Options.TextLeft < 0 Then
    'something HAS GONE HORRIBLY WRONG!! AROOOGLE!, lol
    Call Msgbox2("Text has reached over maximum limit. Cannot send without first removing some text.", vbInformation, "Error")
    Exit Sub
End If

If Client.Verified = False Then MsgBox "Server confirmation not yet recieved. Not sending.": Exit Sub
Client.Verified = False

'send the data
'convert the pictures to text
'Dim strTemp As String
'strTemp = RemovePics(txtSend)
Dim strTemp As String
strTemp = RemovePics(txtSend)
'clear the box
txtSend.text = ""
 'txtSend.TextRTF = ""
'txtSend.Text = ""   'so it resets its textrtf, and removes past sizes and all
'set its focus
txtSend.SetFocus
'reset the textleft
Options.TextLeft = Options.MaxSend - Len(txtSend.TextRTF)
txtSend.Locked = False  'dont even ask if it is locked, just set it to false
SendData "msg" & strTemp  ' & RemovePics(txtSend), this has already been called, OOPS HAHA i was adding a whole second half! oops
'we do not want to automatically add what we wrote to the box, because if the server is
'slow then he will think that everyone else got it even though they did not
End Sub

Private Sub Form_Unload(Cancel As Integer)
'close the program
End
End Sub

Private Sub mnuOptions_Click()
'Load frmOptions
'frmOptions.Show
Call LoadOption(1)
End Sub

Private Sub mnuQuit_Click()
'close the program
Call Form_Unload(0)
End Sub

Private Sub sockMain_Close()
'since it closed, most likely the server died
frmMain.Visible = False
frmConnect.Visible = True

'we need to clear all the controls on frmMain
txtSend.text = ""
txtChat.text = ""
treeUsers.Nodes.Clear

'clear all the information
Server.Connected = False
Client.Verified = True
End Sub

Private Sub sockMain_Connect()
'we are connected
'so we hide the frmConnect and show the frmMain

Client.Verified = True
SendData "nam" & Client.Name  'send our name"
End Sub

Private Sub sockMain_DataArrival(ByVal bytesTotal As Long)
Dim strHeader As String
Dim strData As String

sockMain.GetData strData
strHeader = Left(strData, 3)
strData = Mid(strData, 4)

Select Case strHeader

    Case Is = "ver"
        'the server is sending a verification
    
    Case Is = "nam"
        'the server wants to change the name
        Client.Name = strData
        AddToChat "Server changed your name."
    
    Case Is = "usd" 'used
        'disconnect so the server can use the slot
        sockMain.Close
        'now show a messagebox that will tell the user that the server does not allow miltiple usernames
        'the server said that the client name is used, please change
        Msgbox2 "Nickname in use. Change nickname."
        
    Case Is = "msg"
        'the server sent a message to us, add it
        AddToChat strData, True
        'Beep   'only beep on errors
        
    Case Is = "ms2"
        'message2
        'a messagebox will show
        Msgbox2 strData
        
    Case Is = "us1" 'user lists
        'strdata will have all the users names
        Dim strTemp2() As String
        strTemp2 = Split(strData, ",")

        'If UBound(strTemp) = -1 Then 'they are the first to connect
            'do nothing!
   'elseif
'        If UBound(strTemp2) = 0 Then 'they sent only one person
            'add the user alone
'            treeUsers.Nodes.Add , , , Mid(strData, 2), IIf(Mid(strData, 4, 1) = 0, "here", "away") 'Mid(strData, 4, 1) + 1
        'ElseIf UBound(strTemp2) = -1 Then    'strdata was blank
        If UBound(strTemp2) = -1 Then    'strdata was blank
            'do nothing
            
        Else
            Dim iTemp As Byte
            For iTemp = 0 To UBound(strTemp2)    '-1 for the 0 to 3, that makes 4 not 3
                'add the users to the tree
                treeUsers.Nodes.Add , , , Mid(strTemp2(iTemp), 2), IIf(Left(strTemp2(iTemp), 1) = 0, "here", "away") 'CInt(Left(strTemp2(iTemp), 1)) 'otherwise its looking for key "1"
            Next iTemp
        End If
        
        'Server.Connected = True
        'now make the form visible
        'frmConnect.Visible = False
        'frmMain.Visible = True
        'now that we are actually connected
        'we send a verification
        SendData "co1"
    
    Case Is = "us2"
        'add a user
        treeUsers.Nodes.Add , , , Mid(strData, 2), "here" 'this will always be here 'IIf(Left(strData, 1) = 0, "here", "away") '+1 for the index
        'inform the chat that he has connected
        txtChat.SelStart = Len(txtChat.text)
        txtChat.SelColor = Options.Color.Connect
        txtChat.SelText = Mid(strData, 2) & " has connected." & vbCrLf
        txtChat.SelColor = vbBlack
        
    Case Is = "us3"
        'remove a user
        Dim iIndex As Byte
        'find the index
        For iIndex = 1 To treeUsers.Nodes.Count
            If treeUsers.Nodes(iIndex).text = strData Then
                treeUsers.Nodes.Remove iIndex
                
                txtChat.SelStart = Len(txtChat.text)
                txtChat.SelColor = Options.Color.Disconnect
                txtChat.SelText = strData & " has disconnected." & vbCrLf
                txtChat.SelColor = vbBlack
                
                Exit For
            End If
        Next iIndex
        
    Case Is = "awa"     'away settings for each user
        'the data will look like the following
        '0Therapy for therapy being not away
        '1Doug for doug being away
        
        'first we find who theyre looking for
        Dim iIndex2 As Byte
        'find the index
        For iIndex2 = 1 To treeUsers.Nodes.Count
            If treeUsers.Nodes(iIndex2).text = Mid(strData, 2) Then
                treeUsers.Nodes(iIndex2).Image = IIf(Left(strData, 1) = 0, "here", "away") 'Left(strData, 1) + 1
                Exit For
            End If
        Next iIndex2
    
    Case Is = "opt"     'they want to change an option
        Dim strTemp() As String
        strTemp = Split(strData, ",")
        
        '0 is the first one
        Dim iCount As Integer
        For iCount = 0 To UBound(strTemp)
            Select Case Left(strTemp(iCount), 3)
                Case Is = "max"     'the maximum length
                    Options.MaxSend = Mid(strTemp(iCount), 4)
                    Options.TextLeft = Options.MaxSend - IIf(txtSend.OLEObjects.Count = 0, Len(txtSend.TextRTF), Len(RemovePics(txtSend)))
                    lblLength.Caption = Options.TextLeft - Len(txtSend.TextRTF)
                    
                'Case Is = "other option"
                    'do something
            End Select
        Next iCount
        
        'load the form, this only happ
        'If Me.Visible = False Then
        If Server.Connected = False Then
            Server.Connected = True
            Me.Visible = True
            frmConnect.Visible = False
            SendData "co2"
        End If
        
    Case Else
        MsgBox strHeader & " " & strData
End Select
End Sub

Public Sub SendData(strText As String, Optional Message As Boolean, Optional Index As Integer)
If Message = False Then
    If sockMain.State = 7 Then
        'we are connected so send the data
        sockMain.SendData strText
        DoEvents
    Else
        'we are not connected
        Msgbox2 "attempt to send without connection"
    End If
Else
    If sockMessage(Index).State = 7 Then
        'we are connected so send the data
        sockMessage(Index).SendData strText
        DoEvents
    Else
        'we are not connected
        Msgbox2 "attempt to send without connection"
    End If
End If
End Sub

Private Sub sockMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'msgbox2 "Winsock err: " & Description
End Sub

Private Sub AddToChat(strText As String, Optional splitUser As Boolean)
txtChat.SelStart = Len(txtChat.text)

If splitUser = False Then
    txtChat.SelText = strText
Else
    'split the user from the text
    Dim strTemp As String
    strTemp = Left(strText, InStr(1, strText, ":")) & " " 'add a space between
    If Left(strTemp, Len(strTemp) - 2) = Client.Name Then
        'verify so we can send more
        txtChat.SelColor = Options.Color.Mine
        Client.Verified = True
    Else
        txtChat.SelColor = Options.Color.Others
    End If
    
    txtChat.SelText = strTemp
    txtChat.SelColor = vbBlack
    'change the smiley codes
    'txtChat.SelRTF = ConvertSmile(Mid(strText, InStr(1, strText, ":") + 1), txtChat)
    'Call ConvertSmile(Mid(strText, InStr(1, strText, ":") + 1))
    txtChat.SelRTF = ConvertSmile2(Mid(strText, InStr(1, strText, ":") + 1))
End If
End Sub

Private Sub txtChat_KeyDown(KeyCode As Integer, Shift As Integer)
'if they are selected on this and start typing
'then automatically select the txtsend for them
If KeyCode = 9 Then Exit Sub     'TAB
If KeyCode = 8 Then txtSend.SetFocus: Exit Sub  'backspace
If KeyCode = 93 Then Exit Sub   'RightClick menu
If KeyCode = 46 Then txtSend.SetFocus: Exit Sub 'delete

If KeyCode = 37 Or KeyCode = 38 Or KeyCode = 39 Or KeyCode = 40 Or KeyCode = 16 Or KeyCode = 17 Then 'up, down, left, right
    'do nothing
Else
    If Shift = 0 Then
        txtSend.SelText = LCase(Chr(KeyCode))
    ElseIf Shift = 1 Then
        txtSend.SelText = Chr(KeyCode)
    End If
    txtSend.SetFocus
End If
End Sub

Private Sub txtSend_Change()
Restart:

'get the length of the text, including all of the images
If txtSend.OLEObjects.Count = 0 Then
    'there are no pictures, so set the size to the maxlength -len(txtsend.text)
    lblLength.Caption = Options.MaxSend - Len(txtSend.TextRTF)
    Options.TextLeft = Options.MaxSend - Len(txtSend.TextRTF)
Else
    'make a loop to get the keys of all the pictures
    Dim strTemp As String
    Dim iCount As Integer
    
    'For iCount = 0 To txtSend.OLEObjects.Count - 1
    '    strTemp = strTemp & "(" & txtSend.OLEObjects(iCount).Key & ")"
    'Next iCount
    
    'now that we have the length of all the textrtf
    'it must first set the key!
    iCount = Len(RemovePics(txtSend))
    lblLength.Caption = Options.MaxSend - iCount
    Options.TextLeft = Options.MaxSend - iCount
End If

'now check if they like hit shift+enter (which takes 5 spaces) and they only had one left
If Options.TextLeft < 0 Then    'they did something like that
    
End If
End Sub

Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case Shift
    Case Is = 1 'SHIFT
        Select Case KeyCode
            Case Is = 13 'SHIFT+ENTER (new line)
                'this is 6 spaces, check first the length if it can be done
                If Options.TextLeft > 6 Then
                    'do nothing
                Else
                    'set the keycode to 0 and beep
                    KeyCode = 0
                    Shift = 0
                    Beep
                End If
        End Select
    
    Case Is = 2 'CTRL
        Select Case KeyCode
            Case Is = 13 'CTRL+ENTER (new line)
                If Options.TextLeft > 6 Then
                    'do nothing
                Else
                    KeyCode = 0
                    Shift = 0
                    Beep
                End If
                
            Case Is = 66 'b - Bold (CTRL+B)
                barFont.Buttons("Bold").Value = IIf(barFont.Buttons("Bold").Value = tbrPressed, tbrUnpressed, tbrPressed)
                Call barFont_ButtonClick(barFont.Buttons("Bold"))
                KeyCode = 0
                Shift = 0
            
            Case Is = 73 'i - Italic (CTRL+I)
                barFont.Buttons("Italic").Value = IIf(barFont.Buttons("Italic").Value = tbrPressed, tbrUnpressed, tbrPressed)
                Call barFont_ButtonClick(barFont.Buttons("Italic"))
                KeyCode = 0
                Shift = 0
                
            Case Is = 85 'u - Underline (CTRL+U)
                barFont.Buttons("Underline").Value = IIf(barFont.Buttons("Underline").Value = tbrPressed, tbrUnpressed, tbrPressed)
                Call barFont_ButtonClick(barFont.Buttons("Underline"))
                KeyCode = 0
                Shift = 0
            Case Is = 86 'v - Paste (CTRL+V)
                'we are going to allow to paste pictures
                'check!
                'txtSend.SelText = Clipboard.GetData(vbCFDIB)
                
                'txtSend.SelText = Clipboard.GetData(1)
                If Clipboard.GetFormat(1) = True Then
                    txtSend.SelText = Clipboard.GetText
                ElseIf Clipboard.GetFormat(2) = True Then
                    txtSend.SelText = Clipboard.GetText
                ElseIf Clipboard.GetFormat(3) = True Then
                    'we do not want to set it, for they key cannot be set
                    'Call InsertPicture(txtSend, Clipboard.GetData)
                End If
                KeyCode = 0
                Shift = 0
            Case Is = 187   'CTRL +
                'make size bigger
                Call barFont_ButtonClick(barFont.Buttons("Up"))
                KeyCode = 0
                Shift = 0
            Case 189        'CTRL -
                'make size smaller
                Call barFont_ButtonClick(barFont.Buttons("Down"))
                KeyCode = 0
                Shift = 0
        End Select
End Select

Select Case KeyCode
    Case Is = 8  'BACKSPACE
        'this is to unlock it
        txtSend.Locked = False
    Case Is = 13 'ENTER
        If Shift = 0 Then
            If cmdSend.Enabled = True Then
                Call cmdSend_Click
            End If
            KeyCode = 0 'set it to 0 cuz if they hit enter
                        'so they see they cant send
        End If
    Case Else
        'they hit any other key, for example while regularly typing
        
        'check the length
        If Options.TextLeft <= 0 And txtSend.SelLength > 0 Then
            txtSend.Locked = False
        ElseIf Options.TextLeft <= 0 Then
             txtSend.Locked = True
        End If
End Select

End Sub

Private Sub txtSend_SelChange()
'the selection that we have has changed, so we need to change
'the buttons in the barFont accordingly

Select Case txtSend.SelBold
    Case True
        'barFont.Buttons("Bold").MixedState = False
        barFont.Buttons("Bold").Value = tbrPressed
    Case False
        'barFont.Buttons("Bold").MixedState = False
        barFont.Buttons("Bold").Value = tbrUnpressed
    Case Else
        'barFont.Buttons("Bold").MixedState = True
        barFont.Buttons("Bold").Value = tbrPressed
End Select

Select Case txtSend.SelItalic
    Case True
        'barFont.Buttons("Italic").MixedState = False
        barFont.Buttons("Italic").Value = tbrPressed
    Case False
        'barFont.Buttons("Italic").MixedState = False
        barFont.Buttons("Italic").Value = tbrUnpressed
    Case Else
        'barFont.Buttons("Italic").MixedState = True
        barFont.Buttons("Italic").Value = tbrPressed
End Select

Select Case txtSend.SelUnderline
    Case True
        'barFont.Buttons("Underline").MixedState = False
        barFont.Buttons("Underline").Value = tbrPressed
    Case False
        'barFont.Buttons("Underline").MixedState = False
        barFont.Buttons("Underline").Value = tbrUnpressed
    Case Else
        'barFont.Buttons("Underline").MixedState = True
        barFont.Buttons("Underline").Value = tbrPressed
End Select

End Sub

