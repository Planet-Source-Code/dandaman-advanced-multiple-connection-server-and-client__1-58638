VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMessage 
   Caption         =   "Instant Message"
   ClientHeight    =   3885
   ClientLeft      =   180
   ClientTop       =   750
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar statBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   3630
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6826
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstFontBar 
      Left            =   150
      Top             =   90
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
            Picture         =   "frmMessage.frx":0000
            Key             =   "Emoticon"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessage.frx":0452
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessage.frx":0864
            Key             =   "Normal"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessage.frx":0C76
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessage.frx":1088
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessage.frx":145A
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessage.frx":186C
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessage.frx":1C7E
            Key             =   "Color"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessage.frx":2090
            Key             =   "Face"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessage.frx":251A
            Key             =   "here"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessage.frx":292C
            Key             =   "away"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   660
      Left            =   3300
      TabIndex        =   2
      Top             =   2715
      Width           =   855
   End
   Begin RichTextLib.RichTextBox txtSend 
      Height          =   855
      Left            =   30
      TabIndex        =   0
      Top             =   2730
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1508
      _Version        =   393217
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   2
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMessage.frx":2D3E
   End
   Begin MSComctlLib.Toolbar barFont 
      Height          =   330
      Left            =   30
      TabIndex        =   4
      Top             =   2385
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   582
      ButtonWidth     =   714
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imglstFontBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
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
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   2415
      Left            =   30
      TabIndex        =   1
      Top             =   -30
      Width           =   4095
      _ExtentX        =   7223
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
      TextRTF         =   $"frmMessage.frx":2DC0
   End
   Begin VB.Label lblLength 
      Alignment       =   2  'Center
      Height          =   180
      Left            =   3270
      TabIndex        =   3
      Top             =   3390
      Width           =   870
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iMessageSocket As Byte
Public strMessageTo As String
Private TextLeft As Integer
'Public bRemoteClose As Boolean
'Public tempText As String

Option Base 1
Dim strTemp As String

Private Sub barFont_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case Is = "Emoticon"
        Load frmEmoticon
        With frmEmoticon
            .Top = Me.Top + Me.Height - Me.ScaleHeight + barFont.Top - frmEmoticon.Height - 50
            .Left = Me.Left + Me.Width - Me.ScaleWidth + barFont.Left + Button.Left - 50
            Call .textboxSet(txtSend)
            .Show
            .SetFocus
        End With
    'the font things need 11 spaces in the RTF
    Case Is = "Down"
        If TextLeft > 11 Then
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
        If TextLeft > 11 Then
        Else
            Beep
            Exit Sub
        End If
        
        txtSend.SelFontSize = 8
        If txtSend.SelLength = 0 Then txtSend.Font.Size = 8
    Case Is = "Up"
        If TextLeft > 11 Then
        Else
            Beep
            Exit Sub
        End If
        
        If txtSend.SelFontSize > 21 Then Exit Sub
        If txtSend.SelLength = 0 Then txtSend.Font.Size = txtSend.Font.Size + 2
        txtSend.SelFontSize = txtSend.SelFontSize + 2
    Case Is = "Bold"
        If TextLeft > 7 Then
        Else
            Beep
            Exit Sub
        End If
        
        txtSend.SelBold = CBool(Button.Value)
'        mnuEditBold.Checked = CBool(Button.Value)
    Case Is = "Italic"
        If TextLeft > 7 Then
        Else
            Beep
            Exit Sub
        End If
        
        txtSend.SelItalic = CBool(Button.Value)
'        mnuEditItalic.Checked = CBool(Button.Value)
    Case Is = "Underline"
        If TextLeft > 7 Then
        Else
            Beep
            Exit Sub
        End If
        
        txtSend.SelUnderline = CBool(Button.Value)
'        mnuEditUnderline.Checked = CBool(Button.Value)
    Case Is = "Color"
        If TextLeft > 47 Then
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
            
            If TextLeft > 84 + Len(.FontName) Then
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
End Select
End Sub

Public Sub cmdSend_Click()
If txtSend.text = "" Then Exit Sub
'If Len(txtSend.TextRTF) > 4096 Then
'    'the information is too large to send
'    'this will case two sends, we dont want that
'    Msgbox2 "The data is too large to send.", vbInformation, "Error"
'    Exit Sub
'End If
'first check textleft
If TextLeft < 0 Then
    'something HAS GONE HORRIBLY WRONG!! AROOOGLE!, lol
    Call Msgbox2("Text has reached over maximum limit. Cannot send without first removing some text.", vbInformation, "Error")
    Exit Sub
End If

'If Client.Verified = False Then MsgBox "Server confirmation not yet recieved. Not sending.": Exit Sub
'Client.Verified = False

'send the data
'convert the pictures to text
'Dim strTemp As String
'strTemp = RemovePics(txtSend)

'check if we are connected first...
If Message(iMessageSocket).closesend = True Then
    'they closed their window
    'let us reconnect us together
    'we do not want to yet clear the textbox,
    'for we are going to use its text when we get reconnected
    Status "Reconnecting to " & strMessageTo & "."
    frmMain.SendData "dir" & strMessageTo
    Message(iMessageSocket).closesend = False
    
Else
    Dim strTemp As String
    strTemp = RemovePics(txtSend)
    txtSend.text = ""    'clear the box
    txtSend.SetFocus     'set its focus
    TextLeft = Options.MaxSend - Len(txtSend.TextRTF)   'reset the textleft
    txtSend.Locked = False  'dont even ask if it is locked, just set it to false
    
    Call frmMain.SendData("msg" & Client.Name & ":" & strMessageTo & ":" & strTemp, True, CInt(iMessageSocket)) ' & RemovePics(txtSend), this has already been called, OOPS HAHA i was adding a whole second half! oops
    Status "Message sent."
End If
'we do not want to automatically add what we wrote to the box, because if the server is
'slow then he will think that everyone else got it even though they did not
End Sub

Private Sub Form_Unload(Cancel As Integer)
'i want to unload the form,
'and disconnect the socket
If Message(iMessageSocket).closesend = False Then
    frmMain.SendData "clo" & Client.Name & ":" & strMessageTo, True, CInt(iMessageSocket)
End If
'frmMain.sockMessage(iMessageSocket).Close
'unload and unset everything
'the sockMessage_close will do all the cleanup
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
    TextLeft = Options.MaxSend - Len(txtSend.TextRTF)
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
    TextLeft = Options.MaxSend - iCount
End If

'now check if they like hit shift+enter (which takes 5 spaces) and they only had one left
If TextLeft < 0 Then    'they did something like that
    
End If
End Sub

Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case Shift
    Case Is = 1 'SHIFT
        Select Case KeyCode
            Case Is = 13 'SHIFT+ENTER (new line)
                'this is 6 spaces, check first the length if it can be done
                If TextLeft > 6 Then
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
                If TextLeft > 6 Then
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
        If TextLeft <= 0 And txtSend.SelLength > 0 Then
            txtSend.Locked = False
        ElseIf TextLeft <= 0 Then
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

Private Sub ConvertSmile(strRTF As String)
'search the key names
Dim iTemp As Long
Dim iTemp2 As Long
Dim iTemp3 As Long

Dim iCurrentLocation As Long 'the current location of where the text is
Dim strThing As String
'first we are going to insert the text into the textbox

'AddToChat strRTF

'now that we have done this
'we will replace all the (emoticons) with smilies

iCurrentLocation = 1

'LockWindowUpdate txtChat.hwnd
For iTemp = 1 To frmEmoticon.imglstSmiley.ListImages.Count
    Do
    iTemp2 = InStr(iTemp2 + 1, strRTF, "(" & frmEmoticon.imglstSmiley.ListImages(iTemp).Key & ")")
    'InStr(iTemp2 + 1, strRTF, "(" & frmEmoticon.imglstSmiley.ListImages(iTemp).Key & ")")
    
    If iTemp2 <> 0 Then
        'we found a smiley!
        'what about all the previous text however?
        'lets insert that before we insert the picture
        'txtChat.SelRTF = Mid(strRTF, iCurrentLocation, iTemp2 - iCurrentLocation)
        'strThing = strThing & Mid(strRTF, iCurrentLocation, iTemp2 - iCurrentLocation)
        strRTF = Replace(strRTF, "(" & frmEmoticon.imglstSmiley.ListImages(iTemp).Key & ")", StdPicAsRTF(frmEmoticon.imglstSmiley.ListImages(iTemp).Picture))
        iCurrentLocation = iTemp2 + Len("(" & frmEmoticon.imglstSmiley.ListImages(iTemp).Key & ")")
        'select the text so it gets overwritten
'        txtChat.SelStart = iTemp2 - 1
'        txtChat.SelLength = iCurrentLocation - iTemp2
        'Call InsertPicture(txtChat, frmEmoticon.imglstSmiley.ListImages(iTemp).Picture)
    End If
    Loop Until iTemp2 = 0
Next iTemp
'LockWindowUpdate 0

'now we insert the rest
'txtChat.SelRTF = Mid(strRTF, iCurrentLocation)
txtChat.SelStart = Len(txtChat.text)
txtChat.SelRTF = strRTF
End Sub

Private Function ConvertSmile2(strRTF As String)
'search the key names
Dim iTemp As Long
Dim iTemp2 As Long
Dim iTemp3 As Long

Dim iCurrentLocation As Integer 'the current location of where the text is
Dim strThing As String
'first we are going to insert the text into the textbox

'AddToChat strRTF

'now that we have done this
'we will replace all the (emoticons) with smilies

iCurrentLocation = 1

'LockWindowUpdate txtChat.hwnd
For iTemp = 1 To frmEmoticon.imglstSmiley.ListImages.Count
    Do
    iTemp2 = InStr(iTemp2 + 1, strRTF, "(" & frmEmoticon.imglstSmiley.ListImages(iTemp).Key & ")")
    'InStr(iTemp2 + 1, strRTF, "(" & frmEmoticon.imglstSmiley.ListImages(iTemp).Key & ")")
    
    If iTemp2 <> 0 Then
        'we found a smiley!
        'what about all the previous text however?
        'lets insert that before we insert the picture
        'txtChat.SelRTF = Mid(strRTF, iCurrentLocation, iTemp2 - iCurrentLocation)
        'strThing = strThing & Mid(strRTF, iCurrentLocation, iTemp2 - iCurrentLocation)
        strRTF = Replace(strRTF, "(" & frmEmoticon.imglstSmiley.ListImages(iTemp).Key & ")", StdPicAsRTF(frmEmoticon.imglstSmiley.ListImages(iTemp).Picture))
        iCurrentLocation = iTemp2 + Len("(" & frmEmoticon.imglstSmiley.ListImages(iTemp).Key & ")")
        'select the text so it gets overwritten
        'txtChat.SelStart = iTemp2 - 1
        'txtChat.SelLength = iCurrentLocation - iTemp2
        'Call InsertPicture(txtChat, frmEmoticon.imglstSmiley.ListImages(iTemp).Picture)
    End If
    Loop Until iTemp2 = 0
Next iTemp
'LockWindowUpdate 0

'now we insert the rest
'txtChat.SelRTF = Mid(strRTF, iCurrentLocation)
ConvertSmile2 = strRTF
End Function


Private Sub mnuClose_Click()
'this will close this window
'in this window's close event is the sockmessage(index).close
'and unload sockmessage(index)

Unload Me
End Sub

Public Sub SetIndex(Index As Byte, toward As String)
iMessageSocket = Index
strMessageTo = toward
TextLeft = Options.MaxSend - Len(txtSend.TextRTF)
End Sub

Public Sub AddToChat(strText As String, Optional splitUser As Boolean)
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

Public Sub Status(newText As String)
    statBar.Panels(1).text = newText
End Sub

