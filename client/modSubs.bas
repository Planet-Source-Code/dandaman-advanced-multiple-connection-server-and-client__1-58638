Attribute VB_Name = "modSubs"
Option Explicit

Private Type ServerInfo
    Connected As Boolean
End Type

Private Type ClientInfo
    Name As String
    Verified As Boolean
End Type

Public Server As ServerInfo
Public Client As ClientInfo

'''''''''''''''''''''''''''''''''''''''
Private Type Colors
    Mine As Long    'my name's color
    Others As Long  'other's name's colors
    Connect As Long
    Disconnect As Long
End Type
Private Type ClientOptions
    'this will addtochat all sorts of non-important things
    'like when someone's set their away to ON (1)
    Log As Boolean
    'the colors of the user's name
    Color As Colors
    'the maximum size of a send
    MaxSend As Integer
    'the amount left to type
    TextLeft As Integer
    'a string to hold temporary information
    Temp As String
End Type
Public Options As ClientOptions
Public Const AppName As String = "SchoolChat"

Public Sub LoadOption(iData As Byte)
Load frmOptions

Select Case iData
    Case 1      'Customize
        'set all properties
        With frmOptions
            .cmdColor(0).BackColor = Options.Color.Mine
            .cmdColor(1).BackColor = Options.Color.Others
            .cmdColor(2).BackColor = Options.Color.Connect
            .cmdColor(3).BackColor = Options.Color.Disconnect
            
            .picCustomize.Left = 0
            .picCustomize.Top = 0
            .picCustomize.BorderStyle = 0
            .Left = 0
            .picButton.Top = .picCustomize.Height
            .picButton.BorderStyle = 0
            .Width = .picCustomize.Width
            .Height = .picCustomize.Height + .picButton.Height + (frmOptions.Height - frmOptions.ScaleHeight) '.picButton.Height + (.picButton.Top)
            .picCustomize.Visible = True
        End With
        
    Case 2
        'stuff
        
End Select

frmOptions.Show
End Sub


'Private Sub ConvertSmile(strRTF As String)
''search the key names
'Dim iTemp As Long
'Dim iTemp2 As Long
'Dim iTemp3 As Long
'
'Dim iCurrentLocation As Long 'the current location of where the text is
'Dim strThing As String
''first we are going to insert the text into the textbox
'
''AddToChat strRTF'

''now that we have done this
''we will replace all the (emoticons) with smilies
'
'iCurrentLocation = 1 '
'
''LockWindowUpdate txtChat.hwnd
'For iTemp = 1 To frmEmoticon.imglstSmiley.ListImages.Count
'    Do
'    iTemp2 = InStr(iTemp2 + 1, strRTF, "(" & frmEmoticon.imglstSmiley.ListImages(iTemp).Key & ")")
'    'InStr(iTemp2 + 1, strRTF, "(" & frmEmoticon.imglstSmiley.ListImages(iTemp).Key & ")")
'
'    If iTemp2 <> 0 Then
'        'we found a smiley!
'        'what about all the previous text however?
'        'lets insert that before we insert the picture
'        'txtChat.SelRTF = Mid(strRTF, iCurrentLocation, iTemp2 - iCurrentLocation)
'        'strThing = strThing & Mid(strRTF, iCurrentLocation, iTemp2 - iCurrentLocation)
'        strRTF = Replace(strRTF, "(" & frmEmoticon.imglstSmiley.ListImages(iTemp).Key & ")", StdPicAsRTF(frmEmoticon.imglstSmiley.ListImages(iTemp).Picture))
'        iCurrentLocation = iTemp2 + Len("(" & frmEmoticon.imglstSmiley.ListImages(iTemp).Key & ")")
'        'select the text so it gets overwritten
''        txtChat.SelStart = iTemp2 - 1
''        txtChat.SelLength = iCurrentLocation - iTemp2
'        'Call InsertPicture(txtChat, frmEmoticon.imglstSmiley.ListImages(iTemp).Picture)
'    End If
'    Loop Until iTemp2 = 0
'Next iTemp
''LockWindowUpdate 0'
'
''now we insert the rest
''txtChat.SelRTF = Mid(strRTF, iCurrentLocation)
'txtChat.SelStart = Len(txtChat.Text)
'txtChat.SelRTF = strRTF
'End Sub

Public Function ConvertSmile2(strRTF As String)
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

