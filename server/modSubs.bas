Attribute VB_Name = "modSubs"
Public Enum SendDataOptions
    sendMain = 1
    sendTemp = 2
    sendMessage = 3
End Enum

Private Type ServerOptions
    MaxSend As Integer
End Type
Public Options As ServerOptions

Public Sub LoadOption(iOption As Byte)
Select Case iOption
    Case Is = 1     'the general options
        Load frmOptions
        frmOptions.lblPacketSize.Caption = Options.MaxSend
        frmOptions.Show
End Select
End Sub

Public Sub AddToLog(strText As String)
'the form adds a line in the log so it is not required
frmMain.txtLog.Text = frmMain.txtLog.Text & vbCrLf & strText
frmMain.txtLog.SelStart = Len(frmMain.txtLog.Text)
End Sub

Public Sub AddToChat(strText As String, Optional splitUser As Boolean)
frmMain.txtChat.SelStart = Len(frmMain.txtChat.Text)

If splitUser = False Then
    frmMain.txtChat.SelText = strText
Else
    frmMain.txtChat.SelText = Left(strText, InStr(1, strText, ":")) & " " 'add a space between
    frmMain.txtChat.SelRTF = Mid(strText, InStr(1, strText, ":") + 1)
End If
End Sub
