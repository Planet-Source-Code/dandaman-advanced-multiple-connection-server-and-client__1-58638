Attribute VB_Name = "modPicture"
Public Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_PASTE = &H302
Public iCountObject As Integer

'Public Sub InsertPicture(RTB As RichTextBox, Picture As StdPicture, Optional strKey As String)  'edited my me
'Dim bTemp As Boolean
'Clipboard.Clear
'Clipboard.SetData Picture
'If strKey <> "" Then
'    iCountObject = iCountObject + 1
'End If
''we must unlock it first
'If RTB.Locked = True Then RTB.Locked = False: bTemp = True
'SendMessage RTB.hwnd, WM_PASTE, 0, 0 'Can't open clipboard error 'heres a problem, vb starts the RTB's Change event before setting the key
'If bTemp = True Then RTB.Locked = True
'If strKey <> "" Then
'    RTB.OLEObjects(RTB.OLEObjects.Count - 1).Key = strKey & RTB.OLEObjects.Count 'to keep unique
'End If
'End Sub

Public Sub InsertPicture(RTB As RichTextBox, Picture As StdPicture, Optional strKey As String)
Dim strTemp As String
Dim strKeyTemp As String
strKeyTemp = strKey     'for some reason, after going through StdPicAsRTF, strKey gets erased
strTemp = IIf(strKey <> "", strKey & RTB.OLEObjects.Count, "")
'RTB.OLEObjects.Add , , Picture
'SavePicture Picture, "c:/test.jpg"
'RTB.OLEObjects.Add , , "c:/test.jpg", 65
RTB.SelRTF = StdPicAsRTF(Picture)   'here it inserts the picture, and calls the selchange, and change of txtsend before it even sets the KEY!
strKey = strKeyTemp
If strKey = "" Then strKey = "external"  'they pasted the picture in
RTB.OLEObjects(RTB.OLEObjects.Count - 1).Key = strKey & RTB.OLEObjects.Count 'to keep unique
'how to get the selected one?


'its the rtb.selstart minus the pict

'Dim iTemp1 As Long, iTemp2 As Long
'itemp1 = instrrev(rtb.TextRTF,"{\pict"

'RTB.OLEObjects(RTB.OLEObjects.Count - 1).Key = "(smile)"
End Sub

Public Function RemovePics(RTB As RichTextBox)
Dim strRTF As String, strKey As String
strRTF = RTB.TextRTF

Dim iTemp As Long, iTemp2 As Long, iCountDown As Integer, iCountUp As Integer
iCountDown = RTB.OLEObjects.Count
iCountUp = 0

Dim iSize As Long

Do Until iCountDown = 0
    iTemp = InStr(1, strRTF, "{\pict", vbTextCompare) - 1
    iTemp2 = InStr(iTemp + 1, strRTF, "}", vbTextCompare)
    strKey = RTB.OLEObjects(iCountUp).Key
    'if strKey is blank, then that means they pasted a picture into this

    If strKey = "" Then 'this picture is the one that has just been inserted, and has not yet had a key set 'they pasted the picture in
        strKey = Options.Temp
        'set this permanently, so later we will not get external
        RTB.OLEObjects(iCountUp).Key = Options.Temp
        Options.Temp = ""
        'strKey = "external"   'temp was holding the key from the insertpicture on frmEmoticon
        'strRTF = strRTF 'Left(strRTF, iTemp) & Right(strRTF, Len(strRTF) - iTemp2)
        strRTF = Left(strRTF, iTemp) & "(" & strKey & ")" & Right(strRTF, Len(strRTF) - iTemp2)
        'iSize = iSize + (iTemp2 - iTemp)    'this is just for the foreign picture sends
    Else
        strKey = Left(strKey, InStr(1, strKey, CStr(iCountUp + 1)) - 1)
        strRTF = Left(strRTF, iTemp) & "(" & strKey & ")" & Right(strRTF, Len(strRTF) - iTemp2)
    End If
    
    'strRTF = Left(strRTF, iTemp) & "(" & strKey & ")" & Right(strRTF, Len(strRTF) - iTemp2)
    iCountDown = iCountDown - 1
    iCountUp = iCountUp + 1
Loop

RemovePics = strRTF
End Function



'How to find where the SelStart is for TextRTF?
Public Function FindRTFStart(RTF As RichTextBox) As Long

End Function
