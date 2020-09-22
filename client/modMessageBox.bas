Attribute VB_Name = "modMessageBox"
'new messagebox code so it doesnt stop the flow of the programming
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Const MB_DEFBUTTON1 = &H0&
Const MB_DEFBUTTON2 = &H100&
Const MB_DEFBUTTON3 = &H200&
Const MB_ICONASTERISK = &H40&
Const MB_ICONEXCLAMATION = &H30&
Const MB_ICONHAND = &H10&
Const MB_ICONINFORMATION = MB_ICONASTERISK
Const MB_ICONQUESTION = &H20&
Const MB_ICONSTOP = MB_ICONHAND
Const MB_OK = &H0&
Const MB_OKCANCEL = &H1&
Const MB_YESNO = &H4&
Const MB_YESNOCANCEL = &H3&
Const MB_ABORTRETRYIGNORE = &H2&
Const MB_RETRYCANCEL = &H5&
'end new messagebox

Public Function Msgbox2(Prompt As String, Optional Buttons As VbMsgBoxStyle, Optional Title As String)
Msgbox2 = MessageBox(frmMain.hwnd, Prompt, IIf(IsMissing(Title), App.Title, Title), Buttons)
End Function
