Attribute VB_Name = "modDialog"
Public CommonDialog As New cDialog

Public Function DialogHookFunction(ByVal hDlg As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim ComDlg As cDialog
    Set ComDlg = HookedDialog
    If Not (ComDlg Is Nothing) Then
        DialogHookFunction = ComDlg.DialogHook(hDlg, msg, wParam, lParam)
    End If
End Function

Public Sub ClearHookedDialog()
    m_cHookedDialog = 0
End Sub

Public Function NullTrim(s) As String
    Dim i As Integer
    i = InStr(s, vbNullChar)
    If i > 0 Then s = Left$(s, i - 1)
    s = Trim$(s)
    NullTrim = s
End Function
