Attribute VB_Name = "modDraw"
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

Public Sub DrawForm()
    Dim regionButton, regionTree
    regionButton = CreateRoundRectRgn(2, 2, Command1.Width - 2, Command1.Height - 2, 6, 6)
    regionTree = CreateRoundRectRgn(2, 2, tree.Width - 2, tree.Height - 2, 6, 6)
    
    SetWindowRgn frmMain.cmdSend.hwnd, regionButton, True
    SetWindowRgn frmMain.treeUsers.hwnd, regionTree, True
End Sub
