Attribute VB_Name = "modRedo"
'// View Types
Public Enum ERECViewModes
    ercDefault = 0
    ercWordWrap = 1
    ercWYSIWYG = 2
End Enum
'// Undo Types
Public Enum ERECUndoTypeConstants
    ercUID_UNKNOWN = 0
    ercUID_TYPING = 1
    ercUID_DELETE = 2
    ercUID_DRAGDROP = 3
    ercUID_CUT = 4
    ercUID_PASTE = 5
End Enum
'// Text Modes
Public Enum TextMode
    TM_PLAINTEXT = 1
    TM_RICHTEXT = 2 ' /* default behavior */
    TM_SINGLELEVELUNDO = 4
    tm_multilevelundo = 8 ' /* default behavior */
    TM_SINGLECODEPAGE = 16
    TM_MULTICODEPAGE = 32 ' /* default behavior */
End Enum

Public Const WM_COPY = &H301
Public Const WM_CUT = &H300
Public Const WM_PASTE = &H302

Public Const WM_USER = &H400
Public Const EM_SETTEXTMODE = (WM_USER + 89)
Public Const EM_UNDO = &HC7
Public Const EM_REDO = (WM_USER + 84)
Public Const EM_CANPASTE = (WM_USER + 50)
Public Const EM_CANUNDO = &HC6&
Public Const EM_CANREDO = (WM_USER + 85)
Public Const EM_GETUNDONAME = (WM_USER + 86)
Public Const EM_GETREDONAME = (WM_USER + 87)

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long


Public Property Get UndoType() As ERECUndoTypeConstants
    UndoType = SendMessageLong(frmMain.txtSend.hwnd, EM_GETUNDONAME, 0, 0)
End Property
Public Property Get RedoType() As ERECUndoTypeConstants
    RedoType = SendMessageLong(frmMain.txtSend.hwnd, EM_GETREDONAME, 0, 0)
End Property
Public Property Get CanPaste() As Boolean
   CanPaste = SendMessageLong(frmMain.txtSend.hwnd, EM_CANPASTE, 0, 0)
End Property
Public Property Get CanCopy() As Boolean
   If frmMain.txtSend.hwnd < 0 Then
      CanCopy = True
   End If
End Property
Public Property Get CanUndo() As Boolean
    CanUndo = SendMessageLong(frmMain.txtSend.hwnd, EM_CANUNDO, 0, 0)
End Property
Public Property Get CanRedo() As Boolean
    CanRedo = SendMessageLong(frmMain.txtSend.hwnd, EM_CANREDO, 0, 0)
End Property

 Public Sub Undo()
    SendMessageLong frmMain.txtSend.hwnd, EM_UNDO, 0, 0
End Sub

Public Sub Redo()
    SendMessageLong frmMain.txtSend.hwnd, EM_REDO, 0, 0
End Sub
Public Sub Cut()
   SendMessageLong frmMain.txtSend.hwnd, WM_CUT, 0, 0
End Sub
Public Sub Copy()
   SendMessageLong frmMain.txtSend.hwnd, WM_COPY, 0, 0
End Sub
Public Sub Paste()
   SendMessageLong frmMain.txtSend.hwnd, WM_PASTE, 0, 0
End Sub
Public Sub Clear()
   frmMain.txtSend.SelText = Empty
End Sub
Public Sub UpdateItems()
    Dim bCanUndo As Boolean
    '// Undo/Redo options:
    bCanUndo = CanUndo
    mnuEditUndo.Enabled = bCanUndo
    '// Set Undo Text
    If (bCanUndo) Then
        mnuEditUndo.Caption = "&Undo " & TranslateUndoType(UndoType)
    Else
        mnuEditUndo.Caption = "&Undo"
    End If
    '// Set Redo Text
    bCanUndo = CanRedo
    If (bCanUndo) Then
        mnuEditRedo.Caption = "&Redo " & TranslateUndoType(RedoType)
    Else
        mnuEditRedo.Caption = "&Redo"
    End If
    mnuEditRedo.Enabled = bCanUndo
    tbToolBar.Buttons("Redo").Enabled = bCanUndo
    '// Cut/Copy/Paste/Clear options
    mnuEditCut.Enabled = CanCopy
    mnuEditCopy.Enabled = CanCopy
    mnuEditPaste.Enabled = CanPaste
    mnuEditClear.Enabled = CanCopy
End Sub
'// Returns the undo/redo type
Public Function TranslateUndoType(ByVal eType As ERECUndoTypeConstants) As String
   Select Case eType
   Case ercUID_UNKNOWN
      TranslateUndoType = "Last Action"
   Case ercUID_TYPING
      TranslateUndoType = "Typing"
   Case ercUID_PASTE
      TranslateUndoType = "Paste"
   Case ercUID_DRAGDROP
      TranslateUndoType = "Drag Drop"
   Case ercUID_DELETE
      TranslateUndoType = "Delete"
   Case ercUID_CUT
      TranslateUndoType = "Cut"
   End Select
End Function
