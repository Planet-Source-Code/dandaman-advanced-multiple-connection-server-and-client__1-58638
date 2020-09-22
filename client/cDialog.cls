VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================================
'  Common Dialog Class
'  Accesses COMDLG32.DLL directly without Microsoft Common Dialog.ocx
'=========================================================================================
'  Created By: Joe Hart
'  Published Date: Unknown
'  E-Mail: bghost@ti.cz
'  Legal Copyright: Joe Hart © Unknown
'=========================================================================================

Private Declare Function CommDlgExtendedError Lib "COMDLG32.DLL" () As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ChooseFont Lib "COMDLG32" Alias "ChooseFontA" (pCHOOSEFONT As TCHOOSEFONT) As Long
Private Declare Function ChooseColor Lib "COMDLG32.DLL" Alias "ChooseColorA" (pCHOOSECOLOR As TCHOOSECOLOR) As Long
Private Declare Function GetOpenFileName Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" (pOPENFILENAME As TOPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32.DLL" Alias "GetSaveFileNameA" (pOPENFILENAME As TOPENFILENAME) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)


Private Const WM_DESTROY = &H2
Private Const WM_NOTIFY = &H4E
Private Const WM_COMMAND = &H111
Private Const WM_NCDESTROY = &H82
Private Const WM_GETDLGCODE = &H87
Private Const WM_INITDIALOG = &H110


Private Const H_MAX As Long = &HFFFF + 1
Private Const CDN_LAST = (H_MAX - 699)
Private Const CDN_FIRST = (H_MAX - 601)

Private Const CDN_HELP = (CDN_FIRST - &H4)
Private Const CDN_FILEOK = (CDN_FIRST - &H5)
Private Const CDN_INITDONE = (CDN_FIRST - &H0)
Private Const CDN_SELCHANGE = (CDN_FIRST - &H1)
Private Const CDN_TYPECHANGE = (CDN_FIRST - &H6)
Private Const CDN_INCLUDEITEM = (CDN_FIRST - &H7)
Private Const CDN_FOLDERCHANGE = (CDN_FIRST - &H2)
Private Const CDN_SHAREVIOLATION = (CDN_FIRST - &H3)

Private Const MAX_FILE = 260
Private Const LF_FACESIZE = 32
Private Const SPI_GETWORKAREA = 48

Public Enum EOpenFile
    OFN_READONLY = &H1
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_NOCHANGEDIR = &H8
    OFN_NOVALIDATE = &H100
    OFN_EXPLORER = &H80000
    OFN_HIDEREADONLY = &H4
    OFN_SHAREAWARE = &H4000
    OFN_LONGNAMES = &H200000
    OFN_OVERWRITEPROMPT = &H2
    OFN_ENABLETEMPLATE = &H40
    OFN_PATHMUSTEXIST = &H800
    OFN_CREATEPROMPT = &H2000
    OFN_NOLONGNAMES = &H40000
    OFN_FILEMUSTEXIST = &H1000
    OFN_ENABLESIZING = &H800000
    OFN_ALLOWMULTISELECT = &H200
    OFN_NOREADONLYRETURN = &H8000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOTESTFILECREATE = &H10000
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOREADONLYRETURN_C = &H8000&
    OFN_NODEREFERENCELINKS = &H100000
    OFN_ENABLEINCLUDENOTIFY = &H400000
End Enum

Public Enum EChooseColor
   CC_RGBINIT = &H1
   CC_FULLOPEN = &H2
   CC_ANYCOLOR = &H100
   CC_SOLIDCOLOR = &H80
   CC_ENABLEHOOK = &H10
   CC_ColorShowHelp = &H8
   CC_PREVENTFULLOPEN = &H4
   CC_ENABLETEMPLATE = &H20
   CC_ENABLETEMPLATEHANDLE = &H40
End Enum

Public Enum EChooseFont
   CF_BOTH = &H3
   CF_APPLY = &H200
   CF_USESTYLE = &H80
   CF_EFFECTS = &H100
   CF_ENABLEHOOK = &H8
   CF_TTONLY = &H40000
   CF_ANSIONLY = &H400
   CF_WYSIWYG = &H8000
   CF_SCREENFONTS = &H1
   CF_PRINTERFONTS = &H2
   CF_FONTSHOWHELP = &H4
   CF_NOOEMFONTS = &H800
   CF_LIMITSIZE = &H2000
   CF_NOFACESEL = &H80000
   CF_NOSIZESEL = &H200000
   CF_NOSTYLESEL = &H100000
   CF_NOVECTORFONTS = &H800
   CF_ENABLETEMPLATE = &H10
   CF_NOSIMULATIONS = &H1000
   CF_SCALABLEONLY = &H20000
   CF_NOSCRIPTSEL = &H800000
   CF_SELECTSCRIPT = &H400000
   CF_FIXEDPITCHONLY = &H4000
   CF_NOVERTFONTS = &H1000000
   CF_FORCEFONTEXIST = &H10000
   CF_INITTOLOGFONTSTRUCT = &H40
   CF_ENABLETEMPLATEHANDLE = &H20
End Enum

Public Enum EFontType
    BOLD_FONTTYPE = &H100
    ITALIC_FONTTYPE = &H200
    REGULAR_FONTTYPE = &H400
    SCREEN_FONTTYPE = &H2000
    PRINTER_FONTTYPE = &H4000
    SIMULATED_FONTTYPE = &H8000
End Enum

Public Enum EDialogError
    CDERR_DIALOGFAILURE = &HFFFF
    CDERR_GENERALCODES = &H0&
    CDERR_STRUCTSIZE = &H1&
    CDERR_INITIALIZATION = &H2&
    CDERR_NOTEMPLATE = &H3&
    CDERR_NOHINSTANCE = &H4&
    CDERR_LOADSTRFAILURE = &H5&
    CDERR_FINDRESFAILURE = &H6&
    CDERR_LOADRESFAILURE = &H7&
    CDERR_LOCKRESFAILURE = &H8&
    CDERR_MEMALLOCFAILURE = &H9&
    CDERR_MEMLOCKFAILURE = &HA&
    CDERR_NOHOOK = &HB&
    CDERR_REGISTERMSGFAIL = &HC&
    CFERR_CHOOSEFONTCODES = &H2000&
    CFERR_NOFONTS = &H2001&
    CFERR_MAXLESSTHANMIN = &H2002&
    FNERR_FILENAMECODES = &H3000&
    FNERR_SUBCLASSFAILURE = &H3001&
    FNERR_INVALIDFILENAME = &H3002&
    FNERR_BUFFERTOOSMALL = &H3003&
    CCERR_CHOOSECOLORCODES = &H5000&
End Enum

Private Type TOPENFILENAME
  lStructSize As Long
  hWndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private Type TCHOOSECOLOR
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type

Private Type TCHOOSEFONT
    lStructSize As Long
    hWndOwner As Long
    hdc As Long
    lpLogFont As Long
    iPointSize As Long
    Flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    iAlign As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type NMHDR
    hWndFrom As Long
    idFrom As Long
    Code As Long
End Type

Private Type POINTL
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private m_Font As New StdFont
Private m_CancelError As Boolean
Private m_DefaultExt As String
Private m_DialogTitle As String
Private m_FileName As String
Private m_FileTitle As String
Private m_FilterIndex As Integer
Private m_Filter As String
Private m_flags As Long
Private m_InitDir As String
Private m_MaxFileSize As Integer
Private m_hWnd As Long
Private m_FileExt As Integer
Private m_fHook As Boolean
Private m_FontMinSize As Long
Private m_FontMaxSize As Long
Private m_FontColor As Long
Private m_Color As Long
Private m_ExtendedErr As Long
Private alCustom(0 To 15) As Long

Public Event InitDialog(ByVal hDlg As Long)
Public Event FileChange(ByVal hDlg As Long)
Public Event FolderChange(ByVal hDlg As Long)
Public Event DialogOK(ByRef bCancel As Boolean)
Public Event TypeChange(ByVal hDlg As Long)
Public Event DialogClose()
Public Function DialogHook(ByVal hDlg As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long)
   Dim NotifyMessage As NMHDR
  
   Select Case msg
   Case WM_INITDIALOG
      RaiseEvent InitDialog(hDlg)
   Case WM_NOTIFY
      CopyMemory NotifyMessage, ByVal lParam, Len(NotifyMessage)
      Select Case NotifyMessage.Code
      Case CDN_SELCHANGE
         RaiseEvent FileChange(hDlg)
      Case CDN_FOLDERCHANGE
         RaiseEvent FolderChange(hDlg)
      Case CDN_FILEOK
         SetWindowLong hDlg, 0&, 0&
      Case CDN_HELP
      Case CDN_TYPECHANGE
         RaiseEvent TypeChange(hDlg)
      Case CDN_INCLUDEITEM
      End Select
   Case WM_DESTROY
      RaiseEvent DialogClose
   End Select

End Function
Public Sub CenterDialog(ByVal hDlg As Long, Optional ByRef oCenterTo As Object)
   Dim lhWnd As Long
   Dim WindRect As RECT
   Dim DialogRect As RECT
   Dim tp As POINTL
   Dim hWndCenterTo As Long
   Dim lL As Long
   Dim lT As Long
   Dim lR As Long

   lhWnd = GetParent(hDlg)
   GetWindowRect lhWnd, DialogRect
   On Error Resume Next
   hWndCenterTo = oCenterTo.hwnd
   If (Err.Number = 0) Then
      GetWindowRect hWndCenterTo, WindRect
   Else
      lR = SystemParametersInfo(SPI_GETWORKAREA, 0, WindRect, 0)
      If (lR = 0) Then
         WindRect.Left = 0
         WindRect.Top = 0
         WindRect.Right = Screen.Width \ Screen.TwipsPerPixelX
         WindRect.Bottom = Screen.Height \ Screen.TwipsPerPixelY
      End If
   End If
   On Error GoTo 0
   If (WindRect.Right > 0) And (WindRect.Bottom > 0) Then
        lL = WindRect.Left + (((WindRect.Right - WindRect.Left) - (DialogRect.Right - DialogRect.Left)) \ 2)
        lT = WindRect.Top + (((WindRect.Bottom - WindRect.Top) - (DialogRect.Bottom - DialogRect.Top)) \ 2)
        MoveWindow lhWnd, lL, lT, (DialogRect.Right - DialogRect.Left), (DialogRect.Bottom - DialogRect.Top), 1
    End If
End Sub
Public Property Let FileExt(ByVal vData As Integer)
   m_FileExt = vData
End Property

Public Property Get FileExt() As Integer
   FileExt = m_FileExt
End Property
Public Property Let hwnd(ByVal vData As Long)
   m_hWnd = vData
End Property
Public Property Get hwnd() As Long
   hwnd = m_hWnd
End Property
Public Sub ShowSave()
   Dim OpenFileName As TOPENFILENAME
   Dim L As Long
   With OpenFileName
      .Flags = FileFlags
      .hWndOwner = hwnd
      .hInstance = 0
      .lCustData = 0
      .lpfnHook = 0
      .lpstrDefExt = StrPtr(DefaultExt)
      .lpstrFile = FileName & String$(MAX_FILE - Len(FileName) + 1, vbNullChar)
      .lpstrFileTitle = FileTitle & Space$(256)
      .lpstrFilter = m_Filter
      .lpstrInitialDir = InitDir
      .lpstrTitle = DialogTitle
      .lpTemplateName = 0
      .lStructSize = Len(OpenFileName)
      .nFileExtension = 0
      .nFileOffset = 0
      .nFilterIndex = FilterIndex
      .nMaxCustFilter = 0
      .nMaxFile = MAX_FILE
      .nMaxFileTitle = MAX_FILE
   End With
   L = GetSaveFileName(OpenFileName)
   Select Case L
   Case 1
      With OpenFileName
         FileFlags = .Flags
         DefaultExt = .lpstrDefExt
         FileName = NullTrim(.lpstrFile)
         FileTitle = NullTrim(.lpstrFileTitle)
         FileExt = .nFileExtension
         m_Filter = NullTrim(.lpstrFilter)
         InitDir = NullTrim(.lpstrInitialDir)
         FilterIndex = NullTrim(.nFilterIndex)
      End With
   Case 0
      If CancelError Then Err.Raise 1001, "Run-time error", "Cancel was selected"
   Case Else
        m_ExtendedErr = CommDlgExtendedError()
        Err.Raise m_ExtendedErr
   End Select
End Sub
Public Sub ShowOpen()
   Dim ofn As TOPENFILENAME
   Dim L As Long
   With ofn
      .Flags = m_flags
      .hWndOwner = m_hWnd
      .hInstance = 0
      .lpfnHook = 0
      .lCustData = 0
      .lpstrDefExt = m_DefaultExt
      .lpstrFile = m_FileName & String$(MAX_FILE - Len(m_FileName) + 1, 0)
      .lpstrFileTitle = m_FileTitle & Space$(256)
      .lpstrFilter = m_Filter
      .lpstrInitialDir = m_InitDir
      .lpstrTitle = m_DialogTitle
      .lpTemplateName = 0
      .lStructSize = Len(ofn)
      .nFileExtension = 0
      .nFileOffset = 0
      .nFilterIndex = m_FilterIndex
      .nMaxCustFilter = 0
      .nMaxFile = MAX_FILE
      .nMaxFileTitle = MAX_FILE
      If m_fHook Then
         HookedDialog = Me
         .lpfnHook = HookAddress(AddressOf DialogHookFunction)
         .Flags = .Flags Or OFN_ENABLEHOOK Or OFN_EXPLORER
      End If
   End With
   L = GetOpenFileName(ofn)
   ClearHookedDialog
   Select Case L
   Case 1
      With ofn
         m_flags = .Flags
         m_DefaultExt = .lpstrDefExt
         m_FileName = NullTrim(.lpstrFile)
         m_FileTitle = NullTrim(.lpstrFileTitle)
         m_FileExt = .nFileExtension
         m_Filter = NullTrim(.lpstrFilter)
         m_InitDir = NullTrim(.lpstrInitialDir)
         m_FilterIndex = NullTrim(.nFilterIndex)
      End With
   Case 0
      If m_CancelError Then Err.Raise 1002, "Run-time error", "Cancel was selected"
   Case Else
      m_ExtendedErr = CommDlgExtendedError()
      Err.Raise m_ExtendedErr
   End Select
End Sub
Public Sub ShowFont()
   Dim PrinterDC As Long
   Dim L As Long
    Const CF_FontNotSupported = CF_APPLY Or CF_ENABLEHOOK Or CF_ENABLETEMPLATE
    
    If m_flags And CF_PRINTERFONTS Then PrinterDC = Printer.hdc
    If (m_flags And CF_PRINTERFONTS) = 0 Then m_flags = m_flags Or CF_SCREENFONTS
    
    If m_FontColor > 0 Then m_flags = m_flags Or CF_EFFECTS
    
    If m_FontMinSize > 0 Or m_FontMaxSize > 0 Then m_flags = m_flags Or CF_LIMITSIZE
    
    m_flags = (m_flags Or CF_INITTOLOGFONTSTRUCT) And Not CF_FontNotSupported
    
    Dim LogFnt As LOGFONT
    Const PointsPerTwip = 1440 / 72
    LogFnt.lfHeight = -(m_Font.Size * (PointsPerTwip / Screen.TwipsPerPixelY))
    LogFnt.lfWeight = m_Font.Weight
    LogFnt.lfItalic = m_Font.Italic
    LogFnt.lfUnderline = m_Font.Underline
    LogFnt.lfStrikeOut = m_Font.Strikethrough
    StrToBytes LogFnt.lfFaceName, m_Font.Name

    Dim ChooseFnt As TCHOOSEFONT
    With ChooseFnt
      .lStructSize = Len(ChooseFnt)
      .hWndOwner = m_hWnd
      .hdc = PrinterDC
      .lpLogFont = VarPtr(LogFnt)
      .iPointSize = m_Font.Size * 10
      .Flags = m_flags
      .rgbColors = Color
      .nSizeMin = m_FontMinSize
      .nSizeMax = m_FontMaxSize
    End With
    
    L = ChooseFont(ChooseFnt)
    Select Case L
    Case 1
        m_flags = ChooseFnt.Flags
        m_FontColor = ChooseFnt.rgbColors
        m_Font.Bold = ChooseFnt.nFontType And BOLD_FONTTYPE
        m_Font.Italic = LogFnt.lfItalic
        m_Font.Strikethrough = LogFnt.lfStrikeOut
        m_Font.Underline = LogFnt.lfUnderline
        m_Font.Weight = LogFnt.lfWeight
        m_Font.Size = ChooseFnt.iPointSize / 10
        m_Font.Name = StrConv(LogFnt.lfFaceName, vbUnicode)
    Case 0
      If m_CancelError Then Err.Raise 1001, "Run-time error", "Cancel was selected"
    Case Else
        m_ExtendedErr = CommDlgExtendedError()
        Err.Raise m_ExtendedErr
   End Select
        
End Sub
Sub ShowColor()
    Dim ChooseClr As TCHOOSECOLOR
    Dim afMask As Long
    Dim L As Long
    
    With ChooseClr
      .lStructSize = Len(ChooseClr)
    
      .hWndOwner = m_hWnd
      .rgbResult = m_Color
      
      afMask = CLng(Not (CC_ENABLEHOOK Or _
                       CC_ENABLETEMPLATE))
      .Flags = afMask And m_flags
      .lpCustColors = VarPtr(alCustom(0))
    End With
   L = ChooseColor(ChooseClr)
    
    Select Case L
    Case 1
        m_Color = ChooseClr.rgbResult
    Case 0
        If m_CancelError = True Then Err.Raise 1004, , "Choose Color Dialog was canceled"
        m_Color = -1
    Case Else
        m_ExtendedErr = CommDlgExtendedError()
        Err.Raise m_ExtendedErr
    End Select

End Sub
Private Sub InitColors()
    Dim i As Integer
    For i = 0 To 15
        alCustom(i) = GetSysColor(i)
    Next
End Sub

Public Property Get CustomColor(i As Integer) As Long
    If i >= 0 And i <= 15 Then
        CustomColor = alCustom(i)
    Else
        CustomColor = -1
    End If
End Property
Public Property Let Color(NewColor As Long)
   m_Color = NewColor
End Property
Public Property Get Color() As Long
    Color = m_Color
End Property
Public Property Let FontColor(NewColor As Long)
   m_FontColor = NewColor
End Property
Public Property Get FontColor() As Long
   FontColor = m_FontColor
End Property
Public Property Let FontMinSize(MinSize As Long)
   m_FontMinSize = MinSize
End Property
Public Property Let FontMaxSize(MaxSize As Long)
   m_FontMaxSize = MaxSize
End Property
Private Function HookAddress(Pointer As Long) As Long
    HookAddress = Pointer
End Function
Public Property Let InitDir(ByVal vData As String)
   m_InitDir = vData
End Property
Public Property Get InitDir() As String
   InitDir = m_InitDir
End Property
Public Property Let FileFlags(ByVal vData As EOpenFile)
   m_flags = vData
End Property
Public Property Let Flags(NewFlags As Long)
   m_flags = NewFlags
End Property
Public Property Get Flags() As Long
   Flags = m_flags
End Property
Public Property Get FileFlags() As EOpenFile
   FileFlags = m_flags
End Property
Public Property Let FontFlags(ByVal vData As EChooseFont)
   m_flags = vData
End Property
Public Property Get FontFlags() As EChooseFont
   FontFlags = m_flags
End Property
Public Property Let ColorFlags(ByVal vData As EChooseColor)
   m_flags = vData
End Property
Public Property Get ColorFlags() As EChooseColor
   ColorFlags = m_flags
End Property
Public Property Let Filter(ByVal vData As String)
   Dim pipepos As String
   Do While InStr(vData, "|") > 0
      pipepos = InStr(vData, "|")
      If pipepos > 0 Then
         vData = Left$(vData, pipepos - 1) & vbNullChar & Right$(vData, Len(vData) - pipepos)
      End If
   Loop
   If Right$(vData, 2) <> vbNullChar & vbNullChar Then vData = vData & vbNullChar
   If Right$(vData, 2) <> vbNullChar & vbNullChar Then vData = vData & vbNullChar
   m_Filter = vData
End Property
Public Property Get Filter() As String
   Dim nullpos As String
   Dim tempfilter As String
   tempfilter = m_Filter
   Do While InStr(tempfilter, vbNullChar) > 0
      nullpos = InStr(tempfilter, vbNullChar)
      If nullpos > 0 Then
         tempfilter = Left$(tempfilter, nullpos - 1) & vbNullChar & Right$(tempfilter, Len(tempfilter) - nullpos)
      End If
   Loop
   If Right$(tempfilter, 1) = "|" Then tempfilter = Left$(tempfilter, Len(tempfilter) - 1)
   If Right$(tempfilter, 1) = "|" Then tempfilter = Left$(tempfilter, Len(tempfilter) - 1)
   Filter = tempfilter
End Property
Public Property Let FilterIndex(ByVal vData As Integer)
   m_FilterIndex = vData
End Property
Public Property Get FilterIndex() As Integer
   FilterIndex = m_FilterIndex
End Property
Public Property Let FileTitle(ByVal vData As String)
   m_FileTitle = vData
End Property
Public Property Get FileTitle() As String
   FileTitle = m_FileTitle
End Property
Public Property Let FileName(ByVal vData As String)
   m_FileName = vData
End Property
Public Property Get FileName() As String
   FileName = m_FileName
End Property
Public Property Let DialogTitle(ByVal vData As String)
   m_DialogTitle = vData
End Property
Public Property Get DialogTitle() As String
   DialogTitle = m_DialogTitle
End Property
Public Property Let DefaultExt(ByVal vData As String)
   m_DefaultExt = vData
End Property
Public Property Get DefaultExt() As String
   DefaultExt = m_DefaultExt
End Property
Public Property Let CancelError(ByVal vData As Boolean)
   m_CancelError = vData
End Property
Public Property Get CancelError() As Boolean
   CancelError = m_CancelError
End Property
Private Sub StrToBytes(ab() As Byte, s As String)
    If IsArrayEmpty(ab) Then
        ab = StrConv(s, vbFromUnicode)
    Else
        Dim cab As Long
        cab = UBound(ab) - LBound(ab) + 1
        If Len(s) < cab Then s = s & String$(cab - Len(s), 0)
        CopyMemoryStr ab(LBound(ab)), s, cab
    End If
End Sub
Public Property Get FontBold() As Boolean
    FontBold = m_Font.Bold
End Property
Public Property Let FontBold(ByVal vNewValue As Boolean)
    m_Font.Bold = vNewValue
End Property
Public Property Get FontItalic() As Boolean
    FontItalic = m_Font.Italic
End Property
Public Property Let FontItalic(ByVal vNewValue As Boolean)
    m_Font.Italic = vNewValue
End Property
Public Property Get FontName() As String
    FontName = m_Font.Name
End Property
Public Property Let FontName(ByVal vNewValue As String)
    m_Font.Name = vNewValue
End Property
Public Property Get FontSize() As Long
    FontSize = m_Font.Size
End Property
Public Property Let FontSize(ByVal vNewValue As Long)
    m_Font.Size = vNewValue
End Property
Public Property Get Font() As StdFont
    Set Font = m_Font
End Property
Public Property Let Font(sFont As StdFont)
    Set m_Font = sFont
End Property
Private Sub Class_Initialize()
   CancelError = False
   DefaultExt = ""
   DialogTitle = ""
   FileName = ""
   FileTitle = ""
   Filter = "All Files|*.*"
   FilterIndex = 1
   hwnd = 0
   InitColors
End Sub
Private Function IsArrayEmpty(va As Variant) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = va(LBound(va))
    IsArrayEmpty = (Err <> 0)
End Function
Public Property Get Hook() As Boolean
       Hook = m_fHook
End Property
Public Property Let Hook(NewHook As Boolean)
       m_fHook = NewHook
End Property


