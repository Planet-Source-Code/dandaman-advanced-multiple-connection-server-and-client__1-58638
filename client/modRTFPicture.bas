Attribute VB_Name = "modRTFPicture"
' **************************************************************************
'  EMBEDDING METAFILE and RTF PICTURE EXAMPLE
' **************************************************************************
'
'    AUTHOR: The Hand
'      DATE: June, 2002
'   COMPANY: EliteVB
'
' DESCRIPTION:
'    This example shows the user how to embed a picture (StdPicture object
'    into a metafile, and subsequently create usable RTF code so it can be
'    placed in a rich text box.
'
'    Forget those horribly cheesy Clipboard and OLEObject.Add methods and
'    use this method instead!
'
'    Feel free to use this source in your own projects. You are not allowed
'    to take credit for it, publish it, or wave it around on PSC trying
'    to win a prize without prior consent from the EliteVB team. And just
'    so you know, a few global "replace alls" does not make it 'YOUR' source
'    code. It just makes you a serious lamer, and a very sad individual
'    with little creativity. Give us credit where its due.
'
' **************************************************************************
'   Visit EliteVB.com for more high-powered API and subclassing solutions!
' **************************************************************************

Option Explicit

Private Type Size
    cx As Long
    cy As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

'Private Type METAHEADER
'    mtType As Integer
'    mtHeaderSize As Integer
'    mtVersion As Integer
'    mtSize As Long
'    mtNoObjects As Integer
'    mtMaxRecord As Long
'    mtNoParameters As Integer
'End Type

' Used to create the metafile
Private Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Private Declare Function CloseMetaFile Lib "gdi32" (ByVal hDCMF As Long) As Long
Private Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
' 6 APIs used to render/embed the bitmap in the metafile
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Private Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Private Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
' These APIs are used to BitBlt the bitmap image into the metafile
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

' Used for creating the temporary WMF file
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Const MM_ANISOTROPIC = 8 ' Map mode anisotropic

Public Function StdPicAsRTF(aStdPic As StdPicture) As String

    ' ***********************************************************************
    '  Author: The Hand
    '    Date: June, 2002
    ' Company: EliteVB
    '
    '  Function: StdPicAsRTF
    ' Arguments: aStdPic - Any standard picture object from memory, a
    '                      picturebox, or other source.
    '
    ' Description:
    '    Embeds a standard picture object in a windows metafile and returns
    '    rich text format code (RTF) so it can be placed in a RichTextBox.
    '    Useful for emoticons in chat programs, pics, etc. Currently does
    '    not support icon files, but that is easy enough to add in.
    ' ***********************************************************************
    Dim hMetaDC     As Long
    Dim hMeta       As Long
    Dim hPicDC      As Long
    Dim hOldBmp     As Long
    Dim aBMP        As BITMAP
    Dim aSize       As Size
    Dim aPt         As POINTAPI
    Dim FileName    As String
'    Dim aMetaHdr    As METAHEADER
    Dim screenDC    As Long
    Dim headerStr   As String
    Dim retStr      As String
    Dim byteStr     As String
    Dim bytes()     As Byte
    Dim filenum     As Integer
    Dim numBytes    As Long
    Dim i           As Long
    
    ' Create a metafile to a temporary file in the registered windows TEMP folder
    FileName = getTempName("WMF")
    hMetaDC = CreateMetaFile(FileName)
    
    ' Set the map mode to MM_ANISOTROPIC
    SetMapMode hMetaDC, MM_ANISOTROPIC
    ' Set the metafile origin as 0, 0
    SetWindowOrgEx hMetaDC, 0, 0, aPt
    ' Get the bitmap's dimensions
    GetObject aStdPic.Handle, Len(aBMP), aBMP
    ' Set the metafile width and height
    SetWindowExtEx hMetaDC, aBMP.bmWidth, aBMP.bmHeight, aSize
    ' save the new dimensions
    SaveDC hMetaDC
    ' OK. Now transfer the freakin image to the metafile
    screenDC = GetDC(0)
    hPicDC = CreateCompatibleDC(screenDC)
    ReleaseDC 0, screenDC
    hOldBmp = SelectObject(hPicDC, aStdPic.Handle)
    BitBlt hMetaDC, 0, 0, aBMP.bmWidth, aBMP.bmHeight, hPicDC, 0, 0, vbSrcCopy
    SelectObject hPicDC, hOldBmp
    DeleteDC hPicDC
    DeleteObject hOldBmp
    ' "redraw" the metafile DC
    RestoreDC hMetaDC, True
    ' close it and get the metafile handle
    hMeta = CloseMetaFile(hMetaDC)
    
'    GetObject hMeta, Len(aMetaHdr), aMetaHdr
    ' delete it from memory
    DeleteMetaFile hMeta
    
    ' Do the RTF header for the object. This little bit is sometimes required on
    '  earlier versions of the rich text box and in certain operating systems
    '  (WinNT springs to mind)
    headerStr = "{\rtf1\ansi"
    ' Picture specific tag stuff
    headerStr = headerStr & _
                "{\pict\picscalex100\picscaley100" & _
                "\picw" & aStdPic.Width & "\pich" & aStdPic.Height & _
                "\picwgoal" & aBMP.bmWidth * Screen.TwipsPerPixelX & _
                "\pichgoal" & aBMP.bmHeight * Screen.TwipsPerPixelY & _
                "\wmetafile8"
    
    ' Get the size of the metafile
    numBytes = FileLen(FileName)
    ' Create our byte buffer for reading
    ReDim bytes(1 To numBytes)
    ' get a free file number
    filenum = FreeFile()
    ' open the file for input
    Open FileName For Binary Access Read As #filenum
    ' read the bytes
    Get #filenum, , bytes
    ' close the file
    Close #filenum
    ' Generate our hex encoded byte string
    byteStr = String(numBytes * 2, "0")
    For i = LBound(bytes) To UBound(bytes)
        If bytes(i) > &HF Then
            Mid$(byteStr, 1 + (i - 1) * 2, 2) = Hex$(bytes(i))
        Else
            Mid$(byteStr, 2 + (i - 1) * 2, 1) = Hex$(bytes(i))
        End If
    Next i
    ' stick it all together
    retStr = headerStr & " " & byteStr & "}"
    ' Add in the closing RTF bit
    retStr = retStr & "}"
        
    StdPicAsRTF = retStr
    On Local Error Resume Next
    ' Kill the temporary file
    If Dir(FileName) <> "" Then Kill FileName
End Function

Private Function getTempName(Optional anExt As String = "tmp") As String
    ' ***********************************************************************
    '  Author: The Hand
    '    Date: June, 2002
    ' Company: EliteVB
    '
    '  Function: getTempName
    ' Arguments: anExt - an extension to be used for the temp file. If none
    '                    is provided, the function automatically uses "tmp"
    '                    as the extension. It is up to the procedure that
    '                    uses this temporary name to clean up the file (kill
    '                    it) after it is created.
    '
    ' Description:
    '    Creates a temporary filename in the registered system temp directory
    ' ***********************************************************************
    Dim tempPath    As String
    Dim FileName    As String
    Dim i           As Long
    
    Const validChars As String = "123567890qwertyuiopasdfghjklzxcvbnm"
    
    ' Create a buffer
    tempPath = String$(255, " ")
    ' get the system path
    GetTempPath 255, tempPath
    ' trim off the fat
    tempPath = Left$(tempPath, InStr(tempPath, Chr$(0)) - 1)
    ' Create a buffer
    FileName = Space(12)
    ' Put the non-random stuff into the string
    Mid$(FileName, 1, 1) = "T"
    Mid$(FileName, Len(FileName) - Len(anExt), 1) = "."
    ' Add in the specified extension, if provided ("tmp" is default)
    Mid$(FileName, Len(FileName) - Len(anExt) + 1, Len(anExt)) = anExt
    ' fill the buffer with random stuff
    Randomize
    For i = 2 To Len(FileName) - 4
        Mid$(FileName, i, 1) = Mid$(validChars, CLng(Rnd() * (Len(validChars)) + 1), 1)
    Next i
    tempPath = tempPath & FileName
    ' return the path name
    getTempName = tempPath
    
End Function
'Private Sub Command1_Click()
'    Dim aStr As String
'    aStr = StdPicAsRTF(Picture1.Picture)
'    RichTextBox1.SelRTF = aStr
'End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    RichTextBox1.Text = ""
'End Sub

'Private Sub Form_Resize()
'    Command1.Move 0, Me.ScaleHeight - Command1.Height, Me.ScaleWidth
'    Picture1.Move 0, 0, Me.ScaleWidth / 2, Me.ScaleHeight - Command1.Height
'    RichTextBox1.Move Me.ScaleWidth / 2, 0, Me.ScaleWidth / 2, Me.ScaleHeight - Command1.Height
'End Sub



