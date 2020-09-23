Attribute VB_Name = "MBorrowedCode"
Option Explicit


Public Declare Function recv Lib "WSOCK32.DLL" (ByVal s As Long, ByVal buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long

'Following code from Bruce McKinney's Hardcore VisualBasic
'***Begin******************************************************

Public Const sEmpty As String = ""  'type lib doesn't like this one
Public Const sQuote2 = """"   'or this one
Public Const sBSlash = "\"

Private Declare Function StrSpn Lib "SHLWAPI" Alias "StrSpnW" ( _
        ByVal psz As Long, ByVal pszSet As Long) As Long
Private Declare Function StrCSpn Lib "SHLWAPI" Alias "StrCSpnW" ( _
        ByVal LPSTR As Long, ByVal lpSet As Long) As Long
        

Private Declare Function GetOpenFileName Lib "comdlg32.dll" _
    Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32" _
    Alias "GetSaveFileNameA" (File As OPENFILENAME) As Long

'need this private copymem declare here because fontname won't copy to dlg without it
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' In standard module fNoShlWapi is a Property Get that checks for DLL
Private fNotFirstTime As Boolean, fNoShlWapiI As Boolean ' = False
' Array of custom colors lasts for life of app
Private alCustom(0 To 15) As Long, fNotFirst As Boolean
Public Sock As Integer, WSAStartedUp As Boolean     'Flag to keep track of whether winsock WSAStartup wascalled
Private m_CurrentDirectory As String



Private Property Get fNoShlWapi() As Boolean
    If fNotFirstTime = False Then
        fNotFirstTime = True
        On Error GoTo Fail
        Call StrSpn(StrPtr("a"), StrPtr("a"))
    End If
    Exit Property
Fail:
    fNoShlWapiI = True
End Property

Public Function VBGetOpenFileName(FileName As String, _
                           Optional FileTitle As String, _
                           Optional FileMustExist As Boolean = True, _
                           Optional MultiSelect As Boolean = False, _
                           Optional ReadOnly As Boolean = False, _
                           Optional HideReadOnly As Boolean = False, _
                           Optional filter As String = "All (*.*)| *.*", _
                           Optional FilterIndex As Long = 1, _
                           Optional InitDir As String, _
                           Optional DlgTitle As String, _
                           Optional DefaultExt As String, _
                           Optional owner As Long = -1, _
                           Optional flags As Long = 0) As Boolean

    Dim opfile As OPENFILENAME, s As String, afFlags As Long
    With opfile
        .lStructSize = Len(opfile)
        
        ' Add in specific flags and strip out non-VB flags
        .flags = (-FileMustExist * OFN_FILEMUSTEXIST) Or _
                 (-MultiSelect * OFN_ALLOWMULTISELECT) Or _
                 (-ReadOnly * OFN_READONLY) Or _
                 (-HideReadOnly * OFN_HIDEREADONLY) Or _
                 (flags And CLng(Not (OFN_ENABLEHOOK Or _
                                      OFN_ENABLETEMPLATE)))
        ' Owner can take handle of owning window
        If owner <> -1 Then .hwndOwner = owner
        ' InitDir can take initial directory string
        .lpstrInitialDir = InitDir
        ' DefaultExt can take default extension
        .lpstrDefExt = DefaultExt
        ' DlgTitle can take dialog box title
        .lpstrTitle = DlgTitle

        ' To make Windows-style filter, replace | and : with nulls
        Dim ch As String, i As Long
        For i = 1 To Len(filter)
            ch = Mid$(filter, i, 1)
            If ch = "|" Or ch = ":" Then
                s = s & vbNullChar
            Else
                s = s & ch
            End If
        Next
        ' Put double null at end
        s = s & vbNullChar & vbNullChar
        .lpstrFilter = s
        .nFilterIndex = FilterIndex
    
        ' Pad file and file title buffers to maximum path
        s = FileName & String$(cMaxPath - Len(FileName), 0)
        .lpstrFile = s
        .nMaxFile = cMaxPath
        s = FileTitle '& String$(cMaxFile - Len(FileTitle), 0)
        .lpstrFileTitle = s
        .nMaxFileTitle = cMaxFile
        ' All other fields set to zero

        If GetOpenFileName(opfile) Then
            VBGetOpenFileName = True
            FileName = StrZToStr(.lpstrFile)
            FileTitle = StrZToStr(.lpstrFileTitle)
            flags = .flags
            ' Return the filter index
            FilterIndex = .nFilterIndex
            ' Look up the filter the user selected and return that
            filter = FilterLookup(.lpstrFilter, FilterIndex)
            If (.flags And OFN_READONLY) Then ReadOnly = True
        Else
            VBGetOpenFileName = False
            FileName = sEmpty
            FileTitle = sEmpty
            flags = 0
            FilterIndex = -1
            filter = sEmpty
        End If
    End With
End Function
' ChooseColor wrapper
Function VBChooseColor(Color As Long, _
                       Optional AnyColor As Boolean = True, _
                       Optional FullOpen As Boolean = False, _
                       Optional DisableFullOpen As Boolean = False, _
                       Optional owner As Long = -1, _
                       Optional flags As Long) As Boolean

    Dim chclr As TCHOOSECOLOR
    chclr.lStructSize = Len(chclr)
    
    ' Color must get reference variable to receive result
    ' Flags can get reference variable or constant with bit flags
    ' Owner can take handle of owning window
    If owner <> -1 Then chclr.hwndOwner = owner

    ' Assign color (default uninitialized value of zero is good default)
    chclr.rgbResult = Color

    ' Mask out unwanted bits
    Dim afMask As Long
    afMask = CLng(Not (CC_ENABLEHOOK Or _
                       CC_ENABLETEMPLATE))
    ' Pass in flags
    chclr.flags = afMask And (CC_RGBINIT Or _
                  IIf(AnyColor, CC_ANYCOLOR, CC_SOLIDCOLOR) Or _
                  (-FullOpen * CC_FULLOPEN) Or _
                  (-DisableFullOpen * CC_PREVENTFULLOPEN))

    ' If first time, initialize to white
    If fNotFirst = False Then InitColors

    chclr.lpCustColors = VarPtr(alCustom(0))
    ' All other fields zero
    
    If ChooseColor(chclr) Then
        VBChooseColor = True
        Color = chclr.rgbResult
    Else
        VBChooseColor = False
        Color = -1
    End If

End Function

Private Sub InitColors()
    Dim i As Long
    ' Initialize with first 16 system interface colors
    For i = 0 To 15
        alCustom(i) = GetSysColor(i)
    Next
    fNotFirst = True
End Sub

' Property to read or modify custom colors (use to save colors in registry)
Public Property Get CustomColor(i As Integer) As Long
    ' If first time, initialize to white
    If fNotFirst = False Then InitColors
    If i >= 0 And i <= 15 Then
        CustomColor = alCustom(i)
    Else
        CustomColor = -1
    End If
End Property

Public Property Let CustomColor(i As Integer, iValue As Long)
    ' If first time, initialize to system colors
    If fNotFirst = False Then InitColors
    If i >= 0 And i <= 15 Then
        alCustom(i) = iValue
    End If
End Property

' ChooseFont wrapper   **** modified from original which works of in a dll but
'would not default the original fontname into the dlg as a module function.....
Function VBChooseFont(CurFont As Font, _
                      Optional PrinterDC As Long = -1, _
                      Optional owner As Long = -1, _
                      Optional Color As Long = vbBlack, _
                      Optional MinSize As Long = 0, _
                      Optional MaxSize As Long = 0, _
                      Optional flags As Long = 0) As Boolean

    Dim hMem As Long, pMem As Long, RetVal As Long   ' handle and pointer to memory buffer

    ' Unwanted Flags bits
    Const CF_FontNotSupported = CF_APPLY Or CF_ENABLEHOOK Or CF_ENABLETEMPLATE 'Or CF_NOFACESEL

    ' Flags can get reference variable or constant with bit flags
    ' PrinterDC can take printer DC
    If PrinterDC = -1 Then
        PrinterDC = 0
        If flags And CF_PRINTERFONTS Then PrinterDC = Printer.hdc
    Else
        flags = flags Or CF_PRINTERFONTS
    End If
    ' Must have some fonts
    If (flags And CF_PRINTERFONTS) = 0 Then flags = flags Or CF_SCREENFONTS
    ' Color can take initial color, receive chosen color
    If Color <> vbBlack Then flags = flags Or CF_EFFECTS
    ' MinSize can be minimum size accepted
    If MinSize Then flags = flags Or CF_LIMITSIZE
    ' MaxSize can be maximum size accepted
    If MaxSize Then flags = flags Or CF_LIMITSIZE

    ' Put in required internal flags and remove unsupported
    flags = (flags Or CF_INITTOLOGFONTSTRUCT) And Not CF_FontNotSupported

    ' Initialize LOGFONT variable
    Dim fnt As LOGFONT
    Const PointsPerTwip = 1440 / 72
    fnt.lfHeight = -(CurFont.Size * (PointsPerTwip / Screen.TwipsPerPixelY))
    fnt.lfWeight = CurFont.Weight
    fnt.lfItalic = CurFont.Italic
    fnt.lfUnderline = CurFont.Underline
    fnt.lfStrikeOut = CurFont.Strikethrough

    ' Other fields zero
'''    StrToBytes fnt.lfFaceName, CurFont.Name & vbNullChar  'tLib use
    fnt.lfFaceName = CurFont.Name & vbNullChar

    'added
    ' Create the memory block which will act as the LOGFONT structure buffer.
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(fnt))
    pMem = GlobalLock(hMem)  ' lock and get pointer
    CopyMemory ByVal pMem, fnt, ByVal Len(fnt)  ' copy structure's contents into block


    ' Initialize TCHOOSEFONT variable
    Dim cf As TCHOOSEFONT
    cf.lStructSize = Len(cf)
    If owner <> -1 Then cf.hwndOwner = owner
    cf.hdc = PrinterDC
    cf.lpLogFont = pMem      'VarPtr(fnt)
    cf.iPointSize = CurFont.Size * 10
    cf.flags = flags
    cf.rgbColors = Color
    cf.nSizeMin = MinSize
    cf.nSizeMax = MaxSize
    ' All other fields zero

    If ChooseFont(cf) Then
        VBChooseFont = True
        'added
        CopyMemory fnt, ByVal pMem, ByVal Len(fnt)  ' copy memory back

        flags = cf.flags
        Color = cf.rgbColors
        CurFont.Bold = cf.nFontType And BOLD_FONTTYPE
        'CurFont.Italic = cf.nFontType And ITALIC_FONTTYPE
        CurFont.Italic = fnt.lfItalic
        CurFont.Strikethrough = fnt.lfStrikeOut
        CurFont.Underline = fnt.lfUnderline
        CurFont.Weight = fnt.lfWeight
        CurFont.Size = cf.iPointSize / 10
'''        CurFont.Name = BytesToStr(fnt.lfFaceName)  'typeLib use
        ' Now make the fixed-length string holding the font name into a "normal" string.
        CurFont.Name = Left(fnt.lfFaceName, InStr(fnt.lfFaceName, vbNullChar) - 1)
'Debug.Print CurFont.Name
    Else
        VBChooseFont = False
    End If
    'added  ' Deallocate the memory block we created earlier.  Note that this must
    ' be done whether the function succeeded or not.
    RetVal = GlobalUnlock(hMem)  ' destroy pointer, unlock block
    RetVal = GlobalFree(hMem)  ' free the allocated memory
End Function


Public Sub StrToBytes(ab() As Byte, s As String, Optional bForceUniCode As Boolean = False)
    If IsArrayEmpty(ab) Then
        ' Assign to empty array
        ab = StrConv(s, vbFromUnicode)
Debug.Print "empty ab()"
    Else
        Dim cab As Long
        ' Copy to existing array, padding or truncating if necessary
        cab = UBound(ab) - LBound(ab) + 1
        If Len(s) < cab Then s = s & String$(cab - Len(s), 0)
        If UnicodeTypeLib Or bForceUniCode Then
            Dim st As String
            st = StrConv(s, vbFromUnicode)
            'CopyMemoryStr ab(LBound(ab)), st, ByVal cab
CopyMemory ab(LBound(ab)), st, ByVal cab
        Else
            'CopyMemoryStr ab(LBound(ab)), s, ByVal cab
CopyMemory ab(LBound(ab)), s, ByVal cab
        End If
    End If
    
'Dim i As Integer
'For i = LBound(ab) To UBound(ab)
'Debug.Print Chr$(ab(i))
'Next

End Sub
' Convert an ANSI string in a byte array to a VB Unicode string
Function BytesToStr(ab() As Byte) As String
    BytesToStr = StrConv(ab, vbUnicode)
End Function
' Strip junk at end from null-terminated string
Function StrZToStr(s As String) As String
    StrZToStr = Left$(s, lstrlen(s))
End Function
' Test file existence with error trapping
Public Function ExistFile(sSpec As String) As Boolean
    On Error Resume Next
    Call FileLen(sSpec)
    ExistFile = (Err = 0)
End Function
' Fix provided by Torsten Rendelmann
Function IsArrayEmpty(va As Variant) As Boolean
    Dim i As Long
    On Error Resume Next
    i = LBound(va, 1)
    IsArrayEmpty = (Err <> 0)
    Err = 0
End Function
Private Function FilterLookup(ByVal sFilters As String, ByVal iCur As Long) As String
    Dim iStart As Long, iEnd As Long, s As String
    iStart = 1
    If sFilters = sEmpty Then Exit Function
    Do
        ' Cut out both parts marked by null character
        iEnd = InStr(iStart, sFilters, vbNullChar)
        If iEnd = 0 Then Exit Function
        iEnd = InStr(iEnd + 1, sFilters, vbNullChar)
        If iEnd Then
            s = Mid$(sFilters, iStart, iEnd - iStart)
        Else
            s = Mid$(sFilters, iStart)
        End If
        iStart = iEnd + 1
        If iCur = 1 Then
            FilterLookup = s
            Exit Function
        End If
        iCur = iCur - 1
    Loop While iCur
End Function
' New GetQToken uses faster StrSpn and StrCSpn from SHLWAPI.DLL
Public Function GetQToken(sTarget As String, sSeps As String) As String
    ' GetQToken = sEmpty
    If fNoShlWapi Then
        GetQToken = GetQTokenO(sTarget, sSeps)
        Exit Function
    End If
    
    ' Note that sSave, pSave, pCur, and cSave must be static between calls
    Static sSave As String, pSave As Long, pCur As Long, cSave As Long
    ' First time through save start and length of string
    If sTarget <> sEmpty Then
        ' Save in case sTarget is moveable string (Command$)
        sSave = sTarget
        pSave = StrPtr(sSave)
        pCur = pSave
        cSave = Len(sSave)
    Else
        ' Quit if past end (also catches null or empty target)
        If pCur >= pSave + (cSave * 2) Then Exit Function
    End If
    ' Make sure separators includes quote
    Dim sSepsNew As String, pSeps As Long
    sSepsNew = sSeps & sQuote2
    pSeps = StrPtr(sSepsNew)

    ' Get current character
    Dim pNew As Long, c As Long
    ' Find start of next token
    c = StrSpn(pCur, pSeps)
    ' Set position to start of token
    If c Then pCur = pCur + (c * 2)
    
    Dim ch As Integer
    Const chQuote = 34  ' Asc("""")
    CopyMemory ch, ByVal pCur - 2, 2
    ' Check first character for quote, then find end of token
    If ch = chQuote Then
        c = StrCSpn(pCur, StrPtr(sQuote2))
    Else
        c = StrCSpn(pCur, pSeps)
    End If
    ' If token length is zero, we're at end
    If c = 0 Then Exit Function
    
    ' Cut token out of target string
    GetQToken = String$(c, 0)
    CopyMemory ByVal StrPtr(GetQToken), ByVal pCur, c * 2
    ' Set new starting position
    pCur = pCur + (c * 2)

End Function
' GetQTokenO uses our StrSpan and StrBreak
Private Function GetQTokenO(sTarget As String, sSeps As String) As String
    ' GetQTokenO = sEmpty

    ' Note that sSave and iStart must be static from call to call
    ' If first call, make copy of string
    Static sSave As String, iStart As Integer, cSave As Integer
    Dim iNew As Integer, fQuote As Integer
    If sTarget <> sEmpty Then
        iStart = 1
        sSave = sTarget
        cSave = Len(sSave)
    Else
        If sSave = sEmpty Then Exit Function
    End If
    ' Make sure separators includes quote
    sSeps = sSeps & sQuote2

    ' Find start of next token
    iNew = StrSpan(sSave, iStart, sSeps)
    If iNew Then
        ' Set position to start of token
        iStart = iNew
    Else
        ' If no new token, return empty string
        sSave = sEmpty
        Exit Function
    End If
    
    ' Find end of token
    If iStart = 1 Then
        iNew = StrBreak(sSave, iStart, sSeps)
    ElseIf Mid$(sSave, iStart - 1, 1) = sQuote2 Then
        iNew = StrBreak(sSave, iStart, sQuote2)
    Else
        iNew = StrBreak(sSave, iStart, sSeps)
    End If

    If iNew = 0 Then
        ' If no end of token, set to end of string
        iNew = cSave + 1
    End If
    ' Cut token out of sTarget string
    GetQTokenO = Mid$(sSave, iStart, iNew - iStart)
    
    ' Set new starting position
    iStart = iNew

End Function
' StrBreak and StrSpan are used by GetTokenO, but can be called by clients
Function StrBreak(sTarget As String, ByVal iStart As Integer, _
                  sSeps As String) As Integer
    
    Dim cTarget As Integer
    cTarget = Len(sTarget)
   
    ' Look for end of token (first character that is a separator)
    Do While InStr(sSeps, Mid$(sTarget, iStart, 1)) = 0
        If iStart > cTarget Then
            StrBreak = 0
            Exit Function
        Else
            iStart = iStart + 1
        End If
    Loop
    StrBreak = iStart

End Function
Function StrSpan(sTarget As String, ByVal iStart As Integer, _
                 sSeps As String) As Integer
    
    Dim cTarget As Integer
    cTarget = Len(sTarget)
    ' Look for start of token (character that isn't a separator)
    Do While InStr(sSeps, Mid$(sTarget, iStart, 1))
        If iStart > cTarget Then
            StrSpan = 0
            Exit Function
        Else
            iStart = iStart + 1
        End If
    Loop
    StrSpan = iStart

End Function

Public Function GetFileBaseExt(sFile As String) As String
    Dim iBase As Long, s As String
    If sFile = sEmpty Then Exit Function
    s = GetFullPath(sFile, iBase)
    GetFileBaseExt = Mid$(s, iBase)
End Function
Public Function GetFullPath(sFileName As String, _
                     Optional FilePart As Long, _
                     Optional ExtPart As Long, _
                     Optional DirPart As Long) As String

    Dim c As Long, p As Long, sRet As String
    If sFileName = sEmpty Then Exit Function
    
    ' Get the path size, then create string of that size
    sRet = String(cMaxPath, 0)
    c = GetFullPathName(sFileName, cMaxPath, sRet, p)
    If c = 0 Then ApiRaise Err.LastDllError
    Debug.Assert c <= cMaxPath
    sRet = Left$(sRet, c)

    ' Get the directory, file, and extension positions
    GetDirExt sRet, FilePart, DirPart, ExtPart
    GetFullPath = sRet
    
End Function

Private Sub GetDirExt(sFull As String, iFilePart As Long, _
                      iDirPart As Long, iExtPart As Long)

    Dim iDrv As Long, i As Long, cMax As Long
    cMax = Len(sFull)

    iDrv = Asc(UCase$(Left$(sFull, 1)))

    ' If in format d:\path\name.ext, return 3
    If iDrv <= 90 Then                          ' Less than Z
        If iDrv >= 65 Then                      ' Greater than A
            If Mid$(sFull, 2, 1) = ":" Then     ' Second character is :
                If Mid$(sFull, 3, 1) = "\" Then ' Third character is \
                    iDirPart = 3
                End If
            End If
        End If
    Else

        ' If in format \\machine\share\path\name.ext, return position of \path
        ' First and second character must be \
        If Mid$(sFull, 1, 2) <> "\\" Then ApiRaise ERROR_BAD_PATHNAME

        Dim fFirst As Boolean
        i = 3
        Do
            If Mid$(sFull, i, 1) = "\" Then
                If Not fFirst Then
                    fFirst = True
                Else
                    iDirPart = i
                    Exit Do
                End If
            End If
            i = i + 1
        Loop Until i = cMax
    End If

    ' Start from end and find extension
    iExtPart = cMax + 1       ' Assume no extension
    fFirst = False
    Dim sChar As String
    For i = cMax To iDirPart Step -1
        sChar = Mid$(sFull, i, 1)
        If Not fFirst Then
            If sChar = "." Then
                iExtPart = i
                fFirst = True
            End If
        End If
        If sChar = "\" Then
            iFilePart = i + 1
            Exit For
        End If
    Next
    Exit Sub
FailGetDirExt:
    iFilePart = 0
    iDirPart = 0
    iExtPart = 0
End Sub

Sub ApiRaise(ByVal e As Long)
    Err.Raise vbObjectError + 29000 + e, _
              App.EXEName & ".Windows", ApiError(e)
End Sub

Function ApiError(ByVal e As Long) As String
    Dim s As String, c As Long
    s = String(256, 0)
    c = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                      FORMAT_MESSAGE_IGNORE_INSERTS, _
                      pNull, e, 0&, s, Len(s), ByVal pNull)
    If c Then ApiError = Left$(s, c)
End Function

' Work around limitation of AddressOf
'    Call like this: procVar = GetProc(AddressOf ProcName)
Function GetProc(proc As Long) As Long
    GetProc = proc
End Function
Function StringToPointer(s As String) As Long
    If UnicodeTypeLib Then
        StringToPointer = VarPtr(s)
    Else
        StringToPointer = StrPtr(s)
    End If
End Function
' Make sure path ends in a backslash
Function NormalizePath(sPath As String) As String
    If Right$(sPath, 1) <> sBSlash Then
        NormalizePath = sPath & sBSlash
    Else
        NormalizePath = sPath
    End If
End Function


'***End HardCore VB code********************************************


'following code plucked from "FormShaper" by Mel Grubb II
'Resource region file also made with FormShaper
'I modified it... added the region handle argument
'***Begin*****************************************
Public Sub RegionFromResource(m_lngRegion As Long, ResID As Integer, ResType As String)
    Dim abytRegion() As Byte

    ' Pull the region data from the resource
    abytRegion = LoadResData(ResID, ResType)
    If UBound(abytRegion) > 0 Then
        If m_lngRegion <> 0 Then DeleteObject m_lngRegion
        m_lngRegion = ExtCreateRegion(ByVal 0&, UBound(abytRegion) + 1, abytRegion(0))
    End If
End Sub
Public Sub Apply(ByVal hWnd As Long, ByVal m_lngRegion As Long)
    SetWindowRgn hWnd, m_lngRegion, True
End Sub

'***End FormShaper code********************************************

'********Begin*****************************
    
Public Sub GetAndSaveSnapShot()  'sub renamed
    'KPD-Team 2000
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    '-> Compile this code for better performance
    Dim bi24BitInfo As BITMAPINFO, bBytes() As Byte, Cnt As Long
    Dim iDC As Long, iBitmap As Long, iHPicDC As Long, iHBitmapOld As Long
    
    With bi24BitInfo.bmiHeader
        .biBitCount = 24
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bi24BitInfo.bmiHeader)
        .biWidth = (frmMain.ScaleWidth) \ Screen.TwipsPerPixelX  '- frmMain.tbLeft.Width
        .biHeight = frmMain.ScaleHeight \ Screen.TwipsPerPixelY + GetSystemMetrics(SM_CYCAPTION)
        .biSizeImage = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight
    End With
    ReDim bBytes(1 To bi24BitInfo.bmiHeader.biWidth * _
        bi24BitInfo.bmiHeader.biHeight * 3) As Byte
    'iDC = CreateCompatibleDC(0)
    frmMain.Picture = frmMain.Image  'added... speeds up the procedure greatly
    iDC = CreateCompatibleDC(frmMain.hdc)
    iBitmap = CreateDIBSection(iDC, bi24BitInfo, _
        DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    iHBitmapOld = SelectObject(iDC, iBitmap)
    iHPicDC = GetWindowDC(frmMain.hWnd)
    'StretchBlt iDC, 0, 0, bi24BitInfo.bmiHeader.biWidth, _
        bi24BitInfo.bmiHeader.biHeight, GetDC(0), 0, 0, _
            bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, vbSrcCopy
    StretchBlt iDC, 0, 0, bi24BitInfo.bmiHeader.biWidth, _
        bi24BitInfo.bmiHeader.biHeight, iHPicDC, 0, 0, _
            bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, vbSrcCopy
    GetDIBits iDC, iBitmap, 0, bi24BitInfo.bmiHeader.biHeight, bBytes(1), _
        bi24BitInfo, DIB_RGB_COLORS
    
    'not using pb for saving...
    'SetDIBitsToDevice frmMain.picSnap.hdc, 0, 0, _
        bi24BitInfo.bmiHeader.biWidth, _
        bi24BitInfo.bmiHeader.biHeight, 0, 0, 0, _
        bi24BitInfo.bmiHeader.biHeight, bBytes(1), bi24BitInfo, DIB_RGB_COLORS
    Call SelectObject(iDC, iHBitmapOld)
    DeleteDC iDC
    DeleteObject iBitmap
    Call ReleaseDC(frmMain.hWnd, iHPicDC)
    '***** Added code
    Call SaveBmp2File(bi24BitInfo, bBytes())
End Sub
'**************End *************************

'************Begin code*****************
Public Function RotateText(inObj As Object, x As Single, y As Single, inText As String, _
        Optional inFontName As String = "Arial", Optional inBold As Boolean = False, _
        Optional inItalic As Boolean = False, Optional inFontSize As Integer = 12, _
        Optional iAngle As Long = 0, Optional iOriention As Long = 0, _
        Optional iColor As Long = vbBlack, Optional iROP As Long = vbCopyPen, _
        Optional fDoIndividualChars As Boolean = False, _
        Optional rDoIndDelay As Single = 0) As Boolean
    On Error GoTo errHandler
' TextEffect.frm
'
' By Herman Liu

    'I deleted code that wasn't needed, modified parts of it and added others
    RotateText = False

    Dim L As LOGFONT, HFont As Long, oldROP As Long, wTextParams As DRAWTEXTPARAMS
    Dim mPrevFont As Long, i As Integer, origMode As Integer, sFontName As String
    Dim tmpX As Single, tmpY As Single, iTC As Long, iBG As Long
    Dim mresult As Long, ReturnSL As SIZEL, RC As RECT

     ' For Windows NT to work
    mresult = SetGraphicsMode(inObj.hdc, GM_ADVANCED)

    origMode = inObj.ScaleMode
    inObj.ScaleMode = vbPixels

    If inBold = True And inItalic = True Then
        'L.lfFaceName = inFontName & Space(1) & "Bold" & Space(1) & "Italic" & Chr(0)    ' Must be null terminated
        sFontName$ = inFontName$ & Space(1) & "Bold" & Space(1) & "Italic" & Chr(0)
    ElseIf inBold = True And inItalic = False Then
        'L.lfFaceName = inFontName & Space(1) & "Bold" + Chr$(0)
        sFontName$ = inFontName$ & Space(1) & "Bold" + Chr$(0)
    ElseIf inBold = False And inItalic = True Then
        'L.lfFaceName = inFontName & Space(1) & "Italic" + Chr$(0)
        sFontName$ = inFontName$ & Space(1) & "Italic" + Chr$(0)
    Else
        'L.lfFaceName = inFontName & Chr$(0)
        sFontName$ = inFontName$ & Chr$(0)
    End If
    'Call StrToBytes(L.lfFaceName, sFontName$)
    L.lfFaceName = sFontName$
    
    L.lfOutPrecision = OUT_TT_PRECIS
    L.lfQuality = ANTIALIASED_QUALITY   ' PROOF_QUALITY
    L.lfOrientation = iOriention * 10
    L.lfEscapement = iAngle * 10
    L.lfHeight = inFontSize * -20 / Screen.TwipsPerPixelY

    HFont = CreateFontIndirect(L)
    mPrevFont = SelectObject(inObj.hdc, HFont)
    
    oldROP = SetROP2(inObj.hdc, iROP)
    iTC = SetTextColor(inObj.hdc, iColor)
    iBG = SetBkMode(inObj.hdc, TRANSPARENT)
    wTextParams.cbSize = Len(wTextParams)
    
    'added the following to draw individual chars of a string... js
    'couldn't get the lfEscapement to work with single chars so
    'improvised with some trig and the string width
    If fDoIndividualChars Then
        Dim s As String, sL As String, iCharWidth As Long, iWTotal As Long, rOffset As Single
        For i = 1 To Len(inText)
            s$ = Mid$(inText, i, 1) 'current char to draw
            sL$ = Left$(inText, i)  'string up to the current char-need for rect
            Call GetCharWidth32(inObj.hdc, Asc(s$), Asc(s$), iCharWidth) 'current char width
            iWTotal = iWTotal + iCharWidth  'width total
            rOffset = Tan((PI / 180) * -iAngle) * iWTotal  'char y offset
'Debug.Print s$; " "; iCharWidth; "  "; rOffset
            Call GetTextExtentPoint32(inObj.hdc, sL$, Len(sL$), ReturnSL) 'total extent to current char
            With RC
                .Left = x
                .Top = y + rOffset  ''(i * 1) \ 2 '-iAngle ' 2
                .Right = x + ReturnSL.cx
                .Bottom = y + ReturnSL.cy + rOffset  ''(i * 1) \ 2  '-iAngle ' 2
            End With
            Call DrawTextEx(inObj.hdc, s$, Len(s$), RC, DT_RIGHT Or DT_NOCLIP Or DT_WORDBREAK, wTextParams)   'Or DT_VCENTER  'DT_RIGHT
            If rDoIndDelay <> 0 Then 'add a delay to simulate actual typing
                inObj.Refresh
                Delay rDoIndDelay
            End If
        Next
        'inObj.Refresh
    Else
        Call GetTextExtentPoint32(inObj.hdc, inText, Len(inText), ReturnSL)
        With RC
            .Left = x
            .Top = y
            .Right = x + ReturnSL.cx
            .Bottom = y + ReturnSL.cy
        End With
        'wTextParams.cbSize = Len(wTextParams)
        'TextHeightRet = ReturnSL.cy
        'TextWidthRet = ReturnSL.cx
        Call DrawTextEx(inObj.hdc, inText, Len(inText), RC, DT_CENTER Or DT_NOCLIP Or DT_VCENTER Or DT_WORDBREAK, wTextParams)
    End If
    
    mresult = SelectObject(inObj.hdc, mPrevFont)
    mresult = DeleteObject(HFont)
    inObj.ScaleMode = origMode
    Call SetTextColor(inObj.hdc, iTC)
    Call SetBkMode(inObj.hdc, iBG)
    Call SetROP2(inObj.hdc, oldROP)
    RotateText = True
    Exit Function

errHandler:
    inObj.ScaleMode = origMode
    MsgBox "RotateText Function Error"
End Function
'********************end code********************

'********************Begin Code*******************
'I have modified parts of the following code...js
'Removed proxy stuff, added undeclared variables, converted to function,
'removed the async and callback to textbox mouseUp event...
'save to a temp string variable and return it with the fuction


'********************
'Modifications and improvements by Luis Cantero (2002)
'Modifications: ListenForConnect, Ping, GetMXName, GetDNSInfo, MyIP, SendData, etc.
'http://LCenterprises.net
'********************

'Visual Basic 6.0 Winsock "Header"
'   Alot of the information contained inside this file was originally
'   obtained from ALT.WINSOCK.PROGRAMMING and most of it has since been
'   modified in some way.
'
'Disclaimer: This file is public domain, updated periodically by
'   Topaz, SigSegV@mail.utexas.edu, Use it at your own risk.
'   Neither myself(Topaz) or anyone related to alt.programming.winsock
'   may be held liable for its use, or misuse.
'
'Declare check Aug 27, 1996. (Topaz, SigSegV@mail.utexas.edu)
'   All 16 bit declarations appear correct, even the odd ones that
'   pass longs inplace of in_addr and char buffers. 32 bit functions
'   also appear correct. Some are declared to return integers instead of
'   longs (breaking MS's rules.) however after testing these functions I
'   have come to the conclusion that they do not work properly when declared
'   following MS's rules.
Public Function GetFromInet(strURL As String) As String

  Dim SocketBuffer As SOCKADDR, strPath As String, strHost As String, intPort As Long
  Dim IpAddr As Long, iSlashPos As Long, RC As Long, i As Long
  Dim tmpHost As String, strMsg As String, sStart As String
    
    'Remove leading http or https
    If StrComp(Left$(strURL$, 4), "http", vbTextCompare) = 0 Then
        iSlashPos = InStr(5, strURL$, "/")
        strURL$ = Mid$(strURL$, iSlashPos + 2)
    End If
    
    'Separate URL into Host and Path
    iSlashPos = InStr(1, strURL, "/")
    If iSlashPos = 0 Then iSlashPos = Len(strURL) + 1
    strPath = Mid$(strURL, iSlashPos)
    If strPath = "" Then strPath = "/"
    strHost = Mid$(strURL, 1, iSlashPos - 1)
    intPort = 80

    'sStart winsock
    Call StartWinsock

    'Create socket
    Sock = Socket(AF_INET, SOCK_STREAM, 0)
    If Sock = SOCKET_ERROR Then frmDownLoad.lblStatus.Caption = "SOCKET_ERROR: CreateSocket": Exit Function

    If RC = SOCKET_ERROR Then frmDownLoad.lblStatus.Caption = "SOCKET_ERROR CreateSocket-rc":  Exit Function
    IpAddr = GetHostByNameAlias(strHost)
    If IpAddr = -1 Then
        frmDownLoad.lblStatus.Caption = "Unknown host"
        Exit Function
    End If
    
    With SocketBuffer
        .sin_family = AF_INET
        .sin_port = htons(intPort)
        .sin_addr = IpAddr
        '.sin_zero = String$(8, 0)
        For i = 0 To 7
            .sin_zero(i) = 0  'String$(8, 0)
        Next
    End With
    
    frmDownLoad.lblStatus.Caption = "Connecting to " & strHost
    DoEvents
    
    'Connect to server
    RC = Connect(Sock, SocketBuffer, Len(SocketBuffer))
    
    If RC = SOCKET_ERROR Then
        CloseSocket Sock
        Call EndWinsock
        frmDownLoad.lblStatus.Caption = "Could not connect to " & strHost
        Exit Function
      Else
    End If
    
    frmDownLoad.lblStatus.Caption = "Connected to " & strHost
    DoEvents

'I'm not using a textbox to store the text in so....
''    'Set receive window
''    RC = WSAAsyncSelect(Sock, frmDownLoad.txtReceive.hWnd, _
''         ByVal &H202, ByVal FD_READ Or FD_CLOSE)
''    If RC = SOCKET_ERROR Then
''        CloseSocket Sock
''        Call EndWinsock
''        frmDownLoad.lblStatus.Caption = "SOCKET_ERROR: SetReceiveWindow"
''        Exit Sub
''    End If
    
    'Prepare GET header
    'When to use GET? -> When the amount of data that you
    'need to pass to the server is not much
    strMsg = "GET " & tmpHost & strPath & " HTTP/1.0" & vbCrLf
    strMsg = strMsg & "Accept: */*" & vbCrLf
    strMsg = strMsg & "User-Agent: " & App.Title & vbCrLf
    strMsg = strMsg & "Host: " & strHost & vbCrLf
    strMsg = strMsg & vbCrLf
    
    'Example POST header
    'When to use POST? -> Anytime, it is failsafe since the
    'content-length is passed to the server everytime
    'strCommand = "Temp1=hello&temp2=12345&Etc=hallo"
    'strMsg = "POST " & tmpHost & strPath & " HTTP/1.0" & vbCrLf
    'strMsg = strMsg & "Accept: */*" & vbCrLf
    'strMsg = strMsg & "User-Agent: " & App.Title & vbCrLf
    'strMsg = strMsg & "Host: " & strHost & vbCrLf
    'strMsg = strMsg & "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    'strMsg = strMsg & "Content-Length: " & Len(strCommand) & vbCrLf
    'strMsg = strMsg & vbCrLf & strCommand
    
    frmDownLoad.lblStatus.Caption = "Sending request..."
    DoEvents
    
    'Send request
    SendData Sock, strMsg
    
    If tmpHost = "" Then tmpHost = strHost
    
    'Wait for page to be downloaded
    'Seconds to wait = 10
    sStart = (Format$(Now, "NN") * 60 + Format$(Now, "SS")) + 10
    While Not sStart <= (Format$(Now, "NN") * 60 + Format$(Now, "SS")) And Sock > 0
        frmDownLoad.lblStatus.Caption = "Waiting for response from " & tmpHost '& "..." & sStart - (Format$(Now, "NN") * 60 + Format$(Now, "SS"))
        DoEvents
        
        'You can put a routine that will check if a boolean variable is True here
        'This could indicate that the request has been canceled
        'If CancelFlag = True Then
        '   frmdownload.lblstatus.caption = "Cancelled request"
        '   Exit Sub
        'End If
    Wend
    
    frmDownLoad.lblStatus.Caption = "Ready"
'
'*****Here is where I add the code to save the data in a string variable
'instead of sending it to the textbox. Bits and pieces of this are from vbapi.com
        
    Dim MsgBuffer As String * 8192, sServerResponse As String, iBytes As Integer

    On Error Resume Next

    'A Socket is open
    If Sock > 0 Then
        Do
            DoEvents
            'Receive up to 8192 chars
            iBytes = recv(Sock, ByVal MsgBuffer, 8192, 0)
            If iBytes > 0 And iBytes <> SOCKET_ERROR Then
                sServerResponse$ = sServerResponse$ & Mid$(MsgBuffer, 1, iBytes)
            End If
        Loop Until iBytes = 0
        CloseSocket (Sock)
        Call EndWinsock 'Very important!
        Sock = 0
    End If
      
    GetFromInet$ = sServerResponse$

End Function

Public Sub EndWinsock()

  Dim ret&

    If WSAIsBlocking() Then
        ret = WSACancelBlockingCall()
    End If
    ret = WSACleanup()
    WSAStartedUp = False

End Sub

Private Function StartWinsock(Optional sDescription As String) As Boolean

  Dim StartupData As WSADATA
  Dim RC As Long

    If Not WSAStartedUp Then
        If Not WSAStartup(&H101, StartupData) Then
            RC = WSAStartup(&H101, StartupData)
            WSAStartedUp = True
'Debug.Print "wVersion="; StartupData.wVersion, "wHighVersion="; StartupData.wHighVersion
'Debug.Print "If wVersion = 257 then everything is kewl"
'Debug.Print "szDescription="; StartupData.szDescription
'Debug.Print "szSystemStatus="; StartupData.szSystemStatus
'Debug.Print "iMaxSockets="; StartupData.iMaxSockets, "iMaxUdpDg="; StartupData.iMaxUdpDg
            sDescription = StartupData.szDescription
          Else
            WSAStartedUp = False
        End If
      Else
        Call EndWinsock
        Call StartWinsock
    End If
    StartWinsock = WSAStartedUp

End Function
'returns IP as long, in network byte order
Private Function GetHostByNameAlias(ByVal hostname$) As Long

  Dim phe&
  Dim heDestHost As HOSTENT
  Dim addrList&
  Dim retIP&

    retIP = inet_addr(hostname$)
    If retIP = INADDR_NONE Then
        phe = gethostbyname(hostname$)
        If phe <> 0 Then
            CopyMemory heDestHost, ByVal phe, hostent_size
            CopyMemory addrList, ByVal heDestHost.h_addr_list, 4
            CopyMemory retIP, ByVal addrList, heDestHost.h_length
          Else
            retIP = INADDR_NONE
        End If
    End If
    GetHostByNameAlias = retIP

End Function

Private Function SendData(ByVal intSocket&, vMessage As Variant) As Long

  Dim TheMsg() As Byte, sTemp$

    TheMsg = ""
    Select Case VarType(vMessage)
      Case 8209   'byte array
        sTemp = vMessage
        TheMsg = sTemp
      Case 8      'string, if we receive a string, its assumed we are linemode
        sTemp = StrConv(vMessage, vbFromUnicode)
      Case Else
        sTemp = CStr(vMessage)
        sTemp = StrConv(vMessage, vbFromUnicode)
    End Select
    
    TheMsg = sTemp
    
    If UBound(TheMsg) > -1 Then
        SendData = send(intSocket, TheMsg(0), UBound(TheMsg) + 1, 0)
    End If
    
    If SendData = SOCKET_ERROR Then
        CloseSocket intSocket
        Call EndWinsock
        Exit Function
    End If

End Function
Public Function IsConnected() As Boolean
    'this function will not determine between a inet conn. or LAN...js
    On Error GoTo Err
    IsConnected = InternetGetConnectedState(0&, 0&)
Exit Function
Err:
    IsConnected = True
End Function
'******************End Winsock code*******************
'*****************begin Conn Code**************
'Tip by John Percival From VB - World
Public Function Online() As Boolean
    'If you are online it will return True, otherwise False
    Online = InternetGetConnectedState(0&, 0&)
End Function

Public Function ViaLAN() As Boolean
    Dim SFlags As Long
    'return the flags associated with the connection
    Call InternetGetConnectedState(SFlags, 0&)
    'True if the Sflags has a LAN connection
    ViaLAN = SFlags And INTERNET_CONNECTION_LAN
End Function
Public Function ViaModem() As Boolean
    Dim SFlags As Long
    'return the flags associated with the connection
    Call InternetGetConnectedState(SFlags, 0&)
    'True if the Sflags has a modem connection
    ViaModem = SFlags And INTERNET_CONNECTION_MODEM
End Function

'*****************end conn code***************************

''******************Begin Oleg's code*****************************
'Public Function vbRecv(ByVal lngSocket As Long, strBuffer As String) As Long
''********************************************************************************
''Author    :Oleg Gdalevich
''Date/Time :27-Nov-2001
''Purpose   :Retrieves data from the Winsock buffer.
''Returns   :Number of bytes received.
''Arguments :lngSocket    - the socket connected to the remote host
''           strBuffer    - buffer to read data to
''********************************************************************************
'    '
'    Const MAX_BUFFER_LENGTH As Long = 8192
'    '
'    Dim arrBuffer(1 To MAX_BUFFER_LENGTH)   As Byte
'    Dim lngBytesReceived                    As Long
'    Dim strTempBuffer                       As String
'    '
'    'Check the socket for readabilty with
'    'the IsDataAvailable function
'    If IsDataAvailable(lngSocket) Then
'        '
'        'Call the recv Winsock API function in order to read data from the buffer
'        lngBytesReceived = Recv(lngSocket, arrBuffer(1), MAX_BUFFER_LENGTH, 0&)
'        '
'        If lngBytesReceived > 0 Then
'            '
'            'If we have received some data, convert it to the Unicode
'            'string that is suitable for the Visual Basic String data type
'            strTempBuffer = StrConv(arrBuffer, vbUnicode)
'            '
'            'Remove unused bytes
'            strBuffer = Left$(strTempBuffer, lngBytesReceived)
'            '
'        End If
'        '
'        'If lngBytesReceived is equal to 0 or -1, we have nothing to do with that
'        '
'        vbRecv = lngBytesReceived
'        '
'    Else
'        '
'        'Something wrong with the socket.
'        'Maybe the lngSocket argument is not a valid socket handle at all
'        vbRecv = SOCKET_ERROR
'        '
'    End If
'    '
'End Function
'Public Function IsDataAvailable(ByVal lngSocket As Long) As Boolean
'    '
'    Dim udtRead_fd As fd_set
'    Dim udtWrite_fd As fd_set
'    Dim udtError_fd As fd_set
'    Dim lngSocketCount As Long
'    '
'    udtRead_fd.fd_count = 1
'    udtRead_fd.fd_array(1) = lngSocket
'    '
'    lngSocketCount = vbselect(0&, udtRead_fd, udtWrite_fd, udtError_fd, 0&)
'    '
'    IsDataAvailable = CBool(lngSocketCount)
'    '
'End Function
'
'
''******************End Oleg's code*****************************

'*******************Begin *************************************
'I modified this one too......js
'=====================================================================================
' Browse for a Folder using SHBrowseForFolder API function with a callback
' function BrowseCallbackProc.
'
' Stephen Fonnesbeck
' steev@xmission.com
' http://www.xmission.com/~steev
' Feb 20, 2000
'=============================================================
Public Function BrowseForFolder(Optional ihWnd As Long = 0, _
                                Optional sTitle As String = "Select Folder", _
                                Optional sStartDir As String, _
                                Optional bAddDir As Boolean = False) As String
  'Opens a Treeview control that displays the directories in a computer

  Dim lpIDList As Long
  Dim szTitle As String
  Dim sBuffer As String
  Dim tBrowseInfo As BROWSEINFO
  
  If sStartDir$ = sEmpty Then  'added... js
    If m_CurrentDirectory = sEmpty Then
        m_CurrentDirectory = App.Path
    End If
  Else
    m_CurrentDirectory = sStartDir$ & vbNullChar
  End If
  
  szTitle = sTitle$
  With tBrowseInfo
    .hwndOwner = ihWnd
    .lpszTitle = lstrcat(szTitle, "")
    If bAddDir = True Then  'added...js  Note...isn't working....
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_EDITBOX Or BIF_STATUSTEXT
    Else
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_STATUSTEXT
    End If
    .lpfn = GetAddressofFunction(AddressOf BrowseCallbackProc)  'get address of function.
  End With
  
  CenterDlgBox ihWnd  'added...js
  
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  If (lpIDList) Then
    sBuffer$ = Space(cMaxPath)
    SHGetPathFromIDList lpIDList, sBuffer$
    sBuffer$ = Left(sBuffer$, InStr(sBuffer$, vbNullChar) - 1)
    BrowseForFolder = sBuffer$
    m_CurrentDirectory = sBuffer$
  Else
    BrowseForFolder = ""
  End If
  
End Function
 
Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
  
  Dim lpIDList As Long
  Dim ret As Long
  Dim sBuffer As String
  
  On Error Resume Next  'Sugested by MS to prevent an error from
                        'propagating back into the calling process.
     
  Select Case uMsg
  
    Case BFFM_INITIALIZED
      Call SendMessage(hWnd, BFFM_SETSELECTION, 1, m_CurrentDirectory)
      
    Case BFFM_SELCHANGED
      sBuffer = Space(cMaxPath)
      
      ret = SHGetPathFromIDList(lp, sBuffer)
      If ret = 1 Then
        Call SendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
      End If
      
  End Select
  
  BrowseCallbackProc = 0
  
End Function

' This function allows you to assign a function pointer to a vaiable.
Private Function GetAddressofFunction(add As Long) As Long
  GetAddressofFunction = add
End Function


'*******************End************************************************

'********************Begin code
'Microsoft Knowledge Base Article - Q189170
Function MakeDWord(LoWord As Integer, HiWord As Integer) As Long
   MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)
End Function

Function LoWord(DWord As Long) As Integer
   If DWord And &H8000& Then ' &H8000& = &H00008000
      LoWord = DWord Or &HFFFF0000
   Else
      LoWord = DWord And &HFFFF&
   End If
End Function

Function HiWord(DWord As Long) As Integer
   HiWord = (DWord And &HFFFF0000) \ &H10000
End Function
'********************end code
