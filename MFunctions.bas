Attribute VB_Name = "MFunctions"
Option Explicit



Private hHookCntrMsgBox As Long, lFrmHwndCntrMsgBox As Long

Public Function GetIni(File As String, Section As String, Key As String)
    Dim RetVal As Long
    Dim Value As String
    
    Value = Space(2001) ' Initialize Return String
    ' Query Value
    RetVal = GetPrivateProfileString(Section, Key, "", Value, 2000&, File)
    ' Trim Return String
    Value = Left(Value, RetVal)
    GetIni = Value
End Function

Public Sub WriteIni(File As String, Section As String, Key As String, Value As String)
    WritePrivateProfileString Section, Key, Value, File
End Sub
Public Function GetNumIniKeys(File As String, Section As String) As Integer
    Dim RetVal As Long, RetString As String, NullOffSet As Integer
    Dim Keys As String, Counter As Integer
    
    RetString = Space(2001) ' Initialize Return String
    'RetVal = GetPrivateProfileString(Section, 0&, "", RetString, 2000&, File)
    RetVal = GetPrivateProfileStringKeys(Section, 0&, "", RetString, 2000&, File)
    
    Do
    NullOffSet = InStr(RetString, Chr$(0))
    If NullOffSet > 1 Then
    Counter = Counter + 1 ' Increment Counter
    Keys = Keys & Chr$(10) & Mid$(RetString, 1, NullOffSet - 1)
    RetString = Mid$(RetString, NullOffSet + 1)
    End If
    Loop While NullOffSet > 1
    
    GetNumIniKeys = Counter
End Function

Public Function GetIniKey(File As String, Section As String, KeyNumber As Integer) As String
    Dim RetVal As Long, RetString As String, NullOffSet As Integer
    Dim Keys As String, Counter As Integer
    
    RetString = Space(2001) ' Initialize Return String
    'RetVal = GetPrivateProfileString(Section, 0&, "", RetString, 2000&, File)
    RetVal = GetPrivateProfileStringKeys(Section, 0&, "", RetString, 2000&, File)
    
    Do
    NullOffSet = InStr(RetString, Chr$(0))
    If NullOffSet > 1 Then
    Counter = Counter + 1 ' Increment Counter
    If Counter = KeyNumber Then ' Key Found
    GetIniKey = Mid$(RetString, 1, NullOffSet - 1)
    Exit Do
    End If
    Keys = Keys & Chr$(10) & Mid$(RetString, 1, NullOffSet - 1)
    RetString = Mid$(RetString, NullOffSet + 1)
    End If
    Loop While NullOffSet > 1
End Function
Public Sub GetIniSettings()
    On Error Resume Next
    sSymbol$ = GetIni(sINIsetFile, "DataInfo", "Symbol")
    sFilePath$ = GetIni(sINIsetFile, "Settings", "LastFile")
    sFileName$ = GetFileBaseExt(sFilePath$)
    sDataDir$ = GetIni(sINIsetFile, "DLSettings", "DataDir")
    iBackColor = Val(GetIni(sINIsetFile, "Settings", "BackColor"))
    iForeColor = Val(GetIni(sINIsetFile, "Settings", "ForeColor"))
    iGridColor = Val(GetIni(sINIsetFile, "Settings", "GridColor"))
    iMouseLabelColor = Val(GetIni(sINIsetFile, "Settings", "MouseLabelColor"))
    iTicBodyColor = Val(GetIni(sINIsetFile, "Settings", "TicColor"))
    iTicOpenColor = Val(GetIni(sINIsetFile, "Settings", "OpenColor"))
    iTicCloseColor = Val(GetIni(sINIsetFile, "Settings", "CloseColor"))
    iTicType = Val(GetIni(sINIsetFile, "Settings", "TicType"))
    iTicCandleUpColor = Val(GetIni(sINIsetFile, "Settings", "TicCandleUpColor"))
    iTicCandleDnColor = Val(GetIni(sINIsetFile, "Settings", "TicCandleDnColor"))
    iVolColor = Val(GetIni(sINIsetFile, "Settings", "VolColor"))
    sFontName = GetIni(sINIsetFile, "Settings", "FontName")
    iFontSize = Val(GetIni(sINIsetFile, "Settings", "FontSize"))
    iFontBold = Val(GetIni(sINIsetFile, "Settings", "FontBold"))
    iFontItalic = Val(GetIni(sINIsetFile, "Settings", "FontItalic"))
    iDateMarkerColor = Val(GetIni(sINIsetFile, "Settings", "DateMarkerColor"))
    iCrossHairMode = Val(GetIni(sINIsetFile, "Settings", "CrossHairMode"))
    iCrossHairColor = Val(GetIni(sINIsetFile, "Settings", "CrossHairColor"))
    iBlankSpace = Val(GetIni(sINIsetFile, "Settings", "BlankSpace"))
    iBarSpacing = Val(GetIni(sINIsetFile, "Settings", "BarSpacing"))
    iScrollIncrement = Val(GetIni(sINIsetFile, "Settings", "ScrollIncrement"))
    rSplit1 = Val(GetIni(sINIsetFile, "Settings", "WindowSplit1"))
    rSplit2 = Val(GetIni(sINIsetFile, "Settings", "WindowSplit2"))
    CurrentIndicator = GetIni(sINIsetFile, "Settings", "CurrentIndicator")
    

End Sub

Public Sub SaveIniSettings()

    WriteIni sINIsetFile, "DataInfo", "Symbol", UCase$(sSymbol$)
    WriteIni sINIsetFile, "Settings", "BackColor", CStr(iBackColor)
    WriteIni sINIsetFile, "Settings", "ForeColor", CStr(iForeColor)
    WriteIni sINIsetFile, "Settings", "GridColor", CStr(iGridColor)
    WriteIni sINIsetFile, "Settings", "MouseLabelColor", CStr(iMouseLabelColor)
    WriteIni sINIsetFile, "Settings", "TicColor", CStr(iTicBodyColor)
    WriteIni sINIsetFile, "Settings", "OpenColor", CStr(iTicOpenColor)
    WriteIni sINIsetFile, "Settings", "CloseColor", CStr(iTicCloseColor)
    WriteIni sINIsetFile, "Settings", "TicType", CStr(iTicType)
    WriteIni sINIsetFile, "Settings", "TicCandleUpColor", CStr(iTicCandleUpColor)
    WriteIni sINIsetFile, "Settings", "TicCandleDnColor", CStr(iTicCandleDnColor)
    WriteIni sINIsetFile, "Settings", "VolColor", CStr(iVolColor)
    WriteIni sINIsetFile, "Settings", "FontName", sFontName
    WriteIni sINIsetFile, "Settings", "FontSize", CStr(iFontSize)
    WriteIni sINIsetFile, "Settings", "FontBold", CStr(iFontBold)
    WriteIni sINIsetFile, "Settings", "FontItalic", CStr(iFontItalic)
    WriteIni sINIsetFile, "Settings", "DateMarkerColor", CStr(iDateMarkerColor)
    WriteIni sINIsetFile, "Settings", "CrossHairColor", CStr(iCrossHairColor)
    WriteIni sINIsetFile, "Settings", "CrossHairMode", CStr(iCrossHairMode)
    WriteIni sINIsetFile, "Settings", "BlankSpace", CStr(iBlankSpace)
    WriteIni sINIsetFile, "Settings", "BarSpacing", CStr(iBarSpacing)
    WriteIni sINIsetFile, "Settings", "ScrollIncrement", CStr(iScrollIncrement)
    WriteIni sINIsetFile, "Settings", "WindowSplit1", CStr(rSplit1)
    WriteIni sINIsetFile, "Settings", "WindowSplit2", CStr(rSplit2)
    WriteIni sINIsetFile, "Settings", "CurrentIndicator", CurrentIndicator

End Sub
Public Sub MakeIniFile()
    'use for reseting default values also
    Open sINIsetFile$ For Output Access Write As #1
        Print #1, "[DataInfo]"
        Print #1, "Symbol="
        Print #1, sEmpty
        
        Print #1, "[Settings]"
        Print #1, "LastFile="
        Print #1, "BackColor=0"
        Print #1, "ForeColor=16777215"
        Print #1, "GridColor=10066329"
        Print #1, "MouseLabelColor=65535"
        Print #1, "TicColor=65408"
        Print #1, "OpenColor=65280"
        Print #1, "CloseColor=16777215"
        Print #1, "TicType=3"
        Print #1, "TicCandleUpColor=65280"
        Print #1, "TicCandleDnColor=255"
        Print #1, "VolColor=65280"
        Print #1, "FontName=MS Sans Serif"
        Print #1, "FontSize=10"
        Print #1, "FontBold=-1"
        Print #1, "FontItalic=0"
        Print #1, "DateMarkerColor=255"
        Print #1, "CrossHairMode=15"
        Print #1, "CrossHairColor=16711680"
        Print #1, "BlankSpace=10"
        Print #1, "BarSpacing=3"
        Print #1, "ScrollIncrement=10"
        Print #1, "WindowSplit1=430"
        Print #1, "WindowSplit2=480"
        Print #1, "CurrentIndicator=MACD"
        Print #1, sEmpty
        
        Print #1, "[DLSettings]"
        Print #1, "DataDir=\Data"
        Print #1, "LastURL="
        Print #1, "Source=0"
        Print #1, sEmpty
        
        Print #1, "[Indicators]"
        Print #1, "1=CCI"
        Print #1, "2=MACD"
        Print #1, "3=ROC%"
        Print #1, "4=RSI"
        Print #1, "5=STOCH"
        Print #1, sEmpty
        
        Print #1, "[CciSettings]"
        Print #1, "CciColor=15655586"
        Print #1, "CciAvgColor=13436626"
        Print #1, "CciLen=13"
        Print #1, "CciAvgLen=3"
        Print #1, "CciAvgAvgLen=5"
        Print #1, "PlotWidth=1"
        Print #1, "CciUpperTrig=200"
        Print #1, "CciLowerTrig=-150"
        Print #1, sEmpty
        
        Print #1, "[MACDSettings]"
        Print #1, "MACDColor1=255"
        Print #1, "MACDColor2=16711680"
        Print #1, "MACDColor3=12100980"
        Print #1, "PlotWidth=2"
        Print #1, "MACDLen1=13"
        Print #1, "MACDLen2=34"
        Print #1, "MACDLen3=9"
        Print #1, sEmpty
        
        Print #1, "[ROC%Settings]"
        Print #1, "ROC%Color=255"
        Print #1, "PlotWidth=1"
        Print #1, "ROC%Len=12"
        Print #1, "ROC%UpperTrig=6.5"
        Print #1, "ROC%LowerTrig=-6.5"
        Print #1, sEmpty
        
        Print #1, "[RsiSettings]"
        Print #1, "RsiColor=255"
        Print #1, "PlotWidth=1"
        Print #1, "RsiLen=14"
        Print #1, "RsiUpperTrig=70"
        Print #1, "RsiLowerTrig=30"
        Print #1, sEmpty
        
        Print #1, "[StochSettings]"
        Print #1, "StochKColor=255"
        Print #1, "StochKSlowColor=16711680"
        Print #1, "StochDColor=12100980"
        Print #1, "PlotWidth=1"
        Print #1, "KPeriod=14"
        Print #1, "KPeriodSlow=4"
        Print #1, "DPeriod=3"
        Print #1, "StochUpperTrig=70"
        Print #1, "StochLowerTrig=30"
        Print #1, "StochSlow=true"
        Print #1, sEmpty
        
        Print #1, "[AvgSettings]"
        Print #1, "AvgLen1=21"
        Print #1, "AvgColor1=255"
        Print #1, "AvgWidth1=1"
        Print #1, "AvgLen2=50"
        Print #1, "AvgColor2=16711680"
        Print #1, "AvgWidth2=1"
        Print #1, "AvgLen3=200"
        Print #1, "AvgColor3=65535"
        Print #1, "AvgWidth3=1"
        Print #1, sEmpty
    Close #1
End Sub
Public Function GetColorDlg(iPrevColor As Long) As Long
    Dim f As Boolean, iColor As Long
    
    iColor = iPrevColor
    CenterDlgBox 0
    f = VBChooseColor(Color:=iColor, _
            FullOpen:=True, _
            owner:=0)
    
    If f Then
        GetColorDlg = iColor
    Else
        GetColorDlg = iPrevColor
    End If

End Function
Public Function OpenDataFile() As Boolean
    Dim f As Boolean, sFile As String
    
    CenterDlgBox 0
    f = VBGetOpenFileName( _
            FileName:=sFile$, _
            ReadOnly:=False, _
            filter:="Data Files (*.dat): *.dat|All files (*.*): *.*", _
            DefaultExt:="*.dat", _
            FilterIndex:=1, _
            DlgTitle:="Open Data File", _
            owner:=0, InitDir:=sDataDir$)
    If f And sFile$ <> sEmpty Then
        sFilePath$ = sFile$
        WriteIni sINIsetFile, "Settings", "LastFile", sFilePath$
        sFileName$ = GetFileBaseExt(sFile$)
        Dim p As Long
        p = InStr(sFileName$, "~")  'check for symbol in file name
        If p <> 0 Then
            sSymbol$ = Left$(sFileName$, p - 1)
        Else  'not found...
            sSymbol$ = sUnknownSymbol$
        End If
        WriteIni sINIsetFile, "DataInfo", "Symbol", sSymbol$
    End If
    OpenDataFile = f
End Function
Public Function LoadData() As Boolean
    
    Dim x As Integer, i As Integer, y As Integer, c As Integer, ff As Integer, fSkipLine As Boolean
    Dim sLineFromFile As String, stoken As String, sTemp As String, iType As Integer
    
    If IsDrawing = True Then Exit Function  'if we're drawing a chart exit this function
            
    If Not ExistFile(sFilePath$) Then
        If OpenDataFile = False Then 'cancelled
            Exit Function
        Else 'new file
            
        End If
    End If
    
    If Not frmSplash.Visible Then frmSplash.Show 0, frmMain
    
    ff = FreeFile
    Open sFilePath$ For Input Access Read As ff
    
    Do While Not EOF(ff)
        DoEvents
            Line Input #ff, sLineFromFile$
            If Len(sLineFromFile$) > 2 Then c = c + 1 'line count, make sure not a blank
            If c = 1 Then
                'check the first line for data config
                Select Case sLineFromFile$
                    Case """Date"",""O"",""H"",""L"",""C"",""V"""
                        iType = 1   'typical end of day format
                    Case """Date"",""Time"",""O"",""H"",""L"",""C"",""V"""
                        iType = 2  'Typical intraday format
                    Case """Date"",""Time"",""O"",""H"",""L"",""C"",""U"",""D"""
                        iType = 3  'Omega format
                    Case "Date,Open,High,Low,Close,Volume"
                        iType = 1 'Yahoo EOD format
                    
                End Select
            End If
    Loop
    Close ff
'Debug.Print "c: "; c
    iUBaData = c - 1
    If iType <> 0 Then
        iUBaData = iUBaData - 1  'subtract first line from total
        fSkipLine = True  'set flag to skip the first line
    End If
    ReDim aData(0 To iUBaData)
    
    'parse the data
    Open sFilePath$ For Input Access Read As ff
    Do While Not EOF(ff)
        DoEvents
        Line Input #ff, sLineFromFile$
        If Not fSkipLine And Len(sLineFromFile$) > 2 Then
        
            stoken$ = GetQToken(sLineFromFile$, ",")
            Do While stoken$ <> sEmpty$
'Debug.Print stoken
                Select Case y
                    Case 0  'Date
'Debug.Print stoken
                        aData(x).sDate = stoken$
                        If iType = 1 Then  'no time in this config so we need to bump y +1
                            y = y + 1
                        End If
                    Case 1  'time
                        If Left(stoken$, 3) <> ":" Then _
                            sTemp$ = Left(stoken$, 2) & ":" & Right(stoken$, 2)
                        aData(x).sTime = sTemp$
                    Case 2  ' open
                        aData(x).dOpen = Round(Val(stoken$), 3)
                    Case 3  ' high
                        aData(x).dHigh = Round(Val(stoken$), 3)
                    Case 4  ' low
                        aData(x).dLow = Round(Val(stoken$), 3)
                    Case 5  ' close
                        aData(x).dClose = Round(Val(stoken$), 3)
                    Case 6  ' vol.
                        aData(x).iVol = Val(stoken$)
                    Case 7
                        'Omega data has the vol split into up & dn vol-> add it
                        If iType = 3 Then aData(x).iVol = aData(x).iVol + Val(stoken$)
                    Case Else
'Debug.Print "CaseElse"

                End Select
                y = y + 1
'Debug.Print "y: "; y
                stoken$ = GetQToken(sEmpty$, ",")
            Loop
            x = x + 1
        End If
        fSkipLine = False  'set flag so we can get input lines
        y = 0
    Loop
    Close ff
    
    Call CalculateDataPeriod
    LoadData = True
    
End Function
Private Sub CalculateDataPeriod()
    '*******************Calculate time between data entries
    Dim i1H As Integer, i2H As Integer, i1M As Integer, i2M As Integer
    Dim sTime As String, sTime2 As String, iDifH As Integer, iDifM As Integer
    
    sTime$ = aData(iUBaData).sTime
    sTime2$ = aData(iUBaData - 1).sTime
    'sTime$ = Trim$(Mid$(sTime$, InStr(sTime$, " ") + 1))
    'sTime2$ = Trim$(Mid$(sTime2$, InStr(sTime2$, " ") + 1))
'Debug.Print stime$
'Debug.Print stime2$
    If sTime$ = sTime2$ Then  'daily data
'Debug.Print DateDiff("d", aData(iUBaData - 1).sDate, aData(iUBaData).sDate)
        If DateDiff("d", aData(iUBaData - 1).sDate, aData(iUBaData).sDate) > 3 Then
            iBarDataPeriodMins = -2  'weekly or other
        Else
            iBarDataPeriodMins = -1  'daily
        End If
        Exit Sub
    End If
    i1H = Val(Left$(sTime$, InStr(sTime$, ":") - 1))
    i1M = Val(Mid$(sTime$, InStr(sTime$, ":") + 1))
'Debug.Print i1H; "  "; i1M
    
    i2H = Val(Left$(sTime2$, InStr(sTime2$, ":") - 1))
    i2M = Val(Mid$(sTime2$, InStr(sTime2$, ":") + 1))
    
    iDifH = i1H - i2H
    iDifM = i1M - i2M
'Debug.Print iDifH; "  "; iDifM
    
    iBarDataPeriodMins = iDifH * 60 + iDifM
    
End Sub
Public Sub CenterDlgBox(frmHwnd As Long)
    
    Dim hInst As Long
    Dim Thread As Long

   'Set up the CBT hook
   lFrmHwndCntrMsgBox = frmHwnd
   hInst = GetWindowLong(frmHwnd, GWL_HINSTANCE)
   Thread = GetCurrentThreadId()
   hHookCntrMsgBox = SetWindowsHookEx(WH_CBT, AddressOf CntrMsgBox, hInst, _
                            Thread)
    
End Sub
Private Function CntrMsgBox(ByVal lMsg As Long, ByVal wParam As Long, _
   ByVal lParam As Long) As Long

    Dim rectForm As RECT, rectMsg As RECT
    Dim x As Long, y As Long

   'On HCBT_ACTIVATE, show the MsgBox centered over Form1
   If lMsg = HCBT_ACTIVATE Then
      'Get the coordinates of the form and the message box so that
      'you can determine where the center of the form is located
      If lFrmHwndCntrMsgBox <> 0 Then
        GetWindowRect lFrmHwndCntrMsgBox, rectForm
        GetWindowRect wParam, rectMsg
        x = (rectForm.Left + (rectForm.Right - rectForm.Left) / 2) - _
            ((rectMsg.Right - rectMsg.Left) / 2)
        y = (rectForm.Top + (rectForm.Bottom - rectForm.Top) / 2) - _
            ((rectMsg.Bottom - rectMsg.Top) / 2)
      Else
        GetWindowRect GetDesktopWindow, rectForm
        GetWindowRect wParam, rectMsg
        x = (rectForm.Left + (rectForm.Right - rectForm.Left) / 2) - _
            ((rectMsg.Right - rectMsg.Left) / 2)
        y = (rectForm.Top + (rectForm.Bottom - rectForm.Top) / 2) - _
            ((rectMsg.Bottom - rectMsg.Top) / 2)
      End If
      
      'Position the msgbox
      SetWindowPos wParam, 0, x, y, 0, 0, _
                   SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
      'Release the CBT hook
      UnhookWindowsHookEx hHookCntrMsgBox
   End If
   CntrMsgBox = False

End Function

Public Sub Delay(rSeconds As Single)
    Dim rDelay As Single
    rDelay = Timer
    Do Until Timer - rDelay > rSeconds
        DoEvents
    Loop
End Sub
Public Sub PositionMousePointer(ByVal ihWnd As Long, iXoffsetFromLeft As Long, iYoffsetFromTop As Long, Optional isPixels As Boolean = True)
    'send mouse to specified position... AKA hotspot
    Dim recReturn As RECT, iX As Long, iY As Long
    Call GetWindowRect(ihWnd, recReturn)
    If isPixels = True Then
        iX = recReturn.Left + iXoffsetFromLeft
        iY = recReturn.Top + iYoffsetFromTop
    Else
        iX = recReturn.Left + iXoffsetFromLeft \ Screen.TwipsPerPixelX
        iY = recReturn.Top + iYoffsetFromTop \ Screen.TwipsPerPixelY
    End If
    Call SetCursorPos(iX, iY)

End Sub
Public Sub SaveBmp2File(bi24BitInfo As BITMAPINFO, bBytes() As Byte)
    Dim BmpHeader As BITMAPFILEHEADER, sOutFile As String
    
    sOutFile$ = App.Path & "\Snaps\Snap" & Format(Now, "mmddyyyy@hh.mm.ssa/p") & ".bmp"

    With BmpHeader
        .bfType = &H4D42
        .bfOffBits = Len(BmpHeader) + Len(bi24BitInfo.bmiHeader)
        .bfSize = .bfOffBits + bi24BitInfo.bmiHeader.biSizeImage
    End With
    Open sOutFile$ For Binary As #29
        Put #29, , BmpHeader
        Put #29, , bi24BitInfo.bmiHeader
        Put #29, , bBytes()
    Close #29
End Sub

