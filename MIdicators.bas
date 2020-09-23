Attribute VB_Name = "MIdicators"
Option Explicit

'indicator vars

Public iAvgLen1 As Integer
Public iAvgLen2 As Integer
Public iAvgLen3 As Integer
Public iAvgColor1 As Long
Public iAvgColor2 As Long
Public iAvgColor3 As Long
Public iAvgWidth1 As Integer
Public iAvgWidth2 As Integer
Public iAvgWidth3  As Integer

Public aIndData() As Double

Public MovAvgs As MovAvg
Public MACD1 As MACD
Public CCI1 As CCI
Public ROC1 As ROCPercent
Public RSI1 As RSI
Public STOCH1 As Stoch
Public oCurrentIndicator As Object
Private sCurrentIndicator As String
Public Property Get CurrentIndicator() As String
    CurrentIndicator = sCurrentIndicator
End Property

Public Property Let CurrentIndicator(sCurrentIndicatorA As String)
    'free up resources of current object
    Select Case sCurrentIndicator
        Case "CCI"
            Set CCI1 = Nothing
        Case "MACD"
            Set MACD1 = Nothing
        Case "ROC%"
            Set ROC1 = Nothing
        Case "RSI"
            Set RSI1 = Nothing
        Case "STOCH"
            Set STOCH1 = Nothing
    End Select
    'new indicator
    sCurrentIndicator = sCurrentIndicatorA
    'set new indicator and reference object and get settings for it
    Select Case sCurrentIndicator
        Case "CCI"
            Set CCI1 = New CCI
            Set oCurrentIndicator = CCI1
            CCI1.GetSavedSettings
        Case "MACD"
            Set MACD1 = New MACD
            Set oCurrentIndicator = MACD1
            MACD1.GetSavedSettings
        Case "ROC%"
            Set ROC1 = New ROCPercent
            Set oCurrentIndicator = ROC1
            ROC1.GetSavedSettings
        Case "RSI"
            Set RSI1 = New RSI
            Set oCurrentIndicator = RSI1
            RSI1.GetSavedSettings
        Case "STOCH"
            Set STOCH1 = New Stoch
            Set oCurrentIndicator = STOCH1
            STOCH1.GetSavedSettings
    End Select
End Property

Public Sub GetIndicatorSettings()
    On Error Resume Next
    
    oCurrentIndicator.GetSavedSettings
    
    iAvgLen1 = Val(GetIni(sINIsetFile, "AvgSettings", "AvgLen1"))
    iAvgColor1 = Val(GetIni(sINIsetFile, "AvgSettings", "AvgColor1"))
    iAvgWidth1 = Val(GetIni(sINIsetFile, "AvgSettings", "AvgWidth1"))
    iAvgLen2 = Val(GetIni(sINIsetFile, "AvgSettings", "AvgLen2"))
    iAvgColor2 = Val(GetIni(sINIsetFile, "AvgSettings", "AvgColor2"))
    iAvgWidth2 = Val(GetIni(sINIsetFile, "AvgSettings", "AvgWidth2"))
    iAvgLen3 = Val(GetIni(sINIsetFile, "AvgSettings", "AvgLen3"))
    iAvgColor3 = Val(GetIni(sINIsetFile, "AvgSettings", "AvgColor3"))
    iAvgWidth3 = Val(GetIni(sINIsetFile, "AvgSettings", "AvgWidth3"))
End Sub

Public Sub SaveIndicatorSettings()

    oCurrentIndicator.SaveCurrentSettings

    WriteIni sINIsetFile, "AvgSettings", "AvgLen1", CStr(iAvgLen1)
    WriteIni sINIsetFile, "AvgSettings", "AvgColor1", CStr(iAvgColor1)
    WriteIni sINIsetFile, "AvgSettings", "AvgWidth1", CStr(iAvgWidth1)
    WriteIni sINIsetFile, "AvgSettings", "AvgLen2", CStr(iAvgLen2)
    WriteIni sINIsetFile, "AvgSettings", "AvgColor2", CStr(iAvgColor2)
    WriteIni sINIsetFile, "AvgSettings", "AvgWidth2", CStr(iAvgWidth2)
    WriteIni sINIsetFile, "AvgSettings", "AvgLen3", CStr(iAvgLen3)
    WriteIni sINIsetFile, "AvgSettings", "AvgColor3", CStr(iAvgColor3)
    WriteIni sINIsetFile, "AvgSettings", "AvgWidth3", CStr(iAvgWidth3)
End Sub
Public Sub PlotIndicator()
    Dim i As Long, j As Long, sIndLabel As String, iPlotColors(0 To 3, 0 To 1) As Long
    Dim rString As Single, rStep As Single, rMin As Single, rMax As Single
    Dim rIndRange As Single, iPrevIndW As Integer, dHeightInd As Double
    Dim rIndY1 As Single, iZeroPosition As Long, x As Single, Y1 As Single, Y2 As Single
    Dim iCount As Integer, iUB As Integer, sText As String, X2 As Single, iUB2 As Integer
    Dim rGridValue As Single, rUpperTrig As Single, rLowerTrig As Single, iMaxTxtWidth As Long
'    For i = 0 To iUBaData
'        aTmp(i).dOpen = aData(i).dOpen
'        aTmp(i).dHigh = aData(i).dHigh
'        aTmp(i).dLow = aData(i).dLow
'        aTmp(i).dClose = aData(i).dClose
'        aTmp(i).iVol = aData(i).iVol
'    Next
    
    'iPlotColors() has 2 dimensions, 1st is the color of plot,
    '2nd is the type of plot-> line=0 or histogram=1
    Select Case CurrentIndicator$
        Case "MACD"
            Call LoadDblArray
            sIndLabel$ = CurrentIndicator$ _
                    & "(" & MACD1.MACDLen1 & "," & MACD1.MACDLen2 & "," & MACD1.MACDLen3 & ")"
            rMin = 99999 'set at high level so it will be lowered by real data
            iPlotColors(0, 0) = MACD1.MACDColor1
            iPlotColors(1, 0) = MACD1.MACDColor2
            iPlotColors(2, 0) = MACD1.MACDColor3
            iPlotColors(0, 1) = 0  'line
            iPlotColors(1, 1) = 0  'line
            iPlotColors(2, 1) = 1  'histogram
        Case "CCI"
            Call LoadStockDataArray
            sIndLabel$ = CurrentIndicator$ & "(" & CCI1.CciLen & "," & CCI1.CciAvgLen & "," & CCI1.CciAvgAvgLen & ")"
            rMax = CCI1.UpperTrig   'start max at the trigger
            rMin = CCI1.LowerTrig    'start min at the lower trig
            iPlotColors(0, 0) = CCI1.CciColor
            iPlotColors(1, 0) = CCI1.CciAvgColor
            iPlotColors(0, 1) = 0  'line
            iPlotColors(1, 1) = 0  'line
        Case "ROC%"
            Call LoadDblArray
            sIndLabel$ = CurrentIndicator$ _
                    & "(" & ROC1.ROCLen & ")"
            rMax = ROC1.UpperTrig    'start max at the trigger
            rMin = ROC1.LowerTrig     'start min at the lower trig
            iPlotColors(0, 0) = ROC1.ROCColor
            iPlotColors(0, 1) = 1  'histogram
        Case "RSI"
            Call LoadDblArray
            sIndLabel$ = CurrentIndicator$ _
                    & "(" & RSI1.RsiLen & ")"
            rMax = 100   'set max at 100
            rMin = 0   'set min at 0
            iPlotColors(0, 0) = RSI1.RsiColor
            iPlotColors(0, 1) = 0  'line
        Case "STOCH"
            Call LoadStockDataArray
            sIndLabel$ = CurrentIndicator$ & "(" & STOCH1.KPeriod & "," & STOCH1.DPeriod & ")"
            If STOCH1.StochSlow Then  'plot stochastic slowK
                iPlotColors(0, 0) = STOCH1.StochKSlowColor
                iPlotColors(1, 0) = STOCH1.StochDColor
            Else
                iPlotColors(0, 0) = STOCH1.StochKColor
                iPlotColors(1, 0) = STOCH1.StochKSlowColor
            End If
            iPlotColors(0, 1) = 0  'line
            iPlotColors(1, 1) = 0  'line
            rMax = STOCH1.UpperTrig   'start max at the trigger
            rMin = STOCH1.LowerTrig   'start min at the lower trig
    End Select

    dHeightInd = iBottomPlotMargin - rSplit2 - 15
    
    iUB = UBound(aIndData(), 1)
    iUB2 = UBound(aIndData(), 2)
    For j = 0 To iUB2  'find min & max values
        For i = iStartIndex - iNumBarsPloted To iStartIndex
            If aIndData(i, j) > rMax Then rMax = aIndData(i, j)
            If aIndData(i, j) < rMin Then rMin = aIndData(i, j)
        Next
    Next
    rIndRange = rMax - rMin
    
    With frmMain.ChartBox
        .DrawWidth = 1
        .DrawStyle = vbDot
        .DrawMode = vbCopyPen
        rStep = (((rIndRange / 5) * dHeightInd) / rIndRange) 'grid spacing
        iZeroPosition = 4 + rSplit2 + (((rMax - 0) * dHeightInd) / rIndRange)
        
        'indicator pane Hgrid
        If rMin <= 0 Then
            For Y2 = iZeroPosition + rStep To iBottomPlotMargin - 10 Step rStep 'zero down
                .CurrentY = Y2 - iTextHeight / 2
                If .CurrentY < iBottomPlotMargin - iTextHeight Then
                    'print value of grid line
                    rGridValue = (iZeroPosition - Y2) / (dHeightInd / rIndRange)
                    If rGridValue < -10 Then  'round values > -10 to save space
                        sText$ = Round(rGridValue)
                    Else
                        sText$ = Format(rGridValue, "##.##")
                    End If
                    rString = .TextWidth(sText$)
                    If rString > iMaxTxtWidth Then iMaxTxtWidth = rString
                    .CurrentX = xRightMargin - rString - 10
                    frmMain.ChartBox.Print sText$
                End If
                frmMain.ChartBox.Line (xLeftMargin, Y2)-(xRightMargin - rString - 20, Y2), iGridColor
            Next Y2
        End If
        For Y2 = iZeroPosition - rStep To rSplit2 + 4 Step -rStep 'zero up
            .CurrentY = Y2 - iTextHeight / 2
'Debug.Print .CurrentY; rSplit2
            If .CurrentY < rSplit2 + iTextHeight Then
                .CurrentY = Y2
            End If
            'print value of grid line
            rGridValue = (iZeroPosition - Y2) / (dHeightInd / rIndRange)
            If rGridValue > 10 Then
                sText$ = Round(rGridValue)
            Else
                sText$ = Format(rGridValue, "##.##")
            End If
            rString = .TextWidth(sText$)
            If rString > iMaxTxtWidth Then iMaxTxtWidth = rString
            .CurrentX = xRightMargin - rString - 10
            frmMain.ChartBox.Print sText$
            frmMain.ChartBox.Line (xLeftMargin, Y2)-(xRightMargin - rString - 20, Y2), iGridColor
        Next Y2
        
        'plot zero line and legion text
        rString = .TextWidth(CStr(0))
        If oCurrentIndicator.ZeroLinePlot Then
            .DrawStyle = vbSolid
            frmMain.ChartBox.Line (xLeftMargin, iZeroPosition)-((xRightMargin - rString - 20), iZeroPosition), RGB(0, 0, 160)
        Else
            frmMain.ChartBox.Line (xLeftMargin, iZeroPosition)-((xRightMargin - rString - 20), iZeroPosition), iGridColor
        End If
        .CurrentX = xRightMargin - rString - 10
        rString = iTextHeight / 2
        .CurrentY = .CurrentY - rString
        frmMain.ChartBox.Print CStr(0)  'print zero legion text
        
        'start the ind. data plot
        .DrawStyle = vbSolid
        iPrevIndW = .DrawWidth
        .DrawWidth = oCurrentIndicator.PlotWidth

        If oCurrentIndicator.PlotTriggerLines Then
            'plot upper trigger
            rUpperTrig = oCurrentIndicator.UpperTrig
            rLowerTrig = oCurrentIndicator.LowerTrig
            Y1 = 4 + rSplit2 + (((rMax - rUpperTrig) * dHeightInd) / rIndRange)
            rString = .TextWidth(CStr(rUpperTrig)) + iMaxTxtWidth + 3
            frmMain.ChartBox.Line (xLeftMargin, Y1)-(xRightMargin - rString - 20, Y1), RGB(0, 0, 160)
            .CurrentX = xRightMargin - rString - 10
            rString = iTextHeight / 2
            .CurrentY = .CurrentY - rString + 3
            frmMain.ChartBox.Print CStr(rUpperTrig)
            'lower trigger
            Y1 = 4 + rSplit2 + (((rMax - rLowerTrig) * dHeightInd) / rIndRange)
            rString = .TextWidth(CStr(rLowerTrig)) + iMaxTxtWidth + 3
            frmMain.ChartBox.Line (xLeftMargin, Y1)-(xRightMargin - rString - 20, Y1), RGB(0, 0, 160)
            .CurrentX = xRightMargin - rString - 10
            rString = iTextHeight / 2
            .CurrentY = .CurrentY - rString
            frmMain.ChartBox.Print CStr(rLowerTrig)
        End If
        
        'plot data
        x = rRightSideOffset
        iCount = iStartIndex
        sText$ = sEmpty
        For i = 0 To iUB2
            Do While x > -1 And iCount >= 0
                Y1 = 4 + rSplit2 + (((rMax - aIndData(iCount, i)) _
                    * dHeightInd) / rIndRange)
                
                
                If X2 <> 0 Then
                    If iPlotColors(i, 1) = 0 Then 'line plot
                        frmMain.ChartBox.Line (X2, Y2)-(x, Y1), iPlotColors(i, 0)
                    ElseIf iPlotColors(i, 1) = 1 Then 'histogram
                        frmMain.ChartBox.Line (x, iZeroPosition)-(x, Y1), iPlotColors(i, 0)
                    End If
                End If
                X2 = x
                Y2 = Y1
                iCount = iCount - 1
                x = x - iBarSpacing
            Loop
            x = rRightSideOffset
            iCount = iStartIndex
            X2 = 0
            'get the last ind. value for printout
            sText$ = sText$ & CStr(Format(aIndData(iUB, i), "##.00"))
            If i <> iUB2 Then sText$ = sText$ & "; "
        Next
        .DrawWidth = iPrevIndW
        'print ind. label text
        sText$ = sIndLabel$ & ": " & sText$
        'draw a "blackout rect for better visibility of the text
        frmMain.ChartBox.Line (1, rSplit2 + 3)-(1 + .TextWidth(sText$), rSplit2 + 3 + .TextHeight(sText$)), iBackColor, BF
        .CurrentX = 1
        .CurrentY = rSplit2 + 3
        frmMain.ChartBox.Print sText$
    End With
End Sub
Private Sub LoadDblArray()
    Dim i As Integer, aArray() As Double
    
    ReDim aArray(0 To iUBaData)
    For i = 0 To iUBaData
        aArray(i) = aData(i).dClose  'calculations only need close
    Next
            
    oCurrentIndicator.Calculate aArray()
    
    Select Case CurrentIndicator$
        Case "MACD"
            ReDim aIndData(oCurrentIndicator.RetBoundLo To oCurrentIndicator.RetBoundHi, 0 To 2)
            For i = oCurrentIndicator.RetBoundLo To oCurrentIndicator.RetBoundHi
                aIndData(i, 0) = oCurrentIndicator.RetVal(i)
                aIndData(i, 1) = oCurrentIndicator.RetValSig(i)
                aIndData(i, 2) = oCurrentIndicator.RetValHist(i)
            Next
        Case "RSI", "ROC%"
            ReDim aIndData(oCurrentIndicator.RetBoundLo To oCurrentIndicator.RetBoundHi, 0 To 0)
            For i = oCurrentIndicator.RetBoundLo To oCurrentIndicator.RetBoundHi
                aIndData(i, 0) = oCurrentIndicator.RetVal(i)
            Next
    End Select
        
End Sub

Private Sub LoadStockDataArray()
    Dim i As Integer, aArray() As StockData
    
    ReDim aArray(0 To iUBaData)
    For i = 0 To iUBaData
        aArray(i) = aData(i)
    Next
    
    
    Select Case CurrentIndicator$
        Case "CCI"
            CCI1.Calculate aArray()
            ReDim aIndData(CCI1.RetBoundLo To CCI1.RetBoundHi, 0 To 1)
            For i = CCI1.RetBoundLo To CCI1.RetBoundHi
                aIndData(i, 0) = CCI1.RetValAvg(i)
                aIndData(i, 1) = CCI1.RetValAvgAvg(i)
            Next
        Case "STOCH"
            STOCH1.Calculate aArray()
            ReDim aIndData(STOCH1.RetBoundLo To STOCH1.RetBoundHi, 0 To 1)
            If STOCH1.StochSlow Then  'plot stochastic slowK
                For i = STOCH1.RetBoundLo To STOCH1.RetBoundHi
                    aIndData(i, 0) = STOCH1.RetValKSlow(i)
                    aIndData(i, 1) = STOCH1.RetValD(i)
                Next
            Else
                For i = STOCH1.RetBoundLo To STOCH1.RetBoundHi
                    aIndData(i, 0) = STOCH1.RetValK(i)
                    aIndData(i, 1) = STOCH1.RetValKSlow(i)
                Next
            End If
    End Select
End Sub
Public Sub PlotAvg()
    Dim iPeriod As Long, iColor As Long, j As Long
    Dim aMovAvg() As Double, x As Integer, y As Integer, xPrev As Integer, yPrev As Integer
    Dim i As Integer, iCount As Integer, aTemp() As Double, iOldStyle As Long, iOldWidth As Long
    
    ReDim aTemp(0 To iUBaData)
    For i = 0 To iUBaData
        aTemp(i) = aData(i).dClose
    Next i
    
    With frmMain.ChartBox
        iOldStyle = .DrawStyle  'save current draw settings
        iOldWidth = .DrawWidth
        .DrawStyle = vbSolid
        For j = 1 To 3
            Select Case j
                Case 1
                    .DrawWidth = iAvgWidth1
                    MovAvgs.AvgLen = iAvgLen1
                    iColor = iAvgColor1
                Case 2
                    .DrawWidth = iAvgWidth2
                    MovAvgs.AvgLen = iAvgLen2
                    iColor = iAvgColor2
                Case 3
                    .DrawWidth = iAvgWidth3
                    MovAvgs.AvgLen = iAvgLen3
                    iColor = iAvgColor3
            End Select
            
            MovAvgs.Calculate aTemp()
            ReDim aMovAvg(MovAvgs.RetBoundLo To MovAvgs.RetBoundHi)
            For i = MovAvgs.RetBoundLo To MovAvgs.RetBoundHi
                aMovAvg(i) = MovAvgs.RetVal(i)
            Next i
            
            x = .ScaleWidth - iBlankSpace * 10
            rRightSideOffset = x
            iCount = iStartIndex
            
            Do While x > -1 And iCount > 1
                If aMovAvg(iCount) <> 0 Then
'Debug.Print aMovAvg(iCount)
                    y = 4 + (((dMaxPrice - aMovAvg(iCount)) * dHeightPrice) / dRangePrice)
                    If x < rRightSideOffset And y < rSplit1 Then
                        frmMain.ChartBox.Line (x, y)-(xPrev, yPrev), iColor
                    End If
                    xPrev = x
                    yPrev = y
                End If
                iCount = iCount - 1
                x = x - iBarSpacing
            Loop
        Next
        .DrawStyle = iOldStyle  'restore old settings
        .DrawWidth = iOldWidth
    End With
End Sub

