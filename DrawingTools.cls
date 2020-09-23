VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDrawingTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event DrawingInstructions(sText As String)
Public Event DrawLoopIsRunning()
Public Event DrawingStarted()
Public Event DrawingDone()

Private lDrawModeOld As Long, iOldStyle As Long, lOldWidth As Long, iFillOld As Integer
Private iPicTimer As Long, ltimer As Single, iPanel As Integer
Private rectOrg As RECT, iExtend As Long, iToolFillColor As Long, iOldFillColor As Long
Private iToolMode As Integer, iToolWidth As Long, iToolStyle As Long
Private iToolColor As Long, objpicBxV As PictureBox, objpicBx As PictureBox
Private fMouseClickEnabled As Boolean, iMouseClickNum As Long
Private iXcurr As Long, iYcurr As Long, iYcurrMov As Long, iXcurrMov As Long
Private iToolFillStyle As Integer, fUseOrigin As Boolean
Private X1 As Single, Y1 As Single, X2 As Single, Y2 As Single
Private XextR As Single, YextR As Single, XextL As Single, YextL As Single
Private diffX1 As Single, diffY1 As Single, diffX2 As Single, diffY2 As Single

Public Sub TrendLine(Optional fDrawParallel As Boolean = False)

    Dim rRect As RECT, iHasRun As Integer, iNumOfMousePts As Long
'*******************
    If fDrawParallel Then
        Dim X3 As Single, Y3 As Single
        iNumOfMousePts = 3
    Else
        iNumOfMousePts = 2
    End If
    Call InitDrawingTool
    On Error Resume Next
    RaiseEvent DrawingInstructions("Pick 1st Pt.")
    Do While iMouseClickNum < iNumOfMousePts
        DoEvents
        If fCancelDrawingTool Then Call CancelDrawing: Exit Sub
        If iMouseClickNum = 1 Then
            If iHasRun = 0 Then 'do onetime ops for this leg of the loop
                X1 = Xcurr
                Y1 = Ycurr
                iHasRun = 1  'prevent from entering this section again
                RaiseEvent DrawingInstructions("Pick Last Pt.")
                iPanel = Panel2Draw(Ycurr)
                Call SaveGdiInfo1
                picBxV.DrawStyle = ToolStyle
                picBxV.DrawMode = ToolMode
                picBxV.DrawWidth = ToolWidth
            End If
            'prevent the loop from rapid fire drawing by both setting a minimum
            'elapsed time and looking for mouse movement... if the mouse isn't moving
            'the rubberband line isn't either..... prevents flicker and slow line response
            If Timer - ltimer > 0.1 And ((XcurrMov <> X2) Or (YcurrMov <> Y2)) Then
                picBxV.Refresh
                X2 = XcurrMov
                Y2 = YcurrMov
                picBxV.Line (X1, Y1)-(X2, Y2), ToolColor
                ltimer = Timer
            End If
        ElseIf iMouseClickNum = 2 Then '2nd pt has been clicked
            If iHasRun = 1 Then
                X2 = Xcurr
                Y2 = Ycurr
                If Not fDrawParallel Then Exit Do 'we have all pts for 1 line
                iHasRun = 2  'lockout this if-then section
                RaiseEvent DrawingInstructions("Place Parallel Line")
                diffX1 = ((X1 + X2) / 2) - X1  'get the line segments from the mid pt
                diffY1 = ((Y1 + Y2) / 2) - Y1  'to each end of the 1st line
                diffX2 = ((X1 + X2) / 2) - X2
                diffY2 = ((Y1 + Y2) / 2) - Y2
            End If
            'flicker control
            If Timer - ltimer > 0.1 And ((XcurrMov <> X3) Or (YcurrMov <> Y3)) Then
                picBxV.Refresh
                picBxV.Line (X1, Y1)-(X2, Y2), ToolColor
                X3 = XcurrMov  'pts for the parallel line... it will have the same
                Y3 = YcurrMov  'slope as the 1st. Use these for the mid pt.
                picBxV.Line (X3 - diffX1, Y3 - diffY1)-(X3 - diffX2, Y3 - diffY2), ToolColor
                ltimer = Timer
            End If
        ElseIf iMouseClickNum = 3 Then
            X3 = Xcurr
            Y3 = Ycurr
            diffX1 = X3 - diffX1
            diffY1 = Y3 - diffY1
            diffX2 = X3 - diffX2
            diffY2 = Y3 - diffY2
        End If
        Call DrawLoopEventTimer
    Loop
    Call RestoreSaveGdiInfo2
    picBx.DrawStyle = ToolStyle
    picBx.DrawMode = ToolMode
    picBx.DrawWidth = ToolWidth
    'extension routine.. 0= no ext., 1=right only, 2=left only, 3=both
    If iExtend = 0 Then
        picBx.Line (X1, Y1)-(X2, Y2), ToolColor
        If fDrawParallel Then picBx.Line (diffX1, diffY1)-(diffX2, diffY2), ToolColor
    Else
        If Y1 = Y2 Then 'horz line...  no reason to do the calculations
            If iExtend And 1 Then X2 = iMaxDrawRightX  'ext. right
            If iExtend And 10 Then X1 = 0  'ext left
            picBx.Line (X1, Y1)-(X2, Y2), ToolColor
            If fDrawParallel Then
                picBx.Line (X1, diffY1)-(X2, diffY2), ToolColor
            End If
        Else
            XextR = X2: YextR = Y2
            XextL = X1: YextL = Y1
            'if x>x2 then we want to switch the pts around
            If X1 > X2 Then
                XextR = X1: YextR = Y1
                XextL = X2: YextL = Y2
            End If
            If iExtend And 1 Then Call GetXYExtentLine  'ext. right
            If iExtend And 10 Then Call GetXYExtentLine(False) 'ext left
            picBx.Line (XextL, YextL)-(XextR, YextR), ToolColor
            If fDrawParallel Then
                XextR = diffX2: YextR = diffY2
                XextL = diffX1: YextL = diffY1
                If diffX1 > diffX2 Then
                    XextR = diffX1: YextR = diffY1
                    XextL = diffX2: YextL = diffY2
                End If
                If iExtend And 1 Then Call GetXYExtentLine     'ext. right
                If iExtend And 10 Then Call GetXYExtentLine(False)    'ext left
                picBx.Line (XextL, YextL)-(XextR, YextR), ToolColor
            End If
        End If
    End If

    picBxV.Picture = picBx.Image
    Call RestoreGdiInfo3

End Sub
Private Sub GetXYExtentLine(Optional iDirRight As Long = 1)
    Dim m As Single

    m = (YextR - YextL) / (XextR - XextL) 'slope of line

    If iDirRight = 1 Then 'extend right
        If m > 0 Then 'slope is negative, use lower divider y cross
            'the y coordinate system is backwards so m is positive when slope negative
            If iPanel = 1 Then
                YextR = rSplit1
            ElseIf iPanel = 2 Then
                YextR = rSplit2
            ElseIf iPanel = 3 Then
                YextR = iBottomPlotMargin
            End If
        Else 'slope positive or zero
            If iPanel = 1 Then
                YextR = 0
            ElseIf iPanel = 2 Then
                YextR = rSplit1
            ElseIf iPanel = 3 Then
                YextR = rSplit2
            End If
        End If
        XextR = (YextR - YextL) / m + XextL
        If XextR > iMaxDrawRightX Then
            XextR = iMaxDrawRightX
            YextR = (XextR - XextL) * m + YextL
        End If
    Else  'extend left
        If m > 0 Then 'slope is negative, use upper divider y cross
            'the y coordinate system is backwards so m is positive when slope negative
            If iPanel = 1 Then
                YextL = 0
            ElseIf iPanel = 2 Then
                YextL = rSplit1
            ElseIf iPanel = 3 Then
                YextL = rSplit2
            End If
        Else 'slope positive or zero
            If iPanel = 1 Then
                YextL = rSplit1
            ElseIf iPanel = 2 Then
                YextL = rSplit2
            ElseIf iPanel = 3 Then
                YextL = iBottomPlotMargin
            End If
        End If
        XextL = (YextL - YextR) / m + XextR
        If XextL < 0 Then
            XextL = 0
            YextL = (XextL - XextR) * m + YextR
        End If
    End If
End Sub
Public Function RectAndSquareTool(Optional fDrawSquare As Boolean = False)

    Dim iHasRun As Integer, Pt1 As POINTSINGLE, Pt2 As POINTSINGLE
    Dim rRect As RECT, fValidDraw As Boolean
'*******************

    Call InitDrawingTool
    On Error Resume Next
    RaiseEvent DrawingInstructions("Pick 1st Corner Pt.")
    Do While iMouseClickNum < 2
        DoEvents
        If fCancelDrawingTool Then Call CancelDrawing: Exit Function
        If iMouseClickNum = 1 Then
            If iHasRun = 0 Then
                Pt1.x = Xcurr
                Pt1.y = Ycurr
                iHasRun = 1
                RaiseEvent DrawingInstructions("Pick Last Corner Pt.")
                iPanel = Panel2Draw(Ycurr)
                Call SaveGdiInfo1
                picBxV.DrawStyle = ToolStyle
                picBxV.DrawMode = vbMergePen
                picBxV.DrawWidth = ToolWidth
                picBxV.FillColor = ToolFillColor
                picBxV.FillStyle = ToolFillStyle
            End If
            If Timer - ltimer > 0.1 And ((XcurrMov <> Pt2.x) Or (YcurrMov <> Pt2.y)) Then
                picBxV.Refresh
                Pt2.x = XcurrMov
                Pt2.y = YcurrMov
                If fDrawSquare Then
                    fValidDraw = SqrFrom2PointsF(Pt1, Pt2, rRect)
                Else
                    fValidDraw = RectFrom2PointsF(Pt1, Pt2, rRect)
                End If
                If fValidDraw Then picBxV.Line (rRect.Left, rRect.Top)-(rRect.Right, rRect.Bottom), ToolColor, B
                ltimer = Timer
            End If
        ElseIf iMouseClickNum = 2 Then
            picBxV.DrawMode = ToolMode
            Pt2.x = Xcurr
            Pt2.y = Ycurr
            If fDrawSquare Then
                fValidDraw = SqrFrom2PointsF(Pt1, Pt2, rRect)
            Else
                fValidDraw = RectFrom2PointsF(Pt1, Pt2, rRect)
            End If
        End If
        Call DrawLoopEventTimer
    Loop
    Call RestoreSaveGdiInfo2
    picBx.FillStyle = ToolFillStyle
    picBx.FillColor = ToolFillColor
    picBx.DrawStyle = ToolStyle
    picBx.DrawMode = vbMergePen
    picBx.DrawWidth = ToolWidth
    If fValidDraw Then picBx.Line (rRect.Left, rRect.Top)-(rRect.Right, rRect.Bottom), ToolColor, B
    picBxV.Picture = picBx.Image
    Call RestoreGdiInfo3
End Function

Public Function CircleElipseTool(Optional fDrawCirc As Boolean = False)
    Dim iHasRun As Integer
    Dim rRect As RECT, fValidDraw As Boolean, rDist As Single, oldForeColor As Long
    Dim X1 As Single, Y1 As Single, X2 As Single, Y2 As Single
'*******************

    Call InitDrawingTool
    On Error Resume Next
    RaiseEvent DrawingInstructions("Pick Origin")
    Do While iMouseClickNum < 2
        DoEvents
        If fCancelDrawingTool Then Call CancelDrawing: Exit Function
        If iMouseClickNum = 1 Then
            If iHasRun = 0 Then
                X1 = Xcurr
                Y1 = Ycurr
                iHasRun = 1
                RaiseEvent DrawingInstructions("Pick Last Pt.")
                iPanel = Panel2Draw(Ycurr)
                Call SaveGdiInfo1
                picBxV.DrawStyle = ToolStyle
                picBxV.DrawMode = vbMergePen
                picBxV.DrawWidth = ToolWidth
                picBxV.FillColor = ToolFillColor
                picBxV.FillStyle = ToolFillStyle
            End If
            If Timer - ltimer > 0.1 And ((XcurrMov <> X2) Or (YcurrMov <> Y2)) Then
                picBxV.Refresh
                X2 = XcurrMov
                Y2 = YcurrMov
                If fDrawCirc Then
                    rDist = Sqr((X2 - X1) * (X2 - X1) + (Y2 - Y1) * (Y2 - Y1)) 'radius
                    If fUseOrigin Then
                        picBxV.Circle (X1, Y1), rDist, ToolColor  'origin + radius
                    Else 'opp sides on diameter
                        picBxV.Circle ((X1 + X2) / 2, (Y1 + Y2) / 2), rDist / 2, ToolColor
                    End If
                Else
                    'could select a gdi pen here but using the forecolor works also
                    oldForeColor = picBxV.ForeColor
                    picBxV.ForeColor = ToolColor
                    If fUseOrigin Then
                        'origin + corner
                        Ellipse picBxV.hDC, X1 * 2 - X2, Y1 * 2 - Y2, X2, Y2
                    Else
                        Ellipse picBxV.hDC, X1, Y1, X2, Y2  'diagonal corners
                    End If
                    picBxV.ForeColor = oldForeColor
                End If
                ltimer = Timer
                End If
        ElseIf iMouseClickNum = 2 Then
            picBxV.DrawMode = ToolMode
            X2 = Xcurr
            Y2 = Ycurr
        End If
        Call DrawLoopEventTimer
    Loop
    Call RestoreSaveGdiInfo2
    picBx.FillStyle = ToolFillStyle
    picBx.FillColor = ToolFillColor
    picBx.DrawStyle = ToolStyle
    picBx.DrawMode = vbMergePen
    picBx.DrawWidth = ToolWidth
    If fDrawCirc Then
        rDist = Sqr((X2 - X1) * (X2 - X1) + (Y2 - Y1) * (Y2 - Y1)) 'radius
        If fUseOrigin Then
            picBx.Circle (X1, Y1), rDist, ToolColor  'origin + radius
        Else 'opp sides on diameter
            picBx.Circle ((X1 + X2) / 2, (Y1 + Y2) / 2), rDist / 2, ToolColor
        End If
    Else
        oldForeColor = picBx.ForeColor
        picBx.ForeColor = ToolColor
        If fUseOrigin Then
            'origin + corner
            Ellipse picBx.hDC, X1 * 2 - X2, Y1 * 2 - Y2, X2, Y2
        Else
            Ellipse picBx.hDC, X1, Y1, X2, Y2  'diagonal corners
        End If
        picBx.ForeColor = oldForeColor
    End If
    picBxV.Picture = picBx.Image
    Call RestoreGdiInfo3
End Function

Public Sub FibRetrace()

    Dim rRect As RECT, iHasRun As Integer, iNumOfMousePts As Long, rYdiff As Single
    Dim rFib61 As Single, rFib50 As Single, rFib38 As Single, iXstart As Long
    Dim rXvecDiff As Single, rYvecDiff As Single, rAngle As Single, rYstart As Single
'*******************

    iNumOfMousePts = 2
    Call InitDrawingTool
    On Error Resume Next
    RaiseEvent DrawingInstructions("Pick 1st Extreme.")
    Do While iMouseClickNum < iNumOfMousePts
        DoEvents
        If fCancelDrawingTool Then Call CancelDrawing: Exit Sub
        If iMouseClickNum = 1 Then
            If iHasRun = 0 Then 'do onetime ops for this leg of the loop
                X1 = Xcurr
                Y1 = Ycurr
                iHasRun = 1  'prevent from entering this section again
                RaiseEvent DrawingInstructions("Pick Opposite Extreme")
                iPanel = Panel2Draw(Ycurr)
                Call SaveGdiInfo1
                picBxV.DrawMode = ToolMode
            End If
            'prevent the loop from rapid fire drawing by both setting a minimum
            'elapsed time and looking for mouse movement... if the mouse isn't moving
            'the rubberband line isn't either..... prevents flicker and slow line response
            If Timer - ltimer > 0.1 And ((XcurrMov <> X2) Or (YcurrMov <> Y2)) Then
                picBxV.DrawStyle = vbDot
                picBxV.DrawWidth = 1
                picBxV.Refresh
                X2 = XcurrMov
                Y2 = YcurrMov
                rXvecDiff = X2 - X1
                rYvecDiff = Y2 - Y1
                rAngle = Tan(rYvecDiff / rXvecDiff) * 180 / PI
                picBxV.Line (X1, Y1)-(X2, Y2), ToolColor
                picBxV.Line (X1 + Sin(rAngle), Y1 + Cos(rAngle))-(X2 + Sin(rAngle), Y2 + Cos(rAngle)), ToolColor
                rYdiff = Abs(rYvecDiff)
                rFib61 = rYdiff * 0.618
                rFib50 = rYdiff * 0.5
                rFib38 = rYdiff * 0.382
                If Y1 < Y2 Then
                    rYstart = Y1
                Else
                    rYstart = Y2
                End If
                If X1 > X2 Then
                    iXstart = X2
                Else
                    iXstart = X1
                End If
                picBxV.DrawStyle = ToolStyle
                picBxV.DrawWidth = ToolWidth
                picBxV.Line (iXstart, rYstart)-(iMaxDrawRightX, rYstart), ToolColor
                picBxV.Line (iXstart, rYstart + rFib61)-(iMaxDrawRightX, rYstart + rFib61), ToolColor
                picBxV.Line (iXstart, rYstart + rFib50)-(iMaxDrawRightX, rYstart + rFib50), ToolColor
                picBxV.Line (iXstart, rYstart + rFib38)-(iMaxDrawRightX, rYstart + rFib38), ToolColor
                picBxV.Line (iXstart, rYstart + rYdiff)-(iMaxDrawRightX, rYstart + rYdiff), ToolColor

                ltimer = Timer
            End If
        ElseIf iMouseClickNum = 2 Then '2nd pt has been clicked
            X2 = Xcurr
            Y2 = Ycurr
            rXvecDiff = X2 - X1
            rYvecDiff = Y2 - Y1
            rAngle = Tan(rYvecDiff / rXvecDiff) * 180 / PI
            rYdiff = Abs(rYvecDiff)
            rFib61 = rYdiff * 0.618
            rFib50 = rYdiff * 0.5
            rFib38 = rYdiff * 0.382
            If Y1 < Y2 Then
                rYstart = Y1
            Else
                rYstart = Y2
            End If
            If X1 > X2 Then
                iXstart = X2
            Else
                iXstart = X1
            End If
        End If
        Call DrawLoopEventTimer
    Loop
    Call RestoreSaveGdiInfo2
    picBx.DrawStyle = ToolStyle
    picBx.DrawMode = ToolMode
    picBx.DrawWidth = ToolWidth
    picBx.Line (X1, Y1)-(X2, Y2), ToolColor
    picBx.Line (iXstart, rYstart)-(iMaxDrawRightX, rYstart), ToolColor
    picBx.Line (iXstart, rYstart + rFib61)-(iMaxDrawRightX, rYstart + rFib61), ToolColor
    picBx.Line (iXstart, rYstart + rFib50)-(iMaxDrawRightX, rYstart + rFib50), ToolColor
    picBx.Line (iXstart, rYstart + rFib38)-(iMaxDrawRightX, rYstart + rFib38), ToolColor
    picBx.Line (iXstart, rYstart + rYdiff)-(iMaxDrawRightX, rYstart + rYdiff), ToolColor

    picBxV.Picture = picBx.Image
    Call RestoreGdiInfo3

End Sub
Private Sub DrawLoopEventTimer()
    If Timer - iPicTimer > 1 Then
        RaiseEvent DrawLoopIsRunning
        iPicTimer = Timer
    End If
End Sub
Private Sub InitDrawingTool()
    fCancelDrawingTool = False
    IsDrawing = True
    Screen.MouseIcon = LoadResPicture(101, vbResCursor)
    Screen.MousePointer = vbCustom
    ltimer = Timer
    iPicTimer = Timer
    iMouseClickNum = 0: fMouseClickEnabled = True
    RaiseEvent DrawingStarted
End Sub

Private Sub SaveGdiInfo1()
    'save all original pb settings
    iFillOld = picBxV.FillStyle
    iOldFillColor = picBxV.FillColor
    iOldStyle = picBxV.DrawStyle
    lOldWidth = picBxV.DrawWidth
    lDrawModeOld = picBxV.DrawMode
End Sub
Private Sub RestoreSaveGdiInfo2()
    'restore settings to 1st pb and save 2nd ones
    On Error Resume Next
    picBxV.DrawMode = lDrawModeOld
    picBxV.FillStyle = iFillOld
    picBxV.FillColor = iOldFillColor
    picBxV.DrawWidth = lOldWidth
    picBxV.DrawStyle = iOldStyle
    iOldStyle = picBx.DrawStyle
    lOldWidth = picBx.DrawWidth
    lDrawModeOld = picBx.DrawMode
    iFillOld = picBx.FillStyle
    iOldFillColor = picBx.FillColor
End Sub
Private Sub RestoreGdiInfo3()
    'restore 2nd pb settings and clean up
    On Error Resume Next
    picBx.DrawMode = lDrawModeOld
    picBx.FillStyle = iFillOld
    picBx.FillColor = iOldFillColor
    picBx.DrawWidth = lOldWidth
    picBx.DrawStyle = iOldStyle
    RaiseEvent DrawingDone
    fMouseClickEnabled = False
    IsDrawing = False
    Call ClipCursor(rectOrg)
    Screen.MousePointer = vbDefault
End Sub
Private Sub CancelDrawing()
    Call RestoreSaveGdiInfo2
    Call RestoreGdiInfo3
    picBxV.Cls
    picBx.Cls
    IsDrawing = False
End Sub
Private Function Panel2Draw(iYcurr As Long) As Long
    'deterimine which panel to draw to
    Dim lpRect As RECT, iCaptionHeight As Long

    iCaptionHeight = GetSystemMetrics(SM_CYCAPTION)
    lpRect.Left = picBxV.Left \ Screen.TwipsPerPixelX
    lpRect.Right = (picBxV.Left + picBxV.Width) \ Screen.TwipsPerPixelX
    Select Case iYcurr
        Case Is < rSplit1  'price panel
            lpRect.Top = picBxV.ScaleTop + iCaptionHeight
            lpRect.Bottom = rSplit1 + iCaptionHeight
            Panel2Draw = 1
        Case Is > rSplit2  'indicator panel
            lpRect.Top = (picBxV.Top + rSplit2 + 2) + iCaptionHeight
            lpRect.Bottom = (picBxV.Top + iBottomPlotMargin) + iCaptionHeight
            Panel2Draw = 2
        Case Else  'vol panel
            lpRect.Top = (picBxV.Top + rSplit1 + 2) + iCaptionHeight
            lpRect.Bottom = (picBxV.Top + rSplit2) + iCaptionHeight
            Panel2Draw = 3
    End Select
    'restrict mouse to the panel drawn to
    Call GetClipCursor(rectOrg)
    Call ClipCursor(lpRect)

End Function
Private Function RectFrom2PointsF(Pt1 As POINTSINGLE, Pt2 As POINTSINGLE, rcReturn As RECT) As Boolean
    'Get bounding rect of 2 points
    If Pt1.x = Pt2.x Or Pt1.y = Pt2.y Then Exit Function  'points the same
    If Pt1.x > Pt2.x Then
        rcReturn.Right = Pt1.x
        rcReturn.Left = Pt2.x
    Else
        rcReturn.Right = Pt2.x
        rcReturn.Left = Pt1.x
    End If
    If Pt1.y > Pt2.y Then
        rcReturn.Top = Pt2.y
        rcReturn.Bottom = Pt1.y
    Else
        rcReturn.Top = Pt1.y
        rcReturn.Bottom = Pt2.y
    End If
    RectFrom2PointsF = True
End Function
Private Function SqrFrom2PointsF(Pt1 As POINTSINGLE, Pt2 As POINTSINGLE, rcReturnSQ As RECT) As Boolean
    'Get bounding square of 2 points
    If Pt1.x = Pt2.x Or Pt1.y = Pt2.y Then Exit Function 'points the same
    Dim rDiffX As Single, rDiffY As Single, iUpBnd As Long, iLoBnd As Long

    rDiffX = Abs(Pt1.x - Pt2.x)
    rDiffY = Abs(Pt1.y - Pt2.y)
    'make sure we don't draw out of bounds
    Select Case iPanel
        Case 1
            iUpBnd = 0
            iLoBnd = rSplit1
        Case 2
            iUpBnd = rSplit1
            iLoBnd = rSplit2
        Case 3
            iUpBnd = rSplit2
            iLoBnd = iBottomPlotMargin
    End Select
    If rDiffX > rDiffY Then  'change top-bott
        If Pt1.x > Pt2.x Then
            rcReturnSQ.Right = Pt1.x
            rcReturnSQ.Left = Pt2.x
        Else
            rcReturnSQ.Right = Pt2.x
            rcReturnSQ.Left = Pt1.x
        End If
        If Pt1.y > Pt2.y Then 'y1 is bottom
            rcReturnSQ.Top = Pt1.y - rDiffX
            If rcReturnSQ.Top < iUpBnd Then Exit Function
            rcReturnSQ.Bottom = Pt1.y
        Else 'y1 is top
            rcReturnSQ.Top = Pt1.y
            rcReturnSQ.Bottom = Pt1.y + rDiffX
            If rcReturnSQ.Bottom > iLoBnd Then Exit Function
        End If
    Else  'change right-left
        If Pt1.x > Pt2.x Then  'x1 is right
            rcReturnSQ.Right = Pt1.x
            rcReturnSQ.Left = Pt1.x - rDiffY
            If rcReturnSQ.Left < 0 Then Exit Function
        Else  'x1 is left
            rcReturnSQ.Right = Pt1.x + rDiffY
            If rcReturnSQ.Right > iMaxDrawRightX Then Exit Function
            rcReturnSQ.Left = Pt1.x
        End If
        If Pt1.y > Pt2.y Then
            rcReturnSQ.Top = Pt2.y
            rcReturnSQ.Bottom = Pt1.y
        Else
            rcReturnSQ.Top = Pt1.y
            rcReturnSQ.Bottom = Pt2.y
        End If
    End If

    SqrFrom2PointsF = True
End Function

Public Property Get ToolColor() As Long
    ToolColor = iToolColor
End Property

Public Property Let ToolColor(iToolColorA As Long)
    iToolColor = iToolColorA
End Property
Public Property Get ToolFillColor() As Long
    ToolFillColor = iToolFillColor
End Property

Public Property Let ToolFillColor(iToolFillColorA As Long)
    iToolFillColor = iToolFillColorA
End Property
Public Property Get ToolFillStyle() As Integer
    ToolFillStyle = iToolFillStyle
End Property

Public Property Let ToolFillStyle(iToolFillStyleA As Integer)
    iToolFillStyle = iToolFillStyleA
End Property
Public Property Get ToolStyle() As Long
    ToolStyle = iToolStyle
End Property

Public Property Let ToolStyle(iToolStyleA As Long)
    iToolStyle = iToolStyleA
End Property
Public Property Get ToolWidth() As Long
    ToolWidth = iToolWidth
End Property

Public Property Let ToolWidth(iToolWidthA As Long)
    iToolWidth = iToolWidthA
End Property
Public Property Get ToolMode() As Integer
    ToolMode = iToolMode
End Property

Public Property Let ToolMode(iToolModeA As Integer)
    iToolMode = iToolModeA
End Property
Public Property Get Extend() As Long
    Extend = iExtend
End Property

Public Property Let Extend(iExtendA As Long)
    iExtend = iExtendA
End Property

Public Property Get UseOrigin() As Boolean
    UseOrigin = fUseOrigin
End Property

Public Property Let UseOrigin(fUseOriginA As Boolean)
    fUseOrigin = fUseOriginA
End Property
Public Property Get picBx() As PictureBox
    Set picBx = objpicBx
End Property

Public Property Set picBx(objpicBxA As PictureBox)
    Set objpicBx = objpicBxA
End Property
Public Property Get picBxV() As PictureBox
    Set picBxV = objpicBxV
End Property

Public Property Set picBxV(objpicBxVA As PictureBox)
    Set objpicBxV = objpicBxVA
End Property

Public Property Get MouseClickEnabled() As Boolean
    MouseClickEnabled = fMouseClickEnabled
End Property

Public Property Let MouseClickEnabled(fMouseClickEnabledA As Boolean)
    fMouseClickEnabled = fMouseClickEnabledA
End Property

Public Property Get MouseClickNum() As Long
    MouseClickNum = iMouseClickNum
End Property

Public Property Let MouseClickNum(iMouseClickNumA As Long)
    iMouseClickNum = iMouseClickNumA
End Property
Public Property Get Xcurr() As Long
    Xcurr = iXcurr
End Property

Public Property Let Xcurr(iXcurrA As Long)
    iXcurr = iXcurrA
End Property

Public Property Get Ycurr() As Long
    Ycurr = iYcurr
End Property

Public Property Let Ycurr(iYcurrA As Long)
    iYcurr = iYcurrA
End Property


Public Property Get YcurrMov() As Long
    YcurrMov = iYcurrMov
End Property

Public Property Let YcurrMov(iYcurrMovA As Long)
    iYcurrMov = iYcurrMovA
End Property


Public Property Get XcurrMov() As Long
    XcurrMov = iXcurrMov
End Property

Public Property Let XcurrMov(iXcurrMovA As Long)
    iXcurrMov = iXcurrMovA
End Property


Private Sub Class_Initialize()
    ToolColor = vbRed
    ToolStyle = vbSolid
    ToolWidth = 1
    ToolMode = vbCopyPen
    ToolFillStyle = 1
    ToolFillColor = 10900825
    Extend = False
    fUseOrigin = True
    fMouseClickEnabled = False
    iMouseClickNum = 0
End Sub
