VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RSI"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************
'RSI - Relative Strength Index. Measures the "internal" strength of an equity.
'It is calculated by subtracting from 100 the value of 100 divided by
'1 + an average of up days divided by an average of down days.
'*******************************
Option Explicit

Private iRsiColor As Long
Private iRsiLen As Long
Private iPlotWidth As Long
Private rRsiUpperTrig As Single
Private rRsiLowerTrig As Single
Private rRetVal() As Single
Private iRetBoundHi As Integer
Private iRetBoundLo As Integer
Private iMax As Integer
Private fPlotTriggerLines As Boolean

Public Sub Calculate(aData() As Double)

    Dim x As Integer, y As Integer, i As Integer
    Dim iRsiLenRange As Integer, iStartPosition As Integer
    Dim dAvgUp As Double, dAvgDn As Double, dPrevDn As Double, dPrevUp As Double
    Dim dChg As Double, dRSi As Double, dRSiUp As Double, dRSiDn As Double
    
    iMax = UBound(aData()) + 1

    iRsiLenRange = iMax - iRsiLen
    RetBoundHi = UBound(aData())
    RetBoundLo = LBound(aData())
    iStartPosition = RetBoundLo + iRsiLen - 1
    
    ReDim rRetVal(RetBoundLo To RetBoundHi)
    DoEvents
    For i = iStartPosition To RetBoundHi
        If i = iStartPosition Then
            For y = i To RetBoundLo + 1 Step -1
                dChg = aData(y) - aData(y - 1)
                If dChg >= 0 Then
                    dRSiUp = dRSiUp + dChg
                ElseIf dChg < 0 Then
                    dRSiDn = dRSiDn + Abs(dChg)
                End If
            Next
            dAvgUp = dRSiUp / iRsiLen
            dAvgDn = dRSiDn / iRsiLen
        Else
            
            dChg = aData(i) - aData(i - 1)
            If dChg >= 0 Then
                dRSiUp = dRSiUp + dChg
            ElseIf dChg < 0 Then
                dRSiDn = dRSiDn + Abs(dChg)
            End If
            dAvgUp = (dPrevUp * (iRsiLen - 1) + dRSiUp) / iRsiLen
            dAvgDn = (dPrevDn * (iRsiLen - 1) + dRSiDn) / iRsiLen
        
        End If
        
        If dRSiUp + dRSiDn <> 0 Then
            dRSi = 100 - (100 / (1 + (dAvgUp / dAvgDn)))
        Else 'if zero use prev. value
            dRSi = RetVal(i - 1)
        End If
            
        RetVal(i) = Format((dRSi), "##.000")

        dPrevUp = dAvgUp
        dPrevDn = dAvgDn
        dRSiUp = 0
        dRSiDn = 0
    Next
            
End Sub
Public Sub GetSavedSettings(Optional iNum As Integer = 0)
    Dim sNum As String
    sNum = CStr(iNum)
    If iNum = 0 Then sNum = sEmpty
    iRsiColor = Val(GetIni(sINIsetFile, "RsiSettings", "RsiColor" & sNum))
    iPlotWidth = Val(GetIni(sINIsetFile, "RsiSettings", "PlotWidth" & sNum))
    iRsiLen = Val(GetIni(sINIsetFile, "RsiSettings", "RsiLen" & sNum))
    rRsiUpperTrig = Val(GetIni(sINIsetFile, "RsiSettings", "RsiUpperTrig" & sNum))
    rRsiLowerTrig = Val(GetIni(sINIsetFile, "RsiSettings", "RsiLowerTrig" & sNum))
End Sub
Public Sub SaveCurrentSettings(Optional iNum As Integer = 0)
    Dim sNum As String
    sNum = CStr(iNum)
    If iNum = 0 Then sNum = sEmpty
    WriteIni sINIsetFile, "RsiSettings", "RsiColor" & sNum, CStr(iRsiColor)
    WriteIni sINIsetFile, "RsiSettings", "PlotWidth" & sNum, CStr(iPlotWidth)
    WriteIni sINIsetFile, "RsiSettings", "RsiLen" & sNum, CStr(iRsiLen)
    WriteIni sINIsetFile, "RsiSettings", "RsiUpperTrig" & sNum, CStr(rRsiUpperTrig)
    WriteIni sINIsetFile, "RsiSettings", "RsiLowerTrig" & sNum, CStr(rRsiLowerTrig)
End Sub

Public Property Get RetVal(i As Integer) As Single
    If i >= 0 Then RetVal = rRetVal(i)
End Property

Private Property Let RetVal(i As Integer, rRetValA As Single)
    If i >= 0 Then rRetVal(i) = rRetValA
End Property

Public Property Get RetBoundHi() As Integer
        RetBoundHi = iRetBoundHi
End Property

Private Property Let RetBoundHi(iRetBoundHiA As Integer)
    iRetBoundHi = iRetBoundHiA
End Property
Public Property Get RetBoundLo() As Integer
        RetBoundLo = iRetBoundLo
End Property

Private Property Let RetBoundLo(iRetBoundLoA As Integer)
    iRetBoundLo = iRetBoundLoA
End Property

Public Property Get RsiColor() As Long
    RsiColor = iRsiColor
End Property

Public Property Let RsiColor(iRsiColorA As Long)
    iRsiColor = iRsiColorA
End Property

Public Property Get RsiLen() As Long
    RsiLen = iRsiLen
End Property

Public Property Let RsiLen(iRsiLenA As Long)
    iRsiLen = iRsiLenA
End Property

Public Property Get PlotWidth() As Long
    PlotWidth = iPlotWidth
End Property

Public Property Let PlotWidth(iPlotWidthA As Long)
    iPlotWidth = iPlotWidthA
End Property

Public Property Get UpperTrig() As Single
    UpperTrig = rRsiUpperTrig
End Property

Public Property Let UpperTrig(rRsiUpperTrigA As Single)
    rRsiUpperTrig = rRsiUpperTrigA
End Property

Public Property Get LowerTrig() As Single
    LowerTrig = rRsiLowerTrig
End Property

Public Property Let LowerTrig(rRsiLowerTrigA As Single)
    rRsiLowerTrig = rRsiLowerTrigA
End Property

Public Property Get PlotTriggerLines() As Boolean
    PlotTriggerLines = fPlotTriggerLines
End Property

Public Property Let PlotTriggerLines(fPlotTriggerLinesA As Boolean)
    fPlotTriggerLines = fPlotTriggerLinesA
End Property
Public Property Get ZeroLinePlot() As Boolean
    ZeroLinePlot = True
End Property

Private Sub Class_Initialize()
    fPlotTriggerLines = True
End Sub
