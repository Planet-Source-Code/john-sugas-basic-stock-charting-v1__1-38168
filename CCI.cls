VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCI"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************
'CCI  AKA Commodity Channel Index
'This indicator measures the variation of price from it's
'statistical mean.
'***************************
Option Explicit



Private iCciColor As Long
Private iCciAvgAvgColor As Long
Private iCciAvgColor As Long
Private iCciLen As Integer
Private iCciAvgLen As Integer
Private iCciAvgAvgLen As Integer
Private iPlotWidth As Long
Private iCciUpperTrig As Long
Private iCciLowerTrig As Long
Private rRetVal() As Double
Private rRetValAvgAvg() As Double
Private rRetValAvg() As Double
Private iRetBoundHi As Integer
Private iRetBoundLo As Integer
Private iMax As Integer
Private fPlotTriggerLines As Boolean


Public Sub Calculate(aData() As StockData)

    Dim x As Integer, y As Integer, i As Integer
    Dim iPeriodRange As Integer, rValue As Single, rAvg As Single, rSum As Single, rPrice As Single
    
    
    iMax = UBound(aData()) + 1
    rValue = 0
    iPeriodRange = iMax - iCciLen
    RetBoundHi = UBound(aData())
    RetBoundLo = LBound(aData())

    ReDim rRetVal(RetBoundLo To RetBoundHi)
    'cci calc
    For y = LBound(aData()) To iPeriodRange
        DoEvents
        i = y + iCciLen - 1
        For x = y To i
            rPrice = (aData(x).dHigh + aData(x).dLow + aData(x).dClose) / (3 * iCciLen)
            rAvg = rAvg + rPrice
        Next
        For x = y To i
            rPrice = (aData(x).dHigh + aData(x).dLow + aData(x).dClose) / 3
            rSum = rSum + Abs(rPrice - rAvg)
        Next
        x = i
        rSum = rSum / iCciLen
        rPrice = (aData(x).dHigh + aData(x).dLow + aData(x).dClose) / 3
        rValue = (rPrice - rAvg) / (0.015 * rSum)
        RetVal(i) = Format((rValue), "#.000")
        rValue = 0
        rAvg = 0
        rSum = 0
    Next
    'avg1 calc
    Dim MovAvgs As MovAvg
    Set MovAvgs = New MovAvg
    MovAvgs.AvgLen = iCciAvgLen
    MovAvgs.Calculate rRetVal()
    ReDim rRetValAvg(MovAvgs.RetBoundLo To MovAvgs.RetBoundHi)
    For i = MovAvgs.RetBoundLo To MovAvgs.RetBoundHi
        rRetValAvg(i) = MovAvgs.RetVal(i)
    Next
    'avg of avg
    MovAvgs.AvgLen = iCciAvgAvgLen
    MovAvgs.Calculate rRetValAvg()
    ReDim rRetValAvgAvg(MovAvgs.RetBoundLo To MovAvgs.RetBoundHi)
    For i = MovAvgs.RetBoundLo To MovAvgs.RetBoundHi
        rRetValAvgAvg(i) = MovAvgs.RetVal(i)
    Next
    
    Set MovAvgs = Nothing
End Sub
Public Sub GetSavedSettings(Optional iNum As Integer = 0)
    Dim sNum As String
    sNum = CStr(iNum)
    If iNum = 0 Then sNum = sEmpty
    iCciColor = Val(GetIni(sINIsetFile, "CciSettings", "CciColor" & sNum))
    iCciAvgColor = Val(GetIni(sINIsetFile, "CciSettings", "CciAvgColor" & sNum))
    iCciLen = Val(GetIni(sINIsetFile, "CciSettings", "CciLen" & sNum))
    iCciAvgLen = Val(GetIni(sINIsetFile, "CciSettings", "CciAvgLen" & sNum))
    iCciAvgAvgLen = Val(GetIni(sINIsetFile, "CciSettings", "CciAvgAvgLen" & sNum))
    iPlotWidth = Val(GetIni(sINIsetFile, "CciSettings", "PlotWidth" & sNum))
    iCciUpperTrig = Val(GetIni(sINIsetFile, "CciSettings", "CciUpperTrig" & sNum))
    iCciLowerTrig = Val(GetIni(sINIsetFile, "CciSettings", "CciLowerTrig" & sNum))
End Sub
Public Sub SaveCurrentSettings(Optional iNum As Integer = 0)
    Dim sNum As String
    sNum = CStr(iNum)
    If iNum = 0 Then sNum = sEmpty
    WriteIni sINIsetFile, "CciSettings", "CciColor" & sNum, CStr(iCciColor)
    WriteIni sINIsetFile, "CciSettings", "CciAvgColor" & sNum, CStr(iCciAvgColor)
    WriteIni sINIsetFile, "CciSettings", "CciLen" & sNum, CStr(iCciLen)
    WriteIni sINIsetFile, "CciSettings", "CciAvgLen" & sNum, CStr(iCciAvgLen)
    WriteIni sINIsetFile, "CciSettings", "CciAvgAvgLen" & sNum, CStr(iCciAvgAvgLen)
    WriteIni sINIsetFile, "CciSettings", "PlotWidth" & sNum, CStr(iPlotWidth)
    WriteIni sINIsetFile, "CciSettings", "CciUpperTrig" & sNum, CStr(iCciUpperTrig)
    WriteIni sINIsetFile, "CciSettings", "CciLowerTrig" & sNum, CStr(iCciLowerTrig)
End Sub
Public Property Get RetVal(i As Integer) As Double
    If i >= 0 Then RetVal = rRetVal(i)
End Property

Private Property Let RetVal(i As Integer, rRetValA As Double)
    If i >= 0 Then rRetVal(i) = rRetValA
End Property


Public Property Get RetValAvgAvg(i As Integer) As Double
    If i >= 0 Then RetValAvgAvg = rRetValAvgAvg(i)
End Property

Public Property Let RetValAvgAvg(i As Integer, rRetValAvgAvgA As Double)
    If i >= 0 Then rRetValAvgAvg(i) = rRetValAvgAvgA
End Property

Public Property Get RetValAvg(i As Integer) As Double
    If i >= 0 Then RetValAvg = rRetValAvg(i)
End Property

Public Property Let RetValAvg(i As Integer, rRetValAvgA As Double)
    If i >= 0 Then rRetValAvg(i) = rRetValAvgA
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

Public Property Get CciColor() As Long
    CciColor = iCciColor
End Property

Public Property Let CciColor(iCciColorA As Long)
    iCciColor = iCciColorA
End Property

Public Property Get CciAvgAvgColor() As Long
    CciAvgAvgColor = iCciAvgAvgColor
End Property

Public Property Let CciAvgAvgColor(iCciAvgAvgColorA As Long)
    iCciAvgAvgColor = iCciAvgAvgColorA
End Property

Public Property Get CciAvgColor() As Long
    CciAvgColor = iCciAvgColor
End Property

Public Property Let CciAvgColor(iCciAvgColorA As Long)
    iCciAvgColor = iCciAvgColorA
End Property

Public Property Get CciLen() As Integer
    CciLen = iCciLen
End Property

Public Property Let CciLen(iCciLenA As Integer)
    iCciLen = iCciLenA
End Property

Public Property Get CciAvgLen() As Integer
    CciAvgLen = iCciAvgLen
End Property

Public Property Let CciAvgLen(iCciAvgLenA As Integer)
    iCciAvgLen = iCciAvgLenA
End Property

Public Property Get CciAvgAvgLen() As Integer
    CciAvgAvgLen = iCciAvgAvgLen
End Property

Public Property Let CciAvgAvgLen(iCciAvgAvgLenA As Integer)
    iCciAvgAvgLen = iCciAvgAvgLenA
End Property

Public Property Get PlotWidth() As Long
    PlotWidth = iPlotWidth
End Property

Public Property Let PlotWidth(iPlotWidthA As Long)
    iPlotWidth = iPlotWidthA
End Property

Public Property Get UpperTrig() As Long
    UpperTrig = iCciUpperTrig
End Property

Public Property Let UpperTrig(iCciUpperTrigA As Long)
    iCciUpperTrig = iCciUpperTrigA
End Property

Public Property Get LowerTrig() As Long
    LowerTrig = iCciLowerTrig
End Property

Public Property Let LowerTrig(iCciLowerTrigA As Long)
    iCciLowerTrig = iCciLowerTrigA
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
