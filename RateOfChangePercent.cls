VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ROCPercent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************
'Rate of Change % indicator
'Calculates the difference between current price and price x periods ago divided
'by the price x periods ago... (Close-Close(periods ago))/(Close(periods ago)*100
'Commonly used with a period of 12 or 25 and trigger levels of +-6.5

Option Explicit

Private iROCColor As Long
Private iROCLen As Long
Private iPlotWidth As Long
Private rROCUpperTrig As Single
Private rROCLowerTrig As Single
Private rRetVal() As Double
Private iRetBoundHi As Integer
Private iRetBoundLo As Integer
Private iMax As Integer
Private fPlotTriggerLines As Boolean

Public Sub Calculate(aData() As Double)

    Dim x As Integer, y As Integer, iRocPeriodRange As Integer, rRocValue As Single
    
    
    iMax = UBound(aData())
    rRocValue = 0
    iRocPeriodRange = iMax - iROCLen
    RetBoundHi = UBound(aData())
    RetBoundLo = LBound(aData())

    ReDim rRetVal(LBound(aData()) To UBound(aData()))
    
    For y = iMax To RetBoundLo + iROCLen Step -1
        DoEvents
        rRocValue = (aData(y) - aData(y - iROCLen)) / aData(y - iROCLen) * 100
        rRetVal(y) = Format((rRocValue), "#.000")
    Next
End Sub
Public Sub GetSavedSettings(Optional iNum As Integer = 0)
    Dim sNum As String
    sNum = CStr(iNum)
    If iNum = 0 Then sNum = sEmpty
    iROCColor = Val(GetIni(sINIsetFile, "ROC%Settings", "ROC%Color" & sNum))
    iPlotWidth = Val(GetIni(sINIsetFile, "ROC%Settings", "PlotWidth" & sNum))
    iROCLen = Val(GetIni(sINIsetFile, "ROC%Settings", "ROC%Len" & sNum))
    rROCUpperTrig = Val(GetIni(sINIsetFile, "ROC%Settings", "ROC%UpperTrig" & sNum))
    rROCLowerTrig = Val(GetIni(sINIsetFile, "ROC%Settings", "ROC%LowerTrig" & sNum))
End Sub
Public Sub SaveCurrentSettings(Optional iNum As Integer = 0)
    Dim sNum As String
    sNum = CStr(iNum)
    If iNum = 0 Then sNum = sEmpty
    WriteIni sINIsetFile, "ROC%Settings", "ROC%Color" & sNum, CStr(iROCColor)
    WriteIni sINIsetFile, "ROC%Settings", "PlotWidth" & sNum, CStr(iPlotWidth)
    WriteIni sINIsetFile, "ROC%Settings", "ROC%Len" & sNum, CStr(iROCLen)
    WriteIni sINIsetFile, "ROC%Settings", "ROC%UpperTrig" & sNum, CStr(rROCUpperTrig)
    WriteIni sINIsetFile, "ROC%Settings", "ROC%LowerTrig" & sNum, CStr(rROCLowerTrig)
End Sub
Public Property Get RetVal(i As Integer) As Double
    RetVal = rRetVal(i)
End Property

Private Property Let RetVal(i As Integer, rRetValA As Double)
    rRetVal(i) = rRetValA
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

Public Property Get ROCColor() As Long
    ROCColor = iROCColor
End Property

Public Property Let ROCColor(iROCColorA As Long)
    iROCColor = iROCColorA
End Property

Public Property Get ROCLen() As Long
    ROCLen = iROCLen
End Property

Public Property Let ROCLen(iROCLenA As Long)
    iROCLen = iROCLenA
End Property

Public Property Get PlotWidth() As Long
    PlotWidth = iPlotWidth
End Property

Public Property Let PlotWidth(iPlotWidthA As Long)
    iPlotWidth = iPlotWidthA
End Property

Public Property Get UpperTrig() As Single
    UpperTrig = rROCUpperTrig
End Property

Public Property Let UpperTrig(rROCUpperTrigA As Single)
    rROCUpperTrig = rROCUpperTrigA
End Property

Public Property Get LowerTrig() As Single
    LowerTrig = rROCLowerTrig
End Property

Public Property Let LowerTrig(rROCLowerTrigA As Single)
    rROCLowerTrig = rROCLowerTrigA
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
