VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stoch"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************
'Stochastic. Compares last close with it's price range over a certain length of time.
'Calculated: (Last close- lowest low in X time periods)/
'            (highest hi in X time - lowest low in X time) * 100
'******************************

Option Explicit

Private iStochKColor As Long
Private iStochKSlowColor As Long
Private iStochDColor As Long
Private iPlotWidth As Long
Private rStochUpperTrig As Single
Private rStochLowerTrig As Single
Private iKPeriod As Integer
Private iKPeriodSlow As Integer
Private iDPeriod As Integer
Private rRetValK() As Double
Private rRetValD() As Double
Private rRetValKSlow() As Double
Private iRetBoundHi As Integer
Private iRetBoundLo As Integer
Private iMax As Integer
Private fPlotTriggerLines As Boolean
Private fStochSlow As Boolean

Private Sub Class_Initialize()
    iKPeriodSlow = 3  'internal smoothing value
    fPlotTriggerLines = True
End Sub
Public Sub Calculate(aData() As StockData)

    Dim x As Integer, y As Integer, i As Integer
    Dim iPeriodRange As Integer, rAvg As Single, rSum As Single, rPrice As Single
    Dim rValue1 As Single, rValue2 As Single, rValue3 As Single
    Dim MovAvg1 As New MovAvg, aTemp() As Double
    
    iMax = UBound(aData()) + 1
    iPeriodRange = iMax - KPeriod()
    RetBoundHi = UBound(aData())
    RetBoundLo = LBound(aData())

    ReDim rRetValK(RetBoundLo To RetBoundHi)
    ReDim rRetValKSlow(RetBoundLo To RetBoundHi)
    ReDim rRetValD(RetBoundLo To RetBoundHi)
    DoEvents
    For y = RetBoundLo To iPeriodRange
        'fast K
        rValue1 = 999999
        rValue2 = 0
        i = y + KPeriod() - 1
        For x = y To i
            If rValue1 > aData(x).dLow Then rValue1 = aData(x).dLow
            If rValue2 < aData(x).dHigh Then rValue2 = aData(x).dHigh
            rValue3 = aData(x).dClose
        Next x
        rValue2 = rValue2 - rValue1
        If rValue2 > 0 Then
            rPrice = (rValue3 - rValue1) / rValue2 * 100
            RetValK(i) = Format((rPrice), "#.000")
        Else
            RetValK(i) = Format(0, "#.000")
        End If
        
    Next
    ReDim aTemp(RetBoundLo To RetBoundHi)
    For i = RetBoundLo To RetBoundHi
        aTemp(i) = RetValK(i)
    Next
    'slowK
    MovAvg1.AvgLen = iKPeriodSlow
    MovAvg1.Calculate aTemp()
    For i = RetBoundLo To RetBoundHi
        rRetValKSlow(i) = Format(MovAvg1.RetVal(i), "#.000")
    Next
    'get fastD then slowD
    MovAvg1.AvgLen = DPeriod
    MovAvg1.Calculate aTemp()
    For i = RetBoundLo To RetBoundHi
        aTemp(i) = MovAvg1.RetVal(i)
    Next
    For i = RetBoundLo + 1 To RetBoundHi
        rRetValD(i) = Format(((aTemp(i) + rRetValD(i - 1) * 2) / 3), "#.000")
    Next
    
End Sub
Public Sub GetSavedSettings(Optional iNum As Integer = 0)
    Dim sNum As String
    sNum = CStr(iNum)
    If iNum = 0 Then sNum = sEmpty
    iStochKColor = Val(GetIni(sINIsetFile, "StochSettings", "StochKColor" & sNum))
    iStochKSlowColor = Val(GetIni(sINIsetFile, "StochSettings", "StochKSlowColor" & sNum))
    iStochDColor = Val(GetIni(sINIsetFile, "StochSettings", "StochDColor" & sNum))
    iPlotWidth = Val(GetIni(sINIsetFile, "StochSettings", "PlotWidth" & sNum))
    rStochUpperTrig = Val(GetIni(sINIsetFile, "StochSettings", "StochUpperTrig" & sNum))
    rStochLowerTrig = Val(GetIni(sINIsetFile, "StochSettings", "StochLowerTrig" & sNum))
    iKPeriod = Val(GetIni(sINIsetFile, "StochSettings", "KPeriod" & sNum))
    iKPeriodSlow = Val(GetIni(sINIsetFile, "StochSettings", "KPeriodSlow" & sNum))
    iDPeriod = Val(GetIni(sINIsetFile, "StochSettings", "DPeriod" & sNum))
    fStochSlow = GetIni(sINIsetFile, "StochSettings", "StochSlow" & sNum)
End Sub
Public Sub SaveCurrentSettings(Optional iNum As Integer = 0)
    Dim sNum As String
    sNum = CStr(iNum)
    If iNum = 0 Then sNum = sEmpty
    WriteIni sINIsetFile, "StochSettings", "StochKColor" & sNum, CStr(iStochKColor)
    WriteIni sINIsetFile, "StochSettings", "StochKSlowColor" & sNum, CStr(iStochKSlowColor)
    WriteIni sINIsetFile, "StochSettings", "StochDColor" & sNum, CStr(iStochDColor)
    WriteIni sINIsetFile, "StochSettings", "PlotWidth" & sNum, CStr(iPlotWidth)
    WriteIni sINIsetFile, "StochSettings", "KPeriod" & sNum, CStr(iKPeriod)
    WriteIni sINIsetFile, "StochSettings", "KPeriodSlow" & sNum, CStr(iKPeriodSlow)
    WriteIni sINIsetFile, "StochSettings", "DPeriod" & sNum, CStr(iDPeriod)
    WriteIni sINIsetFile, "StochSettings", "StochUpperTrig" & sNum, CStr(rStochUpperTrig)
    WriteIni sINIsetFile, "StochSettings", "StochLowerTrig" & sNum, CStr(rStochLowerTrig)
    WriteIni sINIsetFile, "StochSettings", "StochSlow" & sNum, (fStochSlow)
End Sub
Public Property Get RetValK(i As Integer) As Single
    If i >= 0 Then RetValK = rRetValK(i)
End Property

Private Property Let RetValK(i As Integer, rRetValKA As Single)
    If i >= 0 Then rRetValK(i) = rRetValKA
End Property

Public Property Get RetValD(i As Integer) As Single
    If i >= 0 Then RetValD = rRetValD(i)
End Property

Private Property Let RetValD(i As Integer, rRetValDA As Single)
    If i >= 0 Then rRetValD(i) = rRetValDA
End Property

Public Property Get RetValKSlow(i As Integer) As Single
    If i >= 0 Then RetValKSlow = rRetValKSlow(i)
End Property

Private Property Let RetValKSlow(i As Integer, rRetValKSlowA As Single)
    If i >= 0 Then rRetValKSlow(i) = rRetValKSlowA
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

Public Property Get KPeriod() As Integer
    KPeriod = iKPeriod
End Property

Public Property Let KPeriod(iKPeriodA As Integer)
    iKPeriod = iKPeriodA
End Property

Public Property Get DPeriod() As Integer
    DPeriod = iDPeriod
End Property

Public Property Let DPeriod(iDPeriodA As Integer)
    iDPeriod = iDPeriodA
End Property

Public Property Get StochKColor() As Long
    StochKColor = iStochKColor
End Property

Public Property Let StochKColor(iStochKColorA As Long)
    iStochKColor = iStochKColorA
End Property

Public Property Get StochKSlowColor() As Long
    StochKSlowColor = iStochKSlowColor
End Property

Public Property Let StochKSlowColor(iStochKSlowColorA As Long)
    iStochKSlowColor = iStochKSlowColorA
End Property

Public Property Get StochDColor() As Long
    StochDColor = iStochDColor
End Property

Public Property Let StochDColor(iStochDColorA As Long)
    iStochDColor = iStochDColorA
End Property

Public Property Get PlotWidth() As Long
    PlotWidth = iPlotWidth
End Property

Public Property Let PlotWidth(iPlotWidthA As Long)
    iPlotWidth = iPlotWidthA
End Property

Public Property Get UpperTrig() As Single
    UpperTrig = rStochUpperTrig
End Property

Public Property Let UpperTrig(rStochUpperTrigA As Single)
    rStochUpperTrig = rStochUpperTrigA
End Property

Public Property Get LowerTrig() As Single
    LowerTrig = rStochLowerTrig
End Property

Public Property Let LowerTrig(rStochLowerTrigA As Single)
    rStochLowerTrig = rStochLowerTrigA
End Property

Public Property Get PlotTriggerLines() As Boolean
    PlotTriggerLines = fPlotTriggerLines
End Property

Public Property Let PlotTriggerLines(fPlotTriggerLinesA As Boolean)
    fPlotTriggerLines = fPlotTriggerLinesA
End Property
Public Property Get ZeroLinePlot() As Boolean
    ZeroLinePlot = False
End Property

Public Property Get StochSlow() As Boolean
    StochSlow = fStochSlow
End Property

Public Property Let StochSlow(fStochSlowA As Boolean)
    fStochSlow = fStochSlowA
End Property
