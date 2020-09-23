VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MACD"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************
'MACD... Moving Average Convergence-Divergence
'This indicator is the difference in value between two different
'period exponential moving averages. The signal line is an average of
'the first value series. The histogram is the difference between the
'first difference data and the signal data.
'*****************************
Option Explicit

Private iMACDColor1 As Long
Private iMACDColor2 As Long
Private iMACDColor3 As Long
Private iPlotWidth As Long
Private iMACDLen1 As Integer
Private iMACDLen2 As Integer
Private iMACDLen3 As Integer
Private rRetVal() As Double
Private rRetValSig() As Double
Private rRetValHist() As Double
Private iRetBoundHi As Integer
Private iRetBoundLo As Integer
Private fPlotTriggerLines As Boolean
Private aTemp() As Double


Public Sub Calculate(aData() As Double)
Attribute Calculate.VB_Description = "aData() is [DaysRequested] X [0 to 7], (date,o,h,l,c,v,,)                                DataOpType:""O"",""H"",""L"",""C"",""V"",""SINGLE"" or ""S"""

    Dim x As Integer, y As Integer, eMov1 As ExpMovAvg, eMov2 As ExpMovAvg
    Dim eMovSig As ExpMovAvg

    Set eMov1 = New ExpMovAvg
    eMov1.ExpAvgLen = iMACDLen1
    eMov1.Calculate aData()

    Set eMov2 = New ExpMovAvg
    eMov2.ExpAvgLen = iMACDLen2
    eMov2.Calculate aData()

    Set eMovSig = New ExpMovAvg
    eMovSig.ExpAvgLen = iMACDLen3
    
    RetBoundHi() = UBound(aData())
    RetBoundLo() = LBound(aData())
    
    ReDim rRetVal(RetBoundLo() To RetBoundHi())
    ReDim rRetValSig(RetBoundLo() To RetBoundHi())
    ReDim rRetValHist(RetBoundLo() To RetBoundHi())
    ReDim aTemp(RetBoundLo() To RetBoundHi())
    
    DoEvents
    For x = RetBoundLo() To RetBoundHi()
        aTemp(x) = eMov1.RetVal(x) - eMov2.RetVal(x)
        RetVal(x) = aTemp(x)
    Next
    
    eMovSig.Calculate aTemp()

    For x = RetBoundLo() To RetBoundHi()
        RetValSig(x) = eMovSig.RetVal(x)
        RetValHist(x) = RetVal(x) - eMovSig.RetVal(x)
    Next
    
    Set eMovSig = Nothing
    Set eMov2 = Nothing
    Set eMov1 = Nothing
    
End Sub
Public Sub GetSavedSettings(Optional iNum As Integer = 0)
    Dim sNum As String
    sNum = CStr(iNum)
    If iNum = 0 Then sNum = sEmpty
    iMACDColor1 = Val(GetIni(sINIsetFile, "MACDSettings", "MACDColor1" & sNum))
    iMACDColor2 = Val(GetIni(sINIsetFile, "MACDSettings", "MACDColor2" & sNum))
    iMACDColor3 = Val(GetIni(sINIsetFile, "MACDSettings", "MACDColor3" & sNum))
    iPlotWidth = Val(GetIni(sINIsetFile, "MACDSettings", "PlotWidth" & sNum))
    iMACDLen1 = Val(GetIni(sINIsetFile, "MACDSettings", "MACDLen1" & sNum))
    iMACDLen2 = Val(GetIni(sINIsetFile, "MACDSettings", "MACDLen2" & sNum))
    iMACDLen3 = Val(GetIni(sINIsetFile, "MACDSettings", "MACDLen3" & sNum))
End Sub
Public Sub SaveCurrentSettings(Optional iNum As Integer = 0)
    Dim sNum As String
    sNum = CStr(iNum)
    If iNum = 0 Then sNum = sEmpty
    WriteIni sINIsetFile, "MACDSettings", "MACDColor1" & sNum, CStr(iMACDColor1)
    WriteIni sINIsetFile, "MACDSettings", "MACDColor2" & sNum, CStr(iMACDColor2)
    WriteIni sINIsetFile, "MACDSettings", "MACDColor3" & sNum, CStr(iMACDColor3)
    WriteIni sINIsetFile, "MACDSettings", "PlotWidth" & sNum, CStr(iPlotWidth)
    WriteIni sINIsetFile, "MACDSettings", "MACDLen1" & sNum, CStr(iMACDLen1)
    WriteIni sINIsetFile, "MACDSettings", "MACDLen2" & sNum, CStr(iMACDLen2)
    WriteIni sINIsetFile, "MACDSettings", "MACDLen3" & sNum, CStr(iMACDLen3)
End Sub

Public Property Get RetVal(i As Integer) As Double
    RetVal = rRetVal(i)
End Property

Private Property Let RetVal(i As Integer, rRetValA As Double)
    rRetVal(i) = rRetValA
End Property
Public Property Get RetValSig(i As Integer) As Double
    RetValSig = rRetValSig(i)
End Property

Private Property Let RetValSig(i As Integer, rRetValSigA As Double)
    rRetValSig(i) = rRetValSigA
End Property
Public Property Get RetValHist(i As Integer) As Double
    RetValHist = rRetValHist(i)
End Property

Private Property Let RetValHist(i As Integer, rRetValHistA As Double)
    rRetValHist(i) = rRetValHistA
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

Public Property Get MACDColor1() As Long
    MACDColor1 = iMACDColor1
End Property

Public Property Let MACDColor1(iMACDColor1A As Long)
    iMACDColor1 = iMACDColor1A
End Property

Public Property Get MACDColor2() As Long
    MACDColor2 = iMACDColor2
End Property

Public Property Let MACDColor2(iMACDColor2A As Long)
    iMACDColor2 = iMACDColor2A
End Property

Public Property Get MACDColor3() As Long
    MACDColor3 = iMACDColor3
End Property

Public Property Let MACDColor3(iMACDColor3A As Long)
    iMACDColor3 = iMACDColor3A
End Property

Public Property Get PlotWidth() As Long
    PlotWidth = iPlotWidth
End Property

Public Property Let PlotWidth(iPlotWidthA As Long)
    iPlotWidth = iPlotWidthA
End Property

Public Property Get MACDLen1() As Integer
    MACDLen1 = iMACDLen1
End Property

Public Property Let MACDLen1(iMACDLen1A As Integer)
    iMACDLen1 = iMACDLen1A
End Property

Public Property Get MACDLen2() As Integer
    MACDLen2 = iMACDLen2
End Property

Public Property Let MACDLen2(iMACDLen2A As Integer)
    iMACDLen2 = iMACDLen2A
End Property

Public Property Get MACDLen3() As Integer
    MACDLen3 = iMACDLen3
End Property

Public Property Let MACDLen3(iMACDLen3A As Integer)
    iMACDLen3 = iMACDLen3A
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
    fPlotTriggerLines = False
End Sub
