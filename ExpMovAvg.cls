VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExpMovAvg"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************
'Exponential moving averages put more "weight" on recent price. A percentage
'of current close is added to yesterdays MA value. Tends to be more responsive
'than simple moving averages
'****************************
Option Explicit

Private iExpAvgLen As Long
Private iExpAvgColor As Long
Private iPlotWidth As Long
Private rRetVal() As Double
Private iRetBoundHi As Integer
Private iRetBoundLo As Integer

Public Sub Calculate(aData() As Double)
Attribute Calculate.VB_Description = "aData() is [DaysRequested] X [0 to 7], (date,o,h,l,c,v,,)                                DataOpType:""O"",""H"",""L"",""C"",""V"",""SINGLE"" or ""S"""

    Dim x As Integer, y As Integer

    RetBoundLo = LBound(aData())
    RetBoundHi = UBound(aData())
    ReDim rRetVal(RetBoundLo To RetBoundHi)

    DoEvents
    For x = LBound(aData()) To RetBoundHi
        If x > LBound(aData()) Then
            RetVal(x) = (aData(x) * ExpPercent) + _
                    (RetVal(x - 1) * ExpPerYesterday)
        Else
            RetVal(x) = aData(x)
        End If
    Next
    
End Sub
Public Sub GetSavedSettings(Optional iNum As Integer = 0)
    Dim sNum As String
    sNum = CStr(iNum)
    If iNum = 0 Then sNum = sEmpty
    iExpAvgLen = Val(GetIni(sINIsetFile, "ExpAvgSettings", "ExpAvgLen" & sNum))
    iExpAvgColor = Val(GetIni(sINIsetFile, "ExpAvgSettings", "ExpAvgColor" & sNum))
    iPlotWidth = Val(GetIni(sINIsetFile, "ExpAvgSettings", "PlotWidth" & sNum))
End Sub
Public Sub SaveCurrentSettings(Optional iNum As Integer = 0)
    Dim sNum As String
    sNum = CStr(iNum)
    If iNum = 0 Then sNum = sEmpty
    WriteIni sINIsetFile, "ExpAvgSettings", "ExpAvgLen & sNum", CStr(ExpAvgLen)
    WriteIni sINIsetFile, "ExpAvgSettings", "ExpAvgColor & sNum", CStr(ExpAvgColor)
    WriteIni sINIsetFile, "ExpAvgSettings", "PlotWidth & sNum", CStr(PlotWidth)
End Sub
Private Function ExpPerYesterday() As Single
    'ExpPerYesterday = Format(1 - ExpPercent, "0.0000")
    ExpPerYesterday = 1 - ExpPercent
End Function

Private Function ExpPercent() As Single
    'ExpPercent = Format(2 / (IndicatorPeriods + 1), "0.0000")
    ExpPercent = 2 / (iExpAvgLen + 1)
End Function

Public Property Get RetVal(i As Integer) As Double
Attribute RetVal.VB_Description = "Return data array of Moving average"
    'If i >= 0 Then RetVal = rRetVal(i)
    RetVal = rRetVal(i)
End Property

Private Property Let RetVal(i As Integer, rRetValA As Double)
    'If i >= 0 Then rRetVal(i) = rRetValA
    rRetVal(i) = rRetValA
End Property

Public Property Get RetBoundHi() As Integer
Attribute RetBoundHi.VB_Description = "Upper bound of return data array"
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

Public Property Get ExpAvgLen() As Long
    ExpAvgLen = iExpAvgLen
End Property

Public Property Let ExpAvgLen(iExpAvgLenA As Long)
    iExpAvgLen = iExpAvgLenA
End Property

Public Property Get ExpAvgColor() As Long
    ExpAvgColor = iExpAvgColor
End Property

Public Property Let ExpAvgColor(iExpAvgColorA As Long)
    iExpAvgColor = iExpAvgColorA
End Property

Public Property Get PlotWidth() As Long
    PlotWidth = iPlotWidth
End Property

Public Property Let PlotWidth(iPlotWidthA As Long)
    iPlotWidth = iPlotWidthA
End Property

