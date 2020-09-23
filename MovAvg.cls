VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MovAvg"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************
'Simple Moving Average. The data is added together for X periods, then divided by X.
'Common values include: 13, 21, 50 ,200
'*************************
Option Explicit

Private iAvgLen As Long
Private iAvgColor As Long
Private iAvgWidth As Long
Private rRetVal() As Double
Private iRetBoundHi As Integer
Private iRetBoundLo As Integer
Private iMax As Integer

Public Sub Calculate(aData() As Double)
Attribute Calculate.VB_Description = "aData() is [DaysRequested] X [0 to 7], (date,o,h,l,c,v,,)                                DataOpType:""O"",""H"",""L"",""C"",""V"",""SINGLE"" or ""S"""

    Dim x As Integer, y As Integer, iMAPeriodRange As Integer, rMAValue As Single
    
    
    iMax = UBound(aData()) + 1
    rMAValue = 0
    iMAPeriodRange = iMax - iAvgLen
    RetBoundHi = UBound(aData())
    RetBoundLo = LBound(aData())

    ReDim rRetVal(LBound(aData()) To UBound(aData()))
    
    For y = RetBoundLo To iMAPeriodRange
        DoEvents
        For x = y To y + iAvgLen - 1
            rMAValue = rMAValue + aData(x)
        Next x
        rRetVal(y + iAvgLen - 1) = Format((rMAValue / iAvgLen), "#.000")
        rMAValue = 0
    Next
    'RetBoundLo = LBound(rRetVal())
End Sub
Public Sub GetSavedSettings(Optional iNum As Integer = 0)
    Dim sNum As String
    sNum = CStr(iNum)
    If iNum = 0 Then sNum = sEmpty
    iAvgLen = Val(GetIni(sINIsetFile, "AvgSettings", "AvgLen" & sNum))
    iAvgColor = Val(GetIni(sINIsetFile, "AvgSettings", "AvgColor" & sNum))
    iAvgWidth = Val(GetIni(sINIsetFile, "AvgSettings", "AvgWidth" & sNum))
End Sub
Public Sub SaveCurrentSettings(Optional iNum As Integer = 0)
    Dim sNum As String
    sNum = CStr(iNum)
    If iNum = 0 Then sNum = sEmpty
    WriteIni sINIsetFile, "AvgSettings", "AvgLen & sNum", CStr(AvgLen)
    WriteIni sINIsetFile, "AvgSettings", "AvgColor & sNum", CStr(AvgColor)
    WriteIni sINIsetFile, "AvgSettings", "AvgWidth & sNum", CStr(AvgWidth)
End Sub
Public Property Get RetVal(i As Integer) As Double
Attribute RetVal.VB_Description = "Return data array of Moving average"
    RetVal = rRetVal(i)
End Property

Private Property Let RetVal(i As Integer, rRetValA As Double)
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

Public Property Get AvgLen() As Long
    AvgLen = iAvgLen
End Property

Public Property Let AvgLen(iAvgLenA As Long)
    iAvgLen = iAvgLenA
End Property

Public Property Get AvgColor() As Long
    AvgColor = iAvgColor
End Property

Public Property Let AvgColor(iAvgColorA As Long)
    iAvgColor = iAvgColorA
End Property

Public Property Get AvgWidth() As Long
    AvgWidth = iAvgWidth
End Property

Public Property Let AvgWidth(iAvgWidthA As Long)
    iAvgWidth = iAvgWidthA
End Property


