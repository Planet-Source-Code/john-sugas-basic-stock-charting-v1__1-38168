Attribute VB_Name = "MPublicVarsFuncs"
Option Explicit



'can't put a floating point number in the type lib
Public Const PI As Double = 3.141592654



'Logfont structure is a problem in the type lib. Fixed strings have to be
'declared as byte arrays. Then the strings have to be fished out of the blobs later
'So here it is... as well as the CreatFontIndirsect which uses the LogFont struct
'as an argument.... (to use it in the TL the struct has to be found before the declare)
'Note: I tried to substute the LogFont with "as Any" but it didn't work properly...

'types
Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 32  'type lib headaches with fixed strings
'    lfFaceName(32) As Byte
End Type

Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long


'Public vars
Public iCrossHairMode As Long  'crosshair drawmode
Public iCrossHairColor As Long
Public iMouseLabelColor As Long
Public iBackColor As Long
Public iGridColor As Long
Public iForeColor As Long
Public iTicType As Long   'type of chart drawn
Public iTicBodyColor As Long
Public iTicOpenColor As Long
Public iTicCloseColor As Long
Public iTicCandleUpColor As Long
Public iTicCandleDnColor As Long
Public iVolColor As Long
Public sFontName As String
Public iFontSize As Long
Public iFontBold As Integer
Public iFontItalic As Integer
Public sDir As String
Public sFileName As String
Public sFilePath As String
Public sDataDir As String
Public sINIsetFile As String
Public sSymbol As String
Public aData() As StockData  'price data array
Public iUBaData As Long   'aData ubound
Public iLboundDataStart As Long  'plotted LBound
Public iStartIndex As Long   'scrolled offset into data array
Public dMaxIndicator As Double   'max price for the current indicator
Public dRangeIndicator As Double  'price range for the current indicator
Public dHeightIndicator As Double  'height of indicator panel
Public dMinIndicator As Double   'min price for the current indicator
Public dMinPrice As Double
Public dMaxPrice As Double
Public dRangePrice As Double
Public dHeightPrice As Double
Public iNumBarsPloted As Long
Public rRightSideOffset As Single  'diffence between total margin and draw margin
Public iMaxDrawRightX As Long   'right side draw margin
Public iBottomPlotMargin As Long
Public iBarDataPeriodMins As Long
Public iDateMarkerColor As Long
Public IsDrawing As Boolean    'inside drawing procedure
Public iBlankSpace As Long   'right side blank space
Public iBarSpacing As Long  'spacing between price bars
Public iScrolledAmount As Long   'number of bars scrolled
Public rSplit1 As Single   'location between price panel & vol panel
Public rSplit2 As Single   'location between vol & lower panels
Public xLeftMargin As Long  'left  margin
Public xRightMargin As Long  'right  margin
Public iTextHeight As Long  'height of picbox text
Public fKillSplash As Boolean  'flag for unloading splash form when using as progress ctrl
Public fCancelDrawingTool As Boolean   'cancel drawing tool loop
Public objDrawingTools As CDrawingTools  'used to pass values to the withevent class
Public iScrollIncrement As Long  'value the chart is scrolled

