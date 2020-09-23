VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Basic Stock Charting V1.0"
   ClientHeight    =   9030
   ClientLeft      =   1515
   ClientTop       =   1530
   ClientWidth     =   10650
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9030
   ScaleWidth      =   10650
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picSnap 
      AutoRedraw      =   -1  'True
      Height          =   1575
      Left            =   7740
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   5220
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtDrawInstruct 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   4620
      TabIndex        =   5
      Top             =   8160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2340
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0442
            Key             =   "ReDraw"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1A9C
            Key             =   "OpenFile"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1FDE
            Key             =   "Camera"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":20F0
            Key             =   "Indicators"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2414
            Key             =   "BlankGray"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2866
            Key             =   "DecBarSpace"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2978
            Key             =   "IncBarSpace"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2A8A
            Key             =   "About"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2E0B
            Key             =   "DrawingTools"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3191
            Key             =   "Options"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":354A
            Key             =   "ScrollData"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":38E6
            Key             =   "DownLoad"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox ChartBox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2685
      Left            =   7860
      MouseIcon       =   "Main.frx":3C7E
      ScaleHeight     =   175
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   4
      Top             =   300
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.PictureBox ChartBoxV 
      BackColor       =   &H00FFFFFF&
      Height          =   3765
      Left            =   420
      MouseIcon       =   "Main.frx":3DD0
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   473
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   7155
      Begin VB.Line Divider 
         BorderColor     =   &H000000FF&
         Index           =   1
         Visible         =   0   'False
         X1              =   20
         X2              =   452
         Y1              =   148
         Y2              =   148
      End
      Begin VB.Line Divider 
         BorderColor     =   &H000000FF&
         Index           =   0
         Visible         =   0   'False
         X1              =   16
         X2              =   448
         Y1              =   88
         Y2              =   88
      End
      Begin VB.Line ChLine1 
         Index           =   0
         Visible         =   0   'False
         X1              =   184
         X2              =   208
         Y1              =   196
         Y2              =   196
      End
      Begin VB.Line ChLine2 
         Index           =   0
         Visible         =   0   'False
         X1              =   196
         X2              =   196
         Y1              =   208
         Y2              =   180
      End
      Begin VB.Label lblMousePrice 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Left            =   5340
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   480
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.Toolbar tbLeft 
      Align           =   3  'Align Left
      Height          =   8655
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   15266
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OpenFile"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DownLoad"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Sep1"
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ReDraw"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "IncBarSpace"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DecBarSpace"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ScrollData"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Options"
            Object.Width           =   300
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Sep2"
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Indicators"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DrawingTools"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Sep3"
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Camera"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Sep4"
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "About"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbBottom 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8655
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5265
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuChartPopUp 
      Caption         =   "ChartPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuPuCancelDrawing 
         Caption         =   "Cancel Drawing"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPuSettingsChart 
         Caption         =   "Chart Options"
      End
      Begin VB.Menu mnuPuCrossHair 
         Caption         =   "CrossHair"
         Begin VB.Menu mnuPuCrossHairColor 
            Caption         =   "Color"
         End
         Begin VB.Menu mnuPuCrossHairMode 
            Caption         =   "Mode"
         End
      End
      Begin VB.Menu mnuPuBarSpacing 
         Caption         =   "Bar Spacing"
      End
      Begin VB.Menu mnuPuBlankSpace 
         Caption         =   "BlankSpace"
      End
      Begin VB.Menu mnuPuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPuIndSettings 
         Caption         =   "Indicator Settings"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Basic stock charting program. John Sugas 2002
'Hopefully this inspire some ideas of your own... no warranties or guarantees
'included. I'm sharing this for personal use only, don't sell this code...
'Code borrowed from others is in the MBorrowedCode module. Thanks to all who have
'shared their code....

Option Explicit

Public WithEvents DrawingTools As CDrawingTools
Attribute DrawingTools.VB_VarHelpID = -1

Private IsChartDrawn As Boolean  'drawing complete
Private fGotData As Boolean   'data is loaded
Private fClickingBarSpacing As Boolean 'changing barspacing flag
Private iMouseDataInfo As Long  'mouse data mode
Private iCrossHair As Long   'crosshair mode
Private iTimeDataInfo As Long 'data info mode
Private iMoveSplit As Long  'panel adjust mode
Private iWhichSplit  As Long  'which panel divider is picked
Private sCaption As String  'original form caption
Private sCapCurrent As String   'caption with current data info
Private dMaxVol As Double  'max vol data range
Private dHeightVol As Double  'height of vol panel
Private iMostRecentBarIndex As Long  'current working bar
Private fSwitch As Boolean   'animation toggle switch
Private iScrolledAmount As Long  'running count of amount chart is scrolled
Private iCalcdAvailBars2Plot As Long  'calculated number of bars to plot





Private Sub DrawingTools_DrawingDone()
    'cleanup after drawing is done
    stbBottom.Panels(3).Picture = Nothing
    txtDrawInstruct.Text = sEmpty
    txtDrawInstruct.Visible = False
End Sub

Private Sub DrawingTools_DrawingInstructions(sText As String)
    txtDrawInstruct.Text = sText$
End Sub

Private Sub DrawingTools_DrawingStarted()
    stbBottom.Panels(3).Picture = LoadResPicture(101, vbResIcon)
    'statusBar fonts not individual to each panel so we put a textbox on it
    'to change to bold for the drawing instructions
    txtDrawInstruct.Top = stbBottom.Top + 30
    txtDrawInstruct.Left = stbBottom.Panels(3).Left + stbBottom.Panels(3).Picture.Width * 0.75
    txtDrawInstruct.Height = stbBottom.Height - 30
    txtDrawInstruct.Width = stbBottom.Panels(3).Width - stbBottom.Panels(3).Picture.Width * 0.75
    txtDrawInstruct.ZOrder 0
    txtDrawInstruct.Visible = True
End Sub

Private Sub DrawingTools_DrawLoopIsRunning()
    'animation routine for visual notification to show that the draw loop is running
    fSwitch = Not fSwitch
    If fSwitch Then
        stbBottom.Panels(3).Picture = LoadResPicture(102, vbResIcon)
    Else
        stbBottom.Panels(3).Picture = LoadResPicture(101, vbResIcon)
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If GetAsyncKeyState(VK_CONTROL) <> 0 Then
        If GetAsyncKeyState(VK_END) <> 0 Then  'ctrl-end reset scroll
            If iScrolledAmount = 0 Then Exit Sub
            iScrolledAmount = 0
            Call ChartBoxDraw
        
        ElseIf GetAsyncKeyState(VK_HOME) <> 0 Then  'ctrl-home Max scroll
            If iScrolledAmount >= iUBaData - iCalcdAvailBars2Plot Then Exit Sub
            iScrolledAmount = iUBaData - iCalcdAvailBars2Plot
            Call tbLeft_MouseUp(1, 1, tbLeft.Buttons("ScrollData").Left + 10, tbLeft.Buttons("ScrollData").Top + 10)
            Call ChartBoxDraw
            
        ElseIf GetAsyncKeyState(VK_LEFT) <> 0 Then  'ctrl-left arrow scroll left 1
            If iScrolledAmount >= iUBaData - iCalcdAvailBars2Plot Then Exit Sub
            Call tbLeft_MouseUp(1, 0, tbLeft.Buttons("ScrollData").Left + 10, tbLeft.Buttons("ScrollData").Top + 10)
            Call ChartBoxDraw
            
        ElseIf GetAsyncKeyState(VK_RIGHT) <> 0 Then  'ctrl-right arrow scroll right 1
            If iScrolledAmount = 0 Then Exit Sub
            Call tbLeft_MouseUp(2, 0, tbLeft.Buttons("ScrollData").Left + 10, tbLeft.Buttons("ScrollData").Top + 10)
            Call ChartBoxDraw
            
        ElseIf GetAsyncKeyState(VK_PRIOR) <> 0 Then  'ctrl-page up scroll left scroll amount
            If iScrolledAmount >= iUBaData - iCalcdAvailBars2Plot Then Exit Sub
            Call tbLeft_MouseUp(1, 1, tbLeft.Buttons("ScrollData").Left + 10, tbLeft.Buttons("ScrollData").Top + 10)
            Call ChartBoxDraw
            
        ElseIf GetAsyncKeyState(VK_NEXT) <> 0 Then  'ctrl-page down scroll right scroll amount
            If iScrolledAmount = 0 Then Exit Sub
            Call tbLeft_MouseUp(2, 1, tbLeft.Buttons("ScrollData").Left + 10, tbLeft.Buttons("ScrollData").Top + 10)
            Call ChartBoxDraw
            
        End If
    End If
End Sub

Private Sub Form_Load()

    frmSplash.Show 0, Me
    DoEvents
    sINIsetFile$ = App.Path & "\Chart.INI"
    If Not ExistFile(sINIsetFile$) Then  'no INI file...create one
        Call MakeIniFile
    End If
    Call GetIniSettings
    Set MovAvgs = New MovAvg
    Call GetIndicatorSettings
    Set DrawingTools = New CDrawingTools
    Set DrawingTools.picBx = ChartBox
    Set DrawingTools.picBxV = ChartBoxV
    
    sCaption = Me.Caption & " - "
    lblMousePrice.Top = ChartBoxV.ScaleTop + 1
    lblMousePrice.Left = ChartBoxV.ScaleWidth
    Call SetColors
    Call SetupToolbar
    
    Show
    DoEvents
    
    If LoadData Then  'try to load data, set flag if success, draw chart
        fGotData = True
        Call SetMargins
        Call ChartBoxDraw
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If IsDrawing Then  'still in drawing loop so set cancel flag, cancel exit to kill loop
        fCancelDrawingTool = True
        Cancel = True
        Exit Sub
    End If
    If UnloadMode = 0 Or UnloadMode = 1 Then
        Dim iResult As Long
        iResult = MsgBox("Do you really want to exit program?", vbQuestion + vbYesNo, "Exiting Program...")
        If iResult = vbNo Then Cancel = True: Exit Sub
    End If
    Call SaveIniSettings
    Call EndWinsock
    Set oCurrentIndicator = Nothing
    Set DrawingTools = Nothing
    Set MovAvgs = Nothing
    Set frmMain = Nothing
End Sub

Public Sub SetColors()
    ChartBox.BackColor = iBackColor
    ChartBox.ForeColor = iForeColor
    ChartBoxV.BackColor = iBackColor
    ChartBoxV.ForeColor = iForeColor
    ChartBox.Font = sFontName
    ChartBox.FontSize = iFontSize
    ChartBox.FontBold = iFontBold
    ChartBox.FontItalic = iFontItalic
    ChartBoxV.Font = sFontName
    ChartBoxV.FontSize = iFontSize
    ChartBoxV.FontBold = iFontBold
    ChartBoxV.FontItalic = iFontItalic
    lblMousePrice.BackColor = iBackColor
    lblMousePrice.ForeColor = iMouseLabelColor

End Sub
Public Sub SetMargins()
    iTextHeight = ChartBox.TextHeight("X")
    xLeftMargin = ChartBox.ScaleLeft
    xRightMargin = ChartBox.ScaleWidth
    rRightSideOffset = ChartBox.ScaleWidth - iBlankSpace * 10
    If iBarDataPeriodMins < 0 Then 'Daily data... need less space for text
        iBottomPlotMargin = ChartBox.ScaleHeight - (iTextHeight)
    Else
        iBottomPlotMargin = ChartBox.ScaleHeight - (iTextHeight * 2)
    End If
    'validate the panel dividers location... make sure we can move them with the mouse
    'especially after setting defaults... a divider may not even be on the screen
    If rSplit1 < iBottomPlotMargin / 2 Then
        rSplit1 = iBottomPlotMargin / 2
        WriteIni sINIsetFile, "Settings", "WindowSplit1", CStr(rSplit1)
    End If
    If rSplit2 > iBottomPlotMargin - 50 Then
        rSplit2 = iBottomPlotMargin - 50
        WriteIni sINIsetFile, "Settings", "WindowSplit2", CStr(rSplit2)
    End If
    If rSplit1 > rSplit2 Then
        rSplit1 = rSplit2 - 10
        WriteIni sINIsetFile, "Settings", "WindowSplit1", CStr(rSplit1)
    End If
    If rSplit2 < rSplit1 Then
        rSplit2 = rSplit1 + 10
        WriteIni sINIsetFile, "Settings", "WindowSplit2", CStr(rSplit2)
    End If
    Divider(0).X1 = ChartBox.ScaleLeft - 5
    Divider(0).Y1 = rSplit1
    Divider(0).X2 = ChartBox.ScaleLeft - 3
    Divider(0).Y2 = rSplit1
    Divider(1).X1 = ChartBox.ScaleLeft - 5
    Divider(1).Y1 = rSplit2
    Divider(1).X2 = ChartBox.ScaleLeft - 3
    Divider(1).Y2 = rSplit2
End Sub
Public Sub Form_Resize()
   
    ChartBoxV.Width = Me.ScaleWidth - (ChartBoxV.Left * 2) + tbLeft.Width
    ChartBoxV.Height = Me.ScaleHeight - (ChartBoxV.Top + stbBottom.Height) ' + 120)
    
    ChartBox.Width = ChartBoxV.Width
    ChartBox.Height = ChartBoxV.Height
    lblMousePrice.Left = ChartBoxV.ScaleWidth - 95
    
    Call SetMargins
    iCalcdAvailBars2Plot = (Int(rRightSideOffset / iBarSpacing) + 1)
    
    If fGotData Then Call ChartBoxDraw
    ChartBoxV.Visible = True

End Sub
Private Function SnapToBar(x As Integer) As Integer
    Dim iXdiff As Integer, iMod As Integer
    Dim iUpperBar As Integer, iLowerBar As Integer
    
    'snap crosshairs to price bar plots
    iXdiff = rRightSideOffset - x
    iMod = iXdiff Mod iBarSpacing
    iUpperBar = x + iMod
    iLowerBar = iUpperBar - iBarSpacing
    
    'split the difference.. send to closest bar
    If x - iLowerBar <= iUpperBar - x Then
        SnapToBar = iLowerBar
    ElseIf x - iLowerBar > iUpperBar - x Then
        SnapToBar = iUpperBar
    End If
    
End Function
Private Sub CrossHairStart(x As Single, y As Single)
    Dim i As Long
    'setup & show crosshairs
    x = SnapToBar(CInt(x))
    ChLine1(0).DrawMode = iCrossHairMode
    ChLine2(0).DrawMode = iCrossHairMode ' 8 '6,8,15
    ChLine1(0).BorderColor = iCrossHairColor
    ChLine2(0).BorderColor = iCrossHairColor
    ChLine1(0).Y1 = y
    ChLine1(0).Y2 = y
    ChLine2(0).X1 = x
    ChLine2(0).X2 = x
    ChLine2(0).Y1 = ChartBox.ScaleHeight
    ChLine2(0).Y2 = ChartBox.ScaleTop
    ChLine1(0).X1 = ChartBox.ScaleLeft
    ChLine1(0).X2 = ChartBox.ScaleWidth
    ChLine1(0).Visible = True
    ChLine2(0).Visible = True
    Call ShowCursor(False)
End Sub

Private Sub CrossHairMoving(x As Single, y As Single)
    Dim i As Long
    'move the crosshairs
    x = SnapToBar(CInt(x))
    ChLine1(0).Y1 = y
    ChLine1(0).Y2 = y
    ChLine2(0).X1 = x
    ChLine2(0).X2 = x
End Sub
Private Function CursorInfo(x As Integer, y As Integer) As String
    Dim ssOutput As String, iX As Integer, ssInfo As String
    
    'get the info for the bar under the crosshairs
    On Error Resume Next
    iX = iUBaData - (rRightSideOffset - x) / iBarSpacing
    If iMouseDataInfo = 1 Then
        ssInfo$ = aData(iX).sDate & vbCrLf & "      " _
        & aData(iX).sTime & vbCrLf _
        & "O: " & aData(iX).dOpen & vbCrLf _
        & "H: " & aData(iX).dHigh & vbCrLf _
        & "L: " & aData(iX).dLow & vbCrLf _
        & "C: " & aData(iX).dClose & vbCrLf _
        & "V: " & aData(iX).iVol
    End If
    
    'only get info if in the price plot panel
    If y < rSplit1 Then
        ssOutput$ = CStr(Format(Round(dMaxPrice - ((y - 4) * dRangePrice) / (rSplit1 - 8), 2), "0.00"))
        If iMouseDataInfo = 1 Then
            ssOutput$ = ssOutput$ & vbCrLf & ssInfo$
        End If
    Else
        CursorInfo = ""
    End If
        
    CursorInfo$ = ssOutput$
End Function
Private Sub ChartBoxv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim iPrevW As Integer

    If Button = 1 Then
        If DrawingTools.MouseClickEnabled = True Then
            DrawingTools.MouseClickNum = DrawingTools.MouseClickNum + 1  'flag and var for drawing tools
        Else
            If iCrossHair = 1 Then  'show crosshair
                Call CrossHairStart(x, y)
                If x <= xRightMargin Then
                    If Timer - iTimeDataInfo < 1.5 Or IsDrawing = True Then
                        iMouseDataInfo = 1  'show info flag
                    End If
                    lblMousePrice.Visible = True
                    lblMousePrice.Caption = CursorInfo(x \ 1, y \ 1)
                End If
            Else
                If rSplit1 > y - 3 And rSplit1 < y + 3 Then 'between price-vol panels
                    iWhichSplit = 1
                    Divider(0).X2 = ChartBoxV.ScaleWidth + 5
                    Divider(0).Visible = True
                ElseIf rSplit2 > y - 3 And rSplit2 < y + 3 Then  'between vol-ind panels
                    iWhichSplit = 2
                    Divider(1).X2 = ChartBoxV.ScaleWidth + 5
                    Divider(1).Visible = True
                End If
            End If
        End If
        DrawingTools.Xcurr = x  'get x&y for drawing tools click
        DrawingTools.Ycurr = y
        iTimeDataInfo = Timer  'set begin time for mouse info label show
    ElseIf Button = 2 Then
        'If iMouseClickNum <> 0 Then Exit Sub
        PopupMenu mnuChartPopUp
    End If
    
End Sub

Private Sub ChartBoxv_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim dy As Double, iOldMode As Integer

    dy = ScaleY(2, vbTwips, vbPixels)
    'Me.Caption = sCapCurrent$ & "       " & _
            CStr(Format(dMaxPrice - ((y * dRangePrice) / (rSplit1)), "0.####"))
    Select Case Button
        Case 1
            If iMoveSplit = 0 Then  'not resizing the panels
                Call CrossHairMoving(x, y)  'crosshairs tracking mouse
                If x <= xRightMargin Then  'show info label if not past price plot
                    lblMousePrice.Visible = True
                    lblMousePrice.Caption = CursorInfo(x \ 1, y \ 1)
                End If
            Else 'resize the plot panels
                If iWhichSplit = 1 Then
                    If y < rSplit2 - 10 And y > iBottomPlotMargin / 2 Then  'set extend bounds
                        rSplit1 = y
                        Divider(0).Y1 = rSplit1
                        Divider(0).Y2 = rSplit1
                    Else
                        Call ResetMousePointerAfterSplitMove
                        Call ChartBoxDraw
                    End If
                ElseIf iWhichSplit = 2 Then
                    If y > rSplit1 + 10 And y < iBottomPlotMargin - 50 Then  'set extend bounds
                        rSplit2 = y
                        Divider(1).Y1 = rSplit2
                        Divider(1).Y2 = rSplit2
                    Else
                        Call ResetMousePointerAfterSplitMove
                        Call ChartBoxDraw
                    End If
                End If
            End If
        Case Else
            If IsDrawing = False Then
                If (y > rSplit1 - 3 And y < rSplit1 + 3) Or _
                            (y > rSplit2 - 3 And y < rSplit2 + 3) Then 'over the divider
                    ChartBoxV.MousePointer = 99  'set curser to Horzsplitter
                    iMoveSplit = 1
                    iCrossHair = 0
                Else  'normal curser status
                    Call ResetMousePointerAfterSplitMove
                End If
            End If
    End Select
    On Error Resume Next
    DrawingTools.XcurrMov = x  'get moving x&y for drawing tools
    DrawingTools.YcurrMov = y
    
End Sub
Private Sub ResetMousePointerAfterSplitMove()
    ChartBoxV.MousePointer = vbDefault
    iMoveSplit = 0
    iCrossHair = 1
End Sub
Private Sub CrossHairStop()
    Dim i As Long
    'hide the crosshairs and info label
    For i = 0 To 2
        ChLine1(0).Visible = False
        ChLine2(0).Visible = False
    Next
    lblMousePrice.Visible = False
    iMouseDataInfo = 0
    Call ShowCursor(True)
End Sub
Private Sub ChartBoxv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If iMoveSplit = 1 Then 'save positions & redraw chart if panels resized
            WriteIni sINIsetFile, "Settings", "WindowSplit1", CStr(rSplit1)
            WriteIni sINIsetFile, "Settings", "WindowSplit2", CStr(rSplit2)
            Divider(0).Visible = False
            Divider(0).X2 = ChartBox.ScaleLeft - 3
            Divider(1).Visible = False
            Divider(1).X2 = ChartBox.ScaleLeft - 3
            Call ChartBoxDraw
        End If
        Call CrossHairStop  'hide crosshairs
        iMoveSplit = 0
    ElseIf Button = 2 Then
        'If iMouseClickNum > 0 Then iMouseClickNum = iMouseClickNum - 1
    End If
    
End Sub


Private Sub cmdBarSpacing_Click(Index As Integer)
    
    If fClickingBarSpacing = 1 Or IsDrawing = 1 Then Exit Sub 'finish current op
    fClickingBarSpacing = 1
    Select Case Index
        Case 0      'increase bar spacing
            If iBarSpacing < 30 Then
                iBarSpacing = iBarSpacing + 1
            End If
        Case 1      'decrease barspacing
            If iBarSpacing > 1 Then
                iBarSpacing = iBarSpacing - 1
            End If
    End Select
    Call ChartBoxDraw
    
End Sub
Private Sub GetDataInfo()
    
    Screen.MousePointer = vbHourglass
    
    'display the last price data on the form caption
    frmMain.Caption = sCaption & aData(iUBaData).sDate & "  " & aData(iUBaData).sTime _
                & "   O: " & aData(iUBaData).dOpen & "   H: " & aData(iUBaData).dHigh _
                & "   L:" & aData(iUBaData).dLow & "   C: " & aData(iUBaData).dClose _
                & "   V: " & aData(iUBaData).iVol

    
    sCapCurrent$ = Me.Caption
    Select Case iBarDataPeriodMins
        Case -1
            stbBottom.Panels(5).Text = "Interval: Daily"
        Case -2
            stbBottom.Panels(5).Text = "Interval: Weekly"
        Case Is > 0
            stbBottom.Panels(5).Text = "Interval: " & iBarDataPeriodMins & " min"
    End Select

    stbBottom.Panels(4).Text = "Symbol: " & sSymbol$
    stbBottom.Panels(1).Text = "File: " & sFileName$

    Screen.MousePointer = vbDefault
End Sub
Public Sub ChartBoxDraw()

    Dim iStyle As Integer, iDmode As Integer, iTimeTrigger As Integer
    Dim iDateTrigger As Integer, iDateSpacing As Integer, iDayOfWeek As Integer, iCurrTime As Long
    Dim sStartTime As String, sTimeDif As String, iTimeH As Integer, iTimeM As Integer
    Dim x As Single, Y1 As Single, Y2 As Single, iCount As Integer, iCnt930 As Integer, iX2 As Integer
    Dim X1 As Long, X2 As Long, sTime As String, iNextX As Integer, sDate As String, sTemp As String
    Dim iLastDx As Long, lPriceLabelWidth As Long, sDateLast As String, iOldDayNum As Integer
    Dim iCnt As Integer, iCnt2 As Integer, rStartPos As Single, rSpacing As Single, iDateCnt As Integer
    Dim iDrWidth As Integer, sDateShort As String, dHeight2RangeRatio As Double, rLegionStartPrice As Single
    Dim iPriceLegionPosX As Long, j As Integer, iWeekTrigger As Integer
    
    Call GetDataInfo
    If fGotData = False Then Exit Sub  'no data... no plot
    If IsDrawing = True Then Exit Sub  'wait till done plotting before plotting again
    IsDrawing = True
    Screen.MousePointer = vbHourglass
On Error Resume Next

    ChartBox.Cls
    iStyle = ChartBox.DrawStyle
    iDmode = ChartBox.DrawMode
    iDrWidth = ChartBox.DrawWidth
    ChartBox.DrawMode = vbCopyPen
    
    
'Debug.Print "rsOff/barSp: "; rRightSideOffset / iBarSpacing
'Debug.Print "iNumBarsPloted: "; iNumBarsPloted
'Debug.Print "calc#:"; (Int(rRightSideOffset / iBarSpacing) + 1)
    
    iStartIndex = (iUBaData - iScrolledAmount)
    If (iStartIndex - iCalcdAvailBars2Plot) > 0 Then
        'we have more data than we are plotting
        iLboundDataStart = (iStartIndex - iCalcdAvailBars2Plot)
        tbLeft.Buttons("ScrollData").Enabled = True
    Else
        'can't plot less than the data we have
        iLboundDataStart = LBound(aData())
        tbLeft.Buttons("ScrollData").Enabled = False 'nothing to scroll
    End If

    
    'check max price against total bars to plot ... find max and min values
    dMaxPrice = 0
    dMinPrice = 999999
    dMaxVol = 1
    For j = iLboundDataStart To iStartIndex
        If aData(j).dHigh > dMaxPrice Then dMaxPrice = aData(j).dHigh
        If aData(j).dLow < dMinPrice And _
                aData(j).dLow > 0 Then dMinPrice = aData(j).dLow   'if 0 in the data ignore
        If aData(j).iVol > dMaxVol Then dMaxVol = aData(j).iVol
    Next j

'Debug.Print dMaxPrice; "  "; dMinPrice
    
    dRangePrice = dMaxPrice - dMinPrice  'price data range
    iNumBarsPloted = 0

    dHeightPrice = rSplit1 - 8 'total plot height with 4 pixels/side margin
    dHeight2RangeRatio = dHeightPrice / dRangePrice  'price per pixel
    dHeightVol = rSplit2 - rSplit1 - 8 'total plot height for vol panel
    lPriceLabelWidth = ChartBox.TextWidth(CStr(Format(dMaxPrice, "##.00"))) 'price legion width
    iPriceLegionPosX = ChartBox.ScaleWidth - lPriceLabelWidth - 5  'start of price legion
    iMaxDrawRightX = iPriceLegionPosX - 3  'max plot in x

    
'****************************************
'**************data pane Hgrid And price labels
'****************************************
    'initalize for grid
    ChartBox.DrawStyle = vbDot
    X1 = ChartBox.ScaleLeft
    X2 = ChartBox.ScaleWidth
    
    'price legion spacing is determined by the max price range.
    Select Case dRangePrice
        Case 0 To 1
            rSpacing = (0.05)
        Case 1 To 2
            rSpacing = (0.1)
        Case 2 To 4
            rSpacing = (0.25)
        Case 4 To 10
            rSpacing = (0.5)
        Case 10 To 20
            rSpacing = (1)
        Case 20 To 30
            rSpacing = (2)
        Case Else
            rSpacing = (3)
    End Select
    'start at the price integer... and work both directions
    'makes a better looking legion using rounded price numbers
    'like 50.70 instead of -> 50.67
    rLegionStartPrice = Round(dMaxPrice - rSpacing, 1)
    rStartPos = 4 + (dMaxPrice - rLegionStartPrice) * dHeight2RangeRatio

    Y1 = Round(rStartPos)
'Debug.Print Round(dMaxPrice - rSpacing, 1)
'Debug.Print rSpacing
'Debug.Print ((dMaxPrice - Int(dMaxPrice)) * dHeight2RangeRatio)
'Debug.Print "y1:"; Y1
'Debug.Print "dMaxPrice:"; dMaxPrice
'Debug.Print "dMin:"; dMinPrice
'Debug.Print "Int(dMaxP:"; Int(dMaxPrice)
'Debug.Print "dH2RRat:"; dHeight2RangeRatio

    Do While Y1 < (rSplit1 - iTextHeight)  'make sure don't print on the divider
        DoEvents
        ChartBox.Line (X1, Y1)-(iPriceLegionPosX, Y1), iGridColor
        ChartBox.CurrentX = iPriceLegionPosX
        ChartBox.CurrentY = ChartBox.CurrentY - iTextHeight / 2 + 1
        ChartBox.Print Format(rLegionStartPrice - (iCnt * rSpacing), "##.00")
        If iCnt = 0 Then 'it is the first horz grid line at the interger val
            If rStartPos - (rSpacing * dHeight2RangeRatio) > 4 Then  'we aren't at the top
                Y2 = rStartPos
                'now we work from the integer val up to the top of the price panel
                Do While Y2 > iTextHeight * 3 '/ 2 + 1
                    DoEvents
                    iCnt2 = iCnt2 + 1 'grid up count
                    Y2 = 4 + ((dMaxPrice - rLegionStartPrice - iCnt2 * rSpacing) * dHeight2RangeRatio)
                    ChartBox.Line (X1, Y2)-(iPriceLegionPosX, Y2), iGridColor
                    ChartBox.CurrentX = iPriceLegionPosX
                    ChartBox.CurrentY = ChartBox.CurrentY - iTextHeight / 2 + 1
                    ChartBox.Print Format(rLegionStartPrice + (iCnt2 * rSpacing), "##.00")
                Loop
            End If
        End If
        iCnt = iCnt + 1 'grid down count
        Y1 = 4 + ((dMaxPrice - rLegionStartPrice + iCnt * rSpacing) * dHeight2RangeRatio)
    Loop
'*****************vol pane Hgrid
    ChartBox.DrawStyle = vbDot
    ChartBox.DrawWidth = 1
    iPriceLegionPosX = ChartBox.ScaleWidth - ChartBox.TextWidth(CStr(dMaxVol)) - 10
    For Y1 = rSplit2 - dMaxVol * (dHeightVol / dMaxVol) To rSplit2 Step 20
        ChartBox.Line (X1, Y1)-(iPriceLegionPosX, Y1), iGridColor
        ChartBox.CurrentX = iPriceLegionPosX
        ChartBox.CurrentY = ChartBox.CurrentY - iTextHeight / 2 + 1
        If ChartBox.CurrentY < rSplit2 - iTextHeight Then _
            ChartBox.Print Round((rSplit2 - Y1) / (dHeightVol / dMaxVol))
    Next Y1

'***************************************************************************
'*****************************************************************************
'*************************start draw data loop
'****************************************************************************
'****************************************************************************

    x = rRightSideOffset
    iNextX = x
    iLastDx = x
    iCount = iStartIndex
    iMostRecentBarIndex = iCount
    'start at right side, go Left. Stop if no more data or we reached left side
    Do While x > 0 And iCount > LBound(aData, 1)
     
'Debug.Print "x: "; X; "  #bars: "; iNumBarsPloted

'*************time legion & vert grid
        
        sTime$ = aData(iCount).sTime
'Debug.Print stime$
        sDate$ = aData(iCount).sDate
        iDayOfWeek = Weekday(sDate$)
'Debug.Print sDate$; "  "; Weekday(sDate$)
'Debug.Print sDatelast$
        
        'need to calculate the first day of week different for end of day data
        If iBarDataPeriodMins < 0 Then
            Dim iPrevDay As Integer
            iPrevDay = Weekday(aData(iCount - 1).sDate)
            Select Case iDayOfWeek
                Case 2, 3 ' mon or tue
                    'if prev day is a thur or fri then week flag=true
                    If iPrevDay = 6 Or iPrevDay = 5 Then iWeekTrigger = 1
                Case Else
                    iWeekTrigger = 0
            End Select
'Debug.Print "iPv:"; iPrevDay; " day:"; iDayOfWeek; " tr:"; iWeekTrigger
        ElseIf iBarDataPeriodMins > 0 Then
            Select Case iOldDayNum
                Case 2, 3 ' mon or tue
                    'if prev day is a thur or fri then week flag=true
                    If iDayOfWeek = 6 Or iDayOfWeek = 5 Then iWeekTrigger = 1
                Case Else
                    iWeekTrigger = 0
            End Select
            iOldDayNum = iDayOfWeek
            
        End If
        sTime$ = Trim(Mid(sTime$, InStr(sTime$, " ") + 1))
'Debug.Print stime$
        Select Case Mid(sTime$, InStr(sTime$, ":") + 1)
            Case "00", "15", "30", "45"  'keep the time legion on "pretty" values
                iTimeTrigger = 1
        End Select
        
        
        ChartBox.DrawStyle = vbSolid
        ChartBox.DrawWidth = 1
'***************Date labels
        Dim iX As Integer  'copy of x pos. for manipulation
        Select Case sTime$
            Case "0930", "09:30", "1600", "16:00"
                'start/end of traditional trading day
                iCnt930 = iCnt930 + 1
                If iCnt930 = 1 Then iX = x  'save current x
        End Select
        If sDate$ <> sDateLast$ Then iDateTrigger = 1 'make sure of new date
'Debug.Print iMostRecentBarIndex
'Debug.Print "iLastDx:"; iLastDx; ChartBox.ScaleWidth
'Debug.Print stime$
'Debug.Print iBarDataPeriodMins
        If iDateTrigger = 1 Then
'Debug.Print "iDateTrigger:"; iDateTrigger
'Debug.Print "x:"; x
            If iLastDx <> rRightSideOffset Or iBarDataPeriodMins < 0 Then
                Dim CurrY As Long
                CurrY = ChartBox.ScaleHeight - (iTextHeight)
                ChartBox.CurrentY = CurrY
                ChartBox.CurrentX = x + 1
                If iBarDataPeriodMins > 0 Then  'minute data
                    sDateShort$ = Left$(sDateLast$, 5)
                    If iLastDx - x > ChartBox.TextWidth(sDateLast$) + 10 Then
                        'we have room for long date string
                        ChartBox.Print sDateLast$
                        ChartBox.CurrentX = x
                        ChartBox.Line (x, ChartBox.ScaleHeight)-(x, iBottomPlotMargin), iDateMarkerColor
                        iLastDx = x
                    ElseIf iLastDx - x > ChartBox.TextWidth(sDateShort$) + 10 Then
                        'use short date string
                        ChartBox.Print sDateShort$
                        ChartBox.CurrentX = x
                        ChartBox.Line (x, ChartBox.ScaleHeight)-(x, iBottomPlotMargin), iDateMarkerColor
                        iLastDx = x
                    Else
                        'iLastDx stays the same until we have room for the string
                    End If
                    If x <> rRightSideOffset Then iDateCnt = iDateCnt + 1
'Debug.Print "x:"; x; " rRSO:"; rRightSideOffset
                ElseIf iBarDataPeriodMins < 0 Then  'daily data
'Debug.Print iLastDx; rRightSideOffset
'Debug.Print iLastDx - x; ChartBox.TextWidth(sDate$)
'Debug.Print iWeekTrigger; sDate
                    sDateShort$ = Left$(sDate$, 5)
                    If iLastDx - x > ChartBox.TextWidth(sDate$) + 10 Or _
                           iCount = iMostRecentBarIndex Then
                        'we have room for long date string
                        ChartBox.Print sDate$
                        ChartBox.CurrentX = x
                        ChartBox.Line (x, ChartBox.ScaleHeight)-(x, iBottomPlotMargin), iDateMarkerColor
                        ChartBox.DrawStyle = vbDot
                        ChartBox.Line (x, ChartBox.ScaleTop)-(x, iBottomPlotMargin), iGridColor
                        iLastDx = x
                    ElseIf iLastDx - x > ChartBox.TextWidth(sDateShort$) + 10 Then
                        'use short date string
                        ChartBox.Print sDateShort$
                        ChartBox.CurrentX = x
                        ChartBox.Line (x, ChartBox.ScaleHeight)-(x, iBottomPlotMargin), iDateMarkerColor
                        ChartBox.DrawStyle = vbDot
                        ChartBox.Line (x, ChartBox.ScaleTop)-(x, iBottomPlotMargin), iGridColor
                        iLastDx = x
                    Else
                        'iLastDx stays the same until we have room for the string
                    End If
                End If
            Else
                
                iLastDx = ChartBox.ScaleWidth
            End If
            
        End If
                
'Debug.Print "iDateCnt"; iDateCnt
'Debug.Print "x:"; x; " rRSO:"; rRightSideOffset
'Debug.Print "x:"; x; " iC:"; iCount; " LB:"; LBound(aData, 1)
        If iDateCnt = 0 Then   'print date if all data is for 1 day
            If x - iBarSpacing <= 0 Or iCount = LBound(aData, 1) + 1 Then
                ChartBox.CurrentY = ChartBox.ScaleHeight - (iTextHeight)
                ChartBox.CurrentX = ChartBox.ScaleLeft + 2
                ChartBox.Print sDateLast$
            End If
        End If
                    
        '***** Time labels
        If iTimeTrigger = 1 Then
            If iNextX >= x Then 'we have room for the time string
                ChartBox.CurrentX = x - (ChartBox.TextWidth(sTime$) / 2)
                ChartBox.CurrentY = iBottomPlotMargin
                ChartBox.Print sTime$
        '*****' vert grid
                ChartBox.DrawStyle = vbDot
                Y1 = ChartBox.ScaleTop
                Y2 = iBottomPlotMargin  ' ChartBox.ScaleHeight - 35
                ChartBox.Line (x, Y1)-(x, Y2), iGridColor
                'short "pointer line" in red to time string
                Y1 = iBottomPlotMargin - 10  'ChartBox.ScaleHeight - 25
                ChartBox.DrawStyle = vbSolid
                ChartBox.Line (x, Y1)-(x, Y2), vbRed
                iNextX = x - ChartBox.TextWidth(sTime$) - 5
            End If
            'sDateLast$ = sDate$
        End If
        ChartBox.DrawStyle = vbDot
        'draw the day marker here so it will be on top of grid
        If iDateTrigger = 1 And iBarDataPeriodMins > 0 Then
            sDateLast$ = sDate$
            ChartBox.Line (x, ChartBox.ScaleTop)-(x, iBottomPlotMargin), iDateMarkerColor '1911939  'iGridColor
        End If
'************************************
''***********************************
'*****price bar plot
'************************************
'************************************
        ChartBox.DrawStyle = vbSolid
        Select Case iTicType
            Case ttHLOC  'standard HLOC bar plot
                'price body
                Y1 = 4 + (dMaxPrice - aData(iCount).dHigh) * dHeight2RangeRatio
                Y2 = 4 + (dMaxPrice - aData(iCount).dLow) * dHeight2RangeRatio
                ChartBox.Line (x, Y1)-(x, Y2), iTicBodyColor
                'open tick
                Y1 = 4 + (dMaxPrice - aData(iCount).dOpen) * dHeight2RangeRatio
                ChartBox.Line (x - 2, Y1)-(x + 1, Y1), iTicOpenColor
                'close tick
                Y1 = 4 + (dMaxPrice - aData(iCount).dClose) * dHeight2RangeRatio
                ChartBox.Line (x, Y1)-(x + 3, Y1), iTicCloseColor
            Case ttCandle  'candle plot
                Dim iCandleColor As Long
                'if close >open then plot color is up color
                If aData(iCount).dClose - aData(iCount).dOpen >= 0 Then
                    iCandleColor = iTicCandleUpColor
                Else
                    iCandleColor = iTicCandleDnColor
                End If
                'price body
                Y1 = 4 + (dMaxPrice - aData(iCount).dOpen) * dHeight2RangeRatio  'open
                Y2 = 4 + (dMaxPrice - aData(iCount).dClose) * dHeight2RangeRatio  'close
                If iBarSpacing > 6 Then  'draw a "fatter" candle body
                    ChartBox.Line (x - 2, Y1)-(x + 3, Y2), iCandleColor, BF
                Else
                    ChartBox.Line (x - 1, Y1)-(x + 1, Y2), iCandleColor, BF
                End If
                'wick  from high to lo
                Y1 = 4 + (dMaxPrice - aData(iCount).dHigh) * dHeight2RangeRatio   'hi
                Y2 = 4 + (dMaxPrice - aData(iCount).dLow) * dHeight2RangeRatio   'lo
                ChartBox.Line (x, Y1)-(x, Y2), iCandleColor
            Case ttLine 'only plot from close to close
                Y1 = 4 + (dMaxPrice - aData(iCount).dClose) * dHeight2RangeRatio
                Y2 = 4 + (dMaxPrice - aData(iCount - 1).dClose) * dHeight2RangeRatio
                ChartBox.Line (x - iBarSpacing, Y2)-(x, Y1), iTicCloseColor
        End Select
'************************vol data
        ChartBox.DrawStyle = vbSolid
        ChartBox.DrawWidth = 2
        Y1 = rSplit2 - 1
        Y2 = rSplit2 - (aData(iCount).iVol * (dHeightVol / dMaxVol))
        ChartBox.Line (x, Y1)-(x, Y2), iVolColor
        
'*******************set-up for next bar
        iNumBarsPloted = iNumBarsPloted + 1
        iTimeTrigger = 0
        iDateTrigger = 0
        iCount = iCount - 1
        x = x - iBarSpacing
    Loop
    
'****print vol data
    sTemp$ = "Volume: " & aData(iCount).iVol
    'draw a "blackout rect for better visibility of the text
    ChartBox.Line (1, rSplit1 + 3)-(1 + ChartBox.TextWidth(sTemp$), rSplit1 + 3 + ChartBox.TextHeight(sTemp$)), iBackColor, BF
    ChartBox.CurrentX = 1
    ChartBox.CurrentY = rSplit1 + 3
    ChartBox.Print "Volume: " & aData(iCount).iVol
    
    iX = 0
'Debug.Print "iNumBarsPloted: "; iNumBarsPloted

'******************************************
'********plot indicators
    Call PlotAvg
    Call PlotIndicator
    
'********draw dividers
    ChartBox.DrawStyle = vbSolid
    ChartBox.Line (0, rSplit1)-(ChartBox.ScaleWidth + 5, rSplit1), vbRed
    ChartBox.Line (0, rSplit2)-(ChartBox.ScaleWidth + 5, rSplit2), vbRed

'****************exit clean up
    ChartBox.DrawMode = iDmode
    ChartBox.DrawStyle = iStyle
    ChartBox.DrawWidth = iDrWidth
    ChartBoxV.Picture = ChartBox.Image
    IsChartDrawn = True
    IsDrawing = 0
    fClickingBarSpacing = 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuPuBarSpacing_Click()
    Dim sText As String, sInpResult As String
    
    sText$ = "Enter Bar Spacing....  " & vbCrLf & vbCrLf _
            & "Current Setting: " & iBarSpacing
    sInpResult$ = InputBox(sText$, sSettingChange$, iBarSpacing)
    
    If sInpResult$ <> "" And IsNumeric(sInpResult$) Then
        If Val(sInpResult$) < 1 Then Exit Sub
        iBarSpacing = CInt(sInpResult$)
        Call ChartBoxDraw
    End If
End Sub

Private Sub mnuPuBlankSpace_Click()
    Dim sText As String, sInpResult As String
    
    sText$ = "Enter Right side of chart 'Blank Space'....10 Minimum. " _
            & vbCrLf & vbCrLf & "Current Setting: " & iBlankSpace
    sInpResult$ = InputBox(sText$, sSettingChange$, iBlankSpace)
    
    If sInpResult <> "" And IsNumeric(sInpResult$) Then
        If Val(sInpResult$) < 10 Then Exit Sub
        iBlankSpace = CInt(sInpResult$)
        Call SetMargins
        Call ChartBoxDraw
    End If
End Sub

Private Sub mnuPuCancelDrawing_Click()
    fCancelDrawingTool = True
End Sub

Private Sub mnuPuCrossHairColor_Click()
    iCrossHairColor = GetColorDlg(iCrossHairColor)
End Sub

Private Sub mnuPuCrossHairMode_Click()
    Dim sText As String, sInpResult As String

    sText$ = "Enter new DrawMode for crosshairs...." _
            & "Any number from 1 to 16.  " _
            & "6,8,15 work best...   15 is default."
    sInpResult$ = InputBox(sText$, sSettingChange$, iCrossHairMode)
    
    If sInpResult$ <> "" And IsNumeric(sInpResult$) Then
        If sInpResult$ > 0 And sInpResult$ < 17 Then _
            iCrossHairMode = CInt(sInpResult$)
    End If
    
End Sub

Private Sub mnuPuIndSettings_Click()
    frmIndicators.Show 1, Me
End Sub

Private Sub mnuPuSettingsChart_Click()
    Call GetOptionsDlg
End Sub

Private Sub stbBottom_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If InStr(stbBottom.Panels(4).Text, sUnknownSymbol$) <> 0 Then
        stbBottom.Panels(4).ToolTipText = "DblClick to edit"
    Else
        stbBottom.Panels(4).ToolTipText = sEmpty
    End If
End Sub

Private Sub stbBottom_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
    Select Case Panel.Index
        Case 4
            'if the symbol in unknown then it can be entered by dblclk the symbol status panel
            If InStr(Panel.Text, sUnknownSymbol$) <> 0 Then MsgBox "left as exercise for....."
    End Select
End Sub

Private Sub SetupToolbar()
    
    tbLeft.ImageList = imgList  'set tb image list
    
    'set button images
    tbLeft.Buttons("ReDraw").Image = "ReDraw"
    tbLeft.Buttons("IncBarSpace").Image = "IncBarSpace"
    tbLeft.Buttons("DecBarSpace").Image = "DecBarSpace"
    tbLeft.Buttons("ScrollData").Image = "ScrollData"
    tbLeft.Buttons("OpenFile").Image = "OpenFile"
    tbLeft.Buttons("DownLoad").Image = "DownLoad"
    tbLeft.Buttons("Options").Image = "Options"
    tbLeft.Buttons("Indicators").Image = "Indicators"
    tbLeft.Buttons("DrawingTools").Image = "DrawingTools"
    tbLeft.Buttons("Camera").Image = "Camera"
    tbLeft.Buttons("About").Image = "About"
    
    'set tb tooltips
    tbLeft.Buttons("ReDraw").ToolTipText = "ReDraw"
    tbLeft.Buttons("IncBarSpace").ToolTipText = "Increase BarSpacing"
    tbLeft.Buttons("DecBarSpace").ToolTipText = "Decrease BarSpacing"
    tbLeft.Buttons("ScrollData").ToolTipText = "Scroll-LButton Left 1-RButton Right 1- +Shift 10+ Incr."
    tbLeft.Buttons("OpenFile").ToolTipText = "OpenFile"
    tbLeft.Buttons("DownLoad").ToolTipText = "Download Data"
    tbLeft.Buttons("Options").ToolTipText = "Options"
    tbLeft.Buttons("Indicators").ToolTipText = "Indicators"
    tbLeft.Buttons("DrawingTools").ToolTipText = "DrawingTools"
    tbLeft.Buttons("Camera").ToolTipText = "ScreenCapture"
    tbLeft.Buttons("About").ToolTipText = "About"
    
End Sub
Private Sub tbLeft_ButtonClick(ByVal Button As MSComctlLib.Button)

'Debug.Print Button.Key
    Select Case Button.Key  'handle tb click events
        Case "OpenFile"
            Call GetDataFile
        
        Case "DownLoad"
            frmDownLoad.Show 1, Me
        Case "ReDraw"
            Call ChartBoxDraw
        
        Case "IncBarSpace"
            If fClickingBarSpacing = True Or IsDrawing = True Then Exit Sub
            fClickingBarSpacing = True
            If iBarSpacing < 30 Then
                iBarSpacing = iBarSpacing + 1
            End If
            iCalcdAvailBars2Plot = (Int(rRightSideOffset / iBarSpacing) + 1)
            WriteIni sINIsetFile, "Settings", "BarSpacing", CStr(iBarSpacing)
            Call ChartBoxDraw
        
        Case "DecBarSpace"
            If fClickingBarSpacing = True Or IsDrawing = True Then Exit Sub
            fClickingBarSpacing = True
            If iBarSpacing > 1 Then
                iBarSpacing = iBarSpacing - 1
            End If
            iCalcdAvailBars2Plot = (Int(rRightSideOffset / iBarSpacing) + 1)
            WriteIni sINIsetFile, "Settings", "BarSpacing", CStr(iBarSpacing)
            Call ChartBoxDraw
        
        Case "ScrollData"
            'need to catch the right button click in the mouse up event

        Case "Options"
            Call GetOptionsDlg
        
        Case "Indicators"
            frmIndicators.Show 1, Me
        
        Case "DrawingTools"
            Set objDrawingTools = DrawingTools
            frmDrawingTools.Show 1, Me
            Set objDrawingTools = Nothing
        
        Case "Camera"
            Call CheckForSnapDir
            Call GetAndSaveSnapShot
        
        Case "About"
            frmAbout.Show 0, Me
    End Select
        
End Sub

Private Sub tbLeft_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'It's an ugly hack to get the right mouse click on the toolbar but since the click
    'won't handle right buttons and it fires after the mouseUp event so we need
    'to find the x-y coord. for the button and determine if is the one we want....
'Debug.Print "X:"; x; " y:"; y
'Debug.Print tbLeft.Buttons("ScrollData").Top; " "; tbLeft.Buttons("ScrollData").Top + tbLeft.Buttons("ScrollData").Height

    If y > tbLeft.Buttons("ScrollData").Top And _
        y < tbLeft.Buttons("ScrollData").Top + tbLeft.Buttons("ScrollData").Height Then
        If Shift Then  'shift button pressed.. large incr.
            If Button = 1 Then
                iScrolledAmount = iScrolledAmount + iScrollIncrement
            ElseIf Button = 2 Then
                iScrolledAmount = iScrolledAmount - iScrollIncrement
            End If
        Else  'normal 1 bar scroll increment
            If Button = 1 Then
                iScrolledAmount = iScrolledAmount + 1
            ElseIf Button = 2 Then
                iScrolledAmount = iScrolledAmount - 1
            End If
        End If
        If iScrolledAmount < 0 Then
            iScrolledAmount = 0
'        ElseIf iScrolledAmount > (iUBaData - iScrolledAmount) - iCalcdAvailBars2Plot Then
'            iScrolledAmount = (iUBaData - iScrolledAmount) - iCalcdAvailBars2Plot
        ElseIf iScrolledAmount > iUBaData - iCalcdAvailBars2Plot Then
            iScrolledAmount = iUBaData - iCalcdAvailBars2Plot
        End If
        Call ChartBoxDraw
        
        'check if button needs to be dis/enabled
        If iScrolledAmount = 0 And iUBaData - iCalcdAvailBars2Plot <= 0 Then
            tbLeft.Buttons("ScrollData").Enabled = False 'nothing to scroll
        Else
            tbLeft.Buttons("ScrollData").Enabled = True 'need to be able to scroll back
        End If
 
    End If
End Sub
Private Sub CheckForSnapDir()
    Dim sPath As String
    sPath$ = App.Path & "\Snaps"   ' Set the path.
    If Dir(sPath$, vbDirectory) = sEmpty$ Then 'not found... make
        MkDir sPath$
    End If
End Sub
Private Sub GetOptionsDlg()
    frmOptions.Show 1, Me
    Call GetIniSettings  'get any new settings
    Call SetColors
    Call SetMargins
    Call ChartBoxDraw
End Sub

Private Sub GetDataFile()
    Static fIn As Boolean
    If fIn Then Exit Sub  'stop DblClk on the toolbar from bring up the open dlg twice
    fIn = True
    sSymbol$ = sEmpty
    If Not OpenDataFile Then fIn = False: Exit Sub
    Call LoadData
    Call SetMargins
    Call ChartBoxDraw
    fKillSplash = True  'flag to unload splash/progress
    fIn = False  'ok to run again
End Sub


