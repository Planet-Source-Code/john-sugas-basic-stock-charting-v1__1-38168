VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   ClientHeight    =   6030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   402
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   402
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   5100
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   0
      Top             =   5160
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   -480
      Top             =   4800
   End
   Begin VB.Timer tmrDelay 
      Interval        =   1000
      Left            =   120
      Top             =   3660
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_lngRegion As Long, fDoButtonPush As Boolean, fDone As Boolean
Private sInfoStrings() As String, tMemStat As MEMORYSTATUS, tSysInfo As SYSTEM_INFO
Private tOSvInfo As OSVERSIONINFO, fStarted As Boolean, iScreenRgn As Long
Private fCancel As Boolean



Private Sub Form_Load()
    Dim x As Long, y As Long, p As Long, q As Long, iYstep As Long, pt As POINTL
    Call GetSystemInfo(tSysInfo)
    tMemStat.dwLength = Len(tMemStat)
    Call GlobalMemoryStatus(tMemStat)
    Call BuildInfoString
    Call GetCursorPos(pt)
    
    picScreen.Left = Me.Left
    picScreen.Top = Me.Top
    picScreen.Width = Me.Width
    picScreen.Height = Me.Height
    Call RegionFromResource(m_lngRegion, 103, "CUSTOM")
    Apply Me.hWnd, m_lngRegion
    Call RegionFromResource(iScreenRgn, 104, "CUSTOM")
    Apply picScreen.hWnd, iScreenRgn
    Show
    DoEvents
    
    pt.x = 273: pt.y = 320
    Call ClientToScreen(Me.hWnd, pt)
    'x = 584: y = 503
    x = pt.x: y = pt.y
    Call GetCursorPos(pt)
'Debug.Print pt.x; " "; pt.y
    iYstep = (y - pt.y) / ((x - pt.x) / 20) 'get the y steps.... inconsistent....
    q = pt.y
    For p = pt.x To x Step 20
'Debug.Print p; " "; q
        Call SetCursorPos(p, q)
        Delay 0.01
        q = q + iYstep
'Call GetCursorPos(pt)
'Debug.Print pt.x; " "; pt.y
    Next
    Call PositionMousePointer(Me.hWnd, 273, 320) 'make sure the mouse *IS* on the button
    Delay 0.3
    Call Form_MouseDown(1, 0, 273, 320)
    Delay 0.2
    Call Form_MouseUp(1, 0, 273, 320)
    Call DrawLED(6060643, vbGreen)
    Call OutputAboutInfoChars(0.02)

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Dim pt As POINTL
'Call GetCursorPos(pt)
'Debug.Print pt.x; " "; pt.y
'Debug.Print "X:"; x; " Y:"; y
    'if within the drawn button show some action
    If x > 264 And x < 284 Then
        If y > 314 And y < 324 Then
            fDoButtonPush = True   'flag for form paint
            Me.Refresh
            If fStarted Then fCancel = True
        End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If fDoButtonPush Then
        fDoButtonPush = False
        If fStarted Then fDone = True  'set exit flag
        Me.Refresh  'form paint
        If fStarted Then  'exit
            picScreen.Cls  'blank the "screen"
            Delay 0.5  'give a few millisecs to draw things
            Unload Me
        End If
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

'Debug.Print KeyAscii
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub Form_Paint()
    If fDoButtonPush Then  'draw button down , LED off
        Call DrawButtonPush
    Else  'draw button up , LED on if not exiting, and print info
        Call DrawButton
    End If
    If fDone Or Not fStarted Then 'turn off LED
        Call DrawLED(vbBlack, 6060643)
    Else
        Call DrawLED(6060643, vbGreen)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    fCancel = True
    DeleteObject m_lngRegion
    DeleteObject iScreenRgn
    Set frmAbout = Nothing
End Sub
Private Sub DrawButton()
    Me.Line (265, 315)-(280, 316), 8355711
    Me.Line -(279, 322), 8355711
    Me.Line -(264, 321), 8355711
    Me.Line -(265, 315), 8355711
    Me.Line (281, 315)-(280, 323), vbBlack
    Me.Line (282, 315)-(281, 323), vbBlack
    Me.Line (283, 315)-(281, 323), vbBlack
    Me.Line (266, 314)-(281, 315), vbBlack
    Me.PSet (273, 315), vbBlack 'fix the line aliasing
End Sub
Private Sub DrawButtonPush()
    Me.Line (265, 315)-(280, 316), 8355711
    Me.Line -(279, 322), 8355711
    Me.Line -(264, 321), 8355711
    Me.Line -(265, 315), 8355711
    Me.Line (266, 315)-(265, 321), vbBlack
    Me.Line (267, 315)-(266, 321), vbBlack
    Me.Line (266, 320)-(280, 321), vbBlack
    Me.Line (266, 321)-(280, 322), vbBlack
End Sub
Private Sub DrawLED(iColor1 As Long, iColor2 As Long)
    'draw offset first then cover with main circle
    Me.FillStyle = 0
    Me.FillColor = iColor1
    'Circle (179, 236), 4, iColor, , , 0.7
    Circle (250, 316), 4, iColor1
    Me.FillColor = iColor2
    Circle (248, 316), 4, iColor2
    
End Sub
Private Sub BuildInfoString()
    Dim i As Long, dw As Long, sSP As String, sMachine As String, sProcessor As String
    Dim sPlatform As String, sWinVersion As String, sName As String, iLen As Long
    Dim szCSDVersion As String
    
    tOSvInfo.dwOSVersionInfoSize = Len(tOSvInfo)
    Call GetVersionEx(tOSvInfo)  'version & platform info code ported from MSDN Cpp code
    With tOSvInfo
'Debug.Print BytesToStr(.szCSDVersion)
        szCSDVersion = BytesToStr(.szCSDVersion)
        Select Case .dwPlatformId
            Case VER_PLATFORM_WIN32_NT
                '// Test for the product.
                If .dwMajorVersion <= 4 Then
                    sPlatform = "Windows NT"
                ElseIf .dwMajorVersion = 5 And .dwMinorVersion = 0 Then
                    sPlatform = "Windows 2K"
                ElseIf .dwMajorVersion = 5 And .dwMinorVersion = 1 Then
                    sPlatform = "Windows XP"
                End If
                sSP$ = Left$(szCSDVersion, InStr(szCSDVersion, Chr$(0)) - 1)
            Case VER_PLATFORM_WIN32_WINDOWS
    
                If .dwMajorVersion = 4 And .dwMinorVersion = 0 Then _
                    sPlatform = "Windows 95"
                    If Left$(szCSDVersion, 1) = "C" Or _
                        Left$(szCSDVersion, 1) = "B" Then _
                        sSP$ = "OSR2"
                
                If .dwMajorVersion = 4 And .dwMinorVersion = 10 Then _
                    sPlatform = "Windows 98"
                    If Left$(szCSDVersion, 1) = "A" Then sSP$ = "SE"
    
                If .dwMajorVersion = 4 And .dwMinorVersion = 90 Then _
                    sPlatform = "Windows Me"
    
            Case VER_PLATFORM_WIN32s
                sPlatform = "Win32s"
        
        End Select
        sWinVersion$ = .dwMajorVersion & "." & .dwMinorVersion _
                        & "." & .dwBuildNumber & &HFFFF
    End With

    iLen = 16
    sName$ = String$(16, 0)
    If GetComputerName(sName$, iLen) Then sMachine$ = Left$(sName$, iLen)

    sProcessor$ = tSysInfo.dwProcessorType
    
    ReDim sInfoStrings(0 To 8)
    sInfoStrings(0) = "Basic Stock Charting"
    sInfoStrings(1) = "V" & App.Major & "." & App.Minor & ", by John Sugas 2002,"
    sInfoStrings(2) = "jsugas@mei.net"
    sInfoStrings(3) = "Machine: " & sMachine$
    sInfoStrings(4) = sPlatform$ & " Ver: " & sWinVersion$
    sInfoStrings(5) = sSP$
    sInfoStrings(6) = "CPU: " & sProcessor$
    sInfoStrings(7) = "Free: " & tMemStat.dwAvailPhys \ 1000000 & "Mb"
    sInfoStrings(8) = "Total: " & tMemStat.dwTotalPhys \ 1000000 & "Mb"
End Sub

Private Sub OutputAboutInfoChars(Optional rDelay As Single)
    picScreen.Cls
    fStarted = True
    RotateText picScreen, 65, 45, sInfoStrings(0), , True, , 18, -3, -5, vbGreen, , True, rDelay
    If fCancel Then Exit Sub
    RotateText picScreen, 81, 75, sInfoStrings(1), , True, , 12, -3, -5, vbWhite, , True, rDelay
    If fCancel Then Exit Sub
    RotateText picScreen, 111, 95, sInfoStrings(2), , True, , 12, -3, -5, vbWhite, , True, rDelay
    If fCancel Then Exit Sub
    RotateText picScreen, 60, 123, sInfoStrings(3), , True, , 12, -4, -5, vbYellow, , True, rDelay
    If fCancel Then Exit Sub
    RotateText picScreen, 63, 147, sInfoStrings(4), , True, , 12, -4, -5, vbWhite, , True, rDelay
    If fCancel Then Exit Sub
    RotateText picScreen, 103, 167, sInfoStrings(5), , True, , 12, -4, -5, vbWhite, , True, rDelay
    If fCancel Then Exit Sub
    RotateText picScreen, 103, 187, sInfoStrings(6), , True, , 12, -4, -5, vbWhite, , True, rDelay
    If fCancel Then Exit Sub
    RotateText picScreen, 101, 207, sInfoStrings(7), , True, , 12, -4, -5, vbWhite, , True, rDelay
    If fCancel Then Exit Sub
    RotateText picScreen, 100, 225, sInfoStrings(8), , True, , 12, -4, -5, vbWhite, , True, rDelay
End Sub

