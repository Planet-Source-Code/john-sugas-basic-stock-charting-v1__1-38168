VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chart Options"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   252
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   456
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBackup 
      Caption         =   "&Backup INI"
      Height          =   375
      Left            =   180
      TabIndex        =   47
      Top             =   3300
      Width           =   1455
   End
   Begin VB.Timer tmrAfterLoad 
      Interval        =   100
      Left            =   2160
      Top             =   3540
   End
   Begin VB.CommandButton cmdSaveExit 
      Caption         =   "Save && E&xit"
      Height          =   375
      Left            =   3540
      TabIndex        =   46
      Top             =   3300
      Width           =   1455
   End
   Begin VB.CommandButton cmdTakeFocus 
      Caption         =   "Command1"
      Height          =   195
      Left            =   3900
      TabIndex        =   45
      Top             =   -1620
      Width           =   315
   End
   Begin VB.Frame Frame4 
      Caption         =   "Chart Font (Click to Change)"
      Height          =   1395
      Left            =   2220
      TabIndex        =   43
      Top             =   120
      Width           =   2715
      Begin VB.TextBox txtFont 
         Alignment       =   2  'Center
         Height          =   1035
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   44
         Top             =   240
         Width           =   2475
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel && Exit"
      Height          =   375
      Left            =   5220
      TabIndex        =   13
      Top             =   3300
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Misc Settings"
      Height          =   3015
      Left            =   5100
      TabIndex        =   26
      Top             =   120
      Width           =   1575
      Begin VB.TextBox txtScrollIncrement 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   49
         Top             =   2580
         Width           =   615
      End
      Begin VB.TextBox txtCrosshairMode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   32
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtDiv2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   31
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtDiv1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   30
         Top             =   540
         Width           =   615
      End
      Begin VB.Label lblScrollIncrement 
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   855
         TabIndex        =   50
         Top             =   2580
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Scroll Increment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   48
         Top             =   2340
         Width           =   1140
      End
      Begin VB.Label lblCHmode 
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   855
         TabIndex        =   42
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label lblDiv2 
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   855
         TabIndex        =   41
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblDiv1 
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   855
         TabIndex        =   40
         Top             =   540
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "CrossHair Mode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   1680
         Width           =   1125
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Divider2 Position"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Divider1 Position"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   300
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Bar Type"
      Height          =   1515
      Left            =   2220
      TabIndex        =   25
      Top             =   1620
      Width           =   2715
      Begin VB.PictureBox picBarTypeCont 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   60
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   173
         TabIndex        =   33
         Top             =   240
         Width           =   2595
         Begin VB.PictureBox picTypeCandle 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Height          =   540
            Left            =   1860
            Picture         =   "frmOptions.frx":0E42
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   36
            Top             =   360
            Width           =   540
         End
         Begin VB.PictureBox picTypeHLOC 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Height          =   540
            Left            =   1020
            Picture         =   "frmOptions.frx":114C
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   35
            Top             =   360
            Width           =   540
         End
         Begin VB.PictureBox picTypeLine 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Height          =   540
            Left            =   180
            Picture         =   "frmOptions.frx":1456
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   34
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Candle"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   2
            Left            =   1860
            TabIndex        =   39
            Top             =   120
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "HLOC"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   1
            Left            =   1035
            TabIndex        =   38
            Top             =   120
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Line"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   0
            Left            =   255
            TabIndex        =   37
            Top             =   120
            Width           =   330
         End
      End
   End
   Begin VB.CommandButton cmdDefaults 
      Caption         =   "&Reset to Defaults"
      Height          =   375
      Left            =   1860
      TabIndex        =   14
      Top             =   3300
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Color Options"
      Height          =   3015
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   1875
      Begin VB.Label lblColor 
         BorderStyle     =   1  'Fixed Single
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   24
         Top             =   2700
         Width           =   255
      End
      Begin VB.Label lblColor 
         BorderStyle     =   1  'Fixed Single
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   23
         Top             =   2460
         Width           =   255
      End
      Begin VB.Label lblColor 
         BorderStyle     =   1  'Fixed Single
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   22
         Top             =   2220
         Width           =   255
      End
      Begin VB.Label lblColor 
         BorderStyle     =   1  'Fixed Single
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   21
         Top             =   1980
         Width           =   255
      End
      Begin VB.Label lblColor 
         BorderStyle     =   1  'Fixed Single
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   1740
         Width           =   255
      End
      Begin VB.Label lblColor 
         BorderStyle     =   1  'Fixed Single
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   1500
         Width           =   255
      End
      Begin VB.Label lblColor 
         BorderStyle     =   1  'Fixed Single
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   1260
         Width           =   255
      End
      Begin VB.Label lblColor 
         BorderStyle     =   1  'Fixed Single
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   1020
         Width           =   255
      End
      Begin VB.Label lblColor 
         BorderStyle     =   1  'Fixed Single
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   780
         Width           =   255
      End
      Begin VB.Label lblColor 
         BorderStyle     =   1  'Fixed Single
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   540
         Width           =   255
      End
      Begin VB.Label lblColor 
         BorderStyle     =   1  'Fixed Single
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "TicCandleUpColor"
         Height          =   195
         Index           =   10
         Left            =   420
         TabIndex        =   11
         Top             =   2220
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "TicCandleDnColor"
         Height          =   195
         Index           =   9
         Left            =   420
         TabIndex        =   10
         Top             =   2460
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "VolColor"
         Height          =   195
         Index           =   8
         Left            =   480
         TabIndex        =   9
         Top             =   2700
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CrossHairColor"
         Height          =   195
         Index           =   7
         Left            =   420
         TabIndex        =   8
         Top             =   1020
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DateMarkerColor"
         Height          =   195
         Index           =   6
         Left            =   420
         TabIndex        =   7
         Top             =   1260
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "TicBodyColor"
         Height          =   195
         Index           =   5
         Left            =   420
         TabIndex        =   6
         Top             =   1500
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Text/ForeColor"
         Height          =   195
         Index           =   4
         Left            =   420
         TabIndex        =   5
         Top             =   540
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "TicCloseColor"
         Height          =   195
         Index           =   3
         Left            =   420
         TabIndex        =   4
         Top             =   1980
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "TicOpenColor"
         Height          =   195
         Index           =   2
         Left            =   420
         TabIndex        =   3
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "GridColor"
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   2
         Top             =   780
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "BackColor"
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   1
         Top             =   300
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private rc1 As RECT, fBarTypeButtonDn As Boolean


Private Sub cmdBackup_Click()
    FileCopy sINIsetFile$, sINIsetFile$ & ".BAK"
    MsgBox "INI File was BackedUp to .BAK in the App Dir...", vbInformation + vbOKOnly, "INI BackUp Complete"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDefaults_Click()
    Dim iResult As Integer
    iResult = MsgBox("Are you sure you want to reset current settings?" & vbCrLf _
            & "This will be permanent...", vbCritical + vbYesNo + vbDefaultButton2, "Reset to Defaults")
    If iResult = vbNo Then Exit Sub
    Call MakeIniFile
    Call GetIniSettings
    Call GetIndicatorSettings  'reload vars
    Call frmMain.SetColors
    Call frmMain.SetMargins
    Call SetUpFontTb
    Call frmMain.ChartBoxDraw
End Sub

Private Sub cmdSaveExit_Click()
    Call SaveIniSettings
    Unload Me
End Sub

Private Sub Form_Load()
    
    lblColor(0).BackColor = iBackColor
    lblColor(1).BackColor = iForeColor
    lblColor(2).BackColor = iGridColor
    lblColor(3).BackColor = iCrossHairColor
    lblColor(4).BackColor = iDateMarkerColor
    lblColor(5).BackColor = iTicBodyColor
    lblColor(6).BackColor = iTicOpenColor
    lblColor(7).BackColor = iTicCloseColor
    lblColor(8).BackColor = iTicCandleUpColor
    lblColor(9).BackColor = iTicCandleDnColor
    lblColor(10).BackColor = iVolColor
    
    txtCrosshairMode.Text = iCrossHairMode
    txtScrollIncrement = iScrollIncrement
    txtDiv1.Text = rSplit1
    txtDiv2.Text = rSplit2
    
    lblDiv1.Caption = "Min:" & iBottomPlotMargin / 2 & vbCrLf & "Max:" & rSplit2 - 10
    lblDiv2.Caption = "Min:" & rSplit1 + 10 & vbCrLf & "Max:" & iBottomPlotMargin - 50
    lblCHmode.Caption = "Min:1" & vbCrLf & "Max:15"
    lblScrollIncrement.Caption = "Min:10" & vbCrLf & "Max:" & (rRightSideOffset \ iBarSpacing) - 1
    
    Call SetUpFontTb

End Sub
Private Sub DrawTicType()
    picBarTypeCont.Cls
    Call DrawState(picTypeLine.hDC, 0, 0, picTypeLine.Picture, 0, 0, 0, 0, 0, DST_ICON Or DSS_DISABLED)
    Call DrawState(picTypeHLOC.hDC, 0, 0, picTypeHLOC.Picture, 0, 0, 0, 0, 0, DST_ICON Or DSS_DISABLED)
    Call DrawState(picTypeCandle.hDC, 0, 0, picTypeCandle.Picture, 0, 0, 0, 0, 0, DST_ICON Or DSS_DISABLED)

    Select Case iTicType
        Case ttLine
            Call DrawState(picTypeLine.hDC, 0, 0, picTypeLine.Picture, 0, 0, 0, 0, 0, DST_ICON)
            Call DrawHiLiteBox(picTypeLine)
        Case ttHLOC
            Call DrawState(picTypeHLOC.hDC, 0, 0, picTypeHLOC.Picture, 0, 0, 0, 0, 0, DST_ICON)
            Call DrawHiLiteBox(picTypeHLOC)
        Case ttCandle
            Call DrawState(picTypeCandle.hDC, 0, 0, picTypeCandle.Picture, 0, 0, 0, 0, 0, DST_ICON)
            Call DrawHiLiteBox(picTypeCandle)
    End Select
    picTypeLine.Refresh
    picTypeHLOC.Refresh
    picTypeCandle.Refresh
End Sub
Private Sub DrawHiLiteBox(pb As PictureBox)
    Dim iOldPen As Long, iHndPen As Long
    
    iHndPen = CreatePen(PS_Solid, 2, vbRed)
    iOldPen = SelectObject(picBarTypeCont.hDC, iHndPen)

    Call Rectangle(picBarTypeCont.hDC, (pb.Left - 2), _
                            (pb.Top - 2), _
                            (pb.Left + pb.Width + 2), _
                            (pb.Top + pb.Height + 2))
    Call SelectObject(picBarTypeCont.hDC, iOldPen)

End Sub
Private Sub DrawBarTypeButtonUp(pb As PictureBox)
    rc1.Left = pb.Left
    rc1.Top = pb.Top
    rc1.Right = pb.Left + pb.Width
    rc1.Bottom = pb.Top + pb.Height
    
    Call InflateRect(rc1, 5, 5)
    Call DrawEdge(picBarTypeCont.hDC, rc1, EDGE_RAISED, BF_TOPLEFT)
    Call DrawEdge(picBarTypeCont.hDC, rc1, EDGE_RAISED, BF_BOTTOMRIGHT)
    picBarTypeCont.Refresh
    fBarTypeButtonDn = True
End Sub
Private Sub DrawBarTypeButtonDn(pb As PictureBox)
    rc1.Left = pb.Left
    rc1.Top = pb.Top
    rc1.Right = pb.Left + pb.Width
    rc1.Bottom = pb.Top + pb.Height
    
    Call InflateRect(rc1, 5, 5)
    Call DrawEdge(picBarTypeCont.hDC, rc1, EDGE_SUNKEN, BF_TOPLEFT)
    Call DrawEdge(picBarTypeCont.hDC, rc1, EDGE_SUNKEN, BF_BOTTOMRIGHT)
    picBarTypeCont.Refresh
    Delay 0.3
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If fBarTypeButtonDn Then
        picBarTypeCont.Cls
        Call DrawTicType
        fBarTypeButtonDn = False
    End If
End Sub

Private Sub Form_Paint()
    Call DrawTicType
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmOptions = Nothing
End Sub

Private Sub lblColor_Click(Index As Integer)
    
    Select Case Index
        Case 0
            iBackColor = GetColorDlg(iBackColor)
            lblColor(0).BackColor = iBackColor
        Case 1
            iForeColor = GetColorDlg(iForeColor)
            lblColor(1).BackColor = iForeColor
            txtFont.ForeColor = iForeColor
        Case 2
            iGridColor = GetColorDlg(iGridColor)
            lblColor(2).BackColor = iGridColor
        Case 3
            iCrossHairColor = GetColorDlg(iCrossHairColor)
            lblColor(3).BackColor = iCrossHairColor
        Case 4
            iDateMarkerColor = GetColorDlg(iDateMarkerColor)
            lblColor(4).BackColor = iDateMarkerColor
        Case 5
            iTicBodyColor = GetColorDlg(iTicBodyColor)
            lblColor(5).BackColor = iTicBodyColor
        Case 6
            iTicOpenColor = GetColorDlg(iTicOpenColor)
            lblColor(6).BackColor = iTicOpenColor
        Case 7
            iTicCloseColor = GetColorDlg(iTicCloseColor)
            lblColor(7).BackColor = iTicCloseColor
        Case 8
            iTicCandleUpColor = GetColorDlg(iTicCandleUpColor)
            lblColor(8).BackColor = iTicCandleUpColor
        Case 9
            iTicCandleDnColor = GetColorDlg(iTicCandleDnColor)
            lblColor(9).BackColor = iTicCandleDnColor
        Case 10
            iVolColor = GetColorDlg(iVolColor)
            lblColor(10).BackColor = iVolColor
            
    End Select
    Call frmMain.SetColors
    Call frmMain.ChartBoxDraw
End Sub

Private Sub picBarTypeCont_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If fBarTypeButtonDn Then
        picBarTypeCont.Cls
        Call DrawTicType
        fBarTypeButtonDn = False
    End If
End Sub

Private Sub picTypeCandle_Click()
    iTicType = ttCandle
    Call DrawBarTypeButtonDn(picTypeCandle)
    Call DrawTicType
    Call DrawBarTypeButtonUp(picTypeCandle)
    Call frmMain.ChartBoxDraw
End Sub

Private Sub picTypeCandle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If fBarTypeButtonDn Then Exit Sub  'Prevent some flicker
    Call DrawBarTypeButtonUp(picTypeCandle)
End Sub

Private Sub picTypeHLOC_Click()
    iTicType = ttHLOC
    Call DrawBarTypeButtonDn(picTypeHLOC)
    Call DrawTicType
    Call DrawBarTypeButtonUp(picTypeHLOC)
    Call frmMain.ChartBoxDraw
End Sub

Private Sub picTypeHLOC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If fBarTypeButtonDn Then Exit Sub  'Prevent some flicker
    Call DrawBarTypeButtonUp(picTypeHLOC)
End Sub

Private Sub picTypeLine_Click()
    iTicType = ttLine
    Call DrawBarTypeButtonDn(picTypeLine)
    Call DrawTicType
    Call DrawBarTypeButtonUp(picTypeLine)
    Call frmMain.ChartBoxDraw
End Sub

Private Sub picTypeLine_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If fBarTypeButtonDn Then Exit Sub  'Prevent some flicker
    Call DrawBarTypeButtonUp(picTypeLine)
End Sub

Private Sub txtCrosshairMode_Change()
    If Len(txtCrosshairMode.Text) = 2 And txtCrosshairMode.SelLength < 2 Then
        txtCrosshairMode.Text = Right$(txtCrosshairMode.Text, 2) 'don't allow a len value > 2
    End If
    If Val(txtCrosshairMode.Text) > 15 Then
        txtCrosshairMode.Text = 15
    ElseIf Val(txtCrosshairMode.Text) < 1 And Len(txtCrosshairMode.Text) > 0 Then
        txtCrosshairMode.Text = 1
    End If
    
End Sub
Private Sub txtCrosshairMode_Click()
    txtCrosshairMode.SelStart = 0
    txtCrosshairMode.SelLength = Len(txtCrosshairMode.Text)
End Sub
Private Sub txtCrosshairMode_KeyDown(KeyCode As Integer, Shift As Integer)
    
'Debug.Print KeyCode
    Select Case KeyCode
        Case 48 To 57, vbKeyDelete, vbKeyLeft, vbKeyRight, vbKeyBack   ' numerical, delete, r-l arrows, backspace
            'do nothing
        Case vbKeyEscape
            txtCrosshairMode.Text = iCrossHairMode
            txtCrosshairMode.SelStart = Len(txtCrosshairMode.Text)
        Case vbKeyReturn
            If Val(txtCrosshairMode.Text) = 0 Then
                iCrossHairMode = 15
            Else
                iCrossHairMode = Val(txtCrosshairMode.Text)
            End If
        Case Else
            KeyCode = 0
    End Select
End Sub


Private Sub txtDiv1_Click()
    txtDiv1.SelStart = 0
    txtDiv1.SelLength = Len(txtDiv1.Text)
End Sub
Private Sub txtDiv1_KeyDown(KeyCode As Integer, Shift As Integer)

'Debug.Print KeyCode
    Select Case KeyCode
        Case 48 To 57, vbKeyDelete, vbKeyLeft, vbKeyRight, vbKeyBack  ' numerical, delete, r-l arrows, backspace
            'do nothing
        Case vbKeyEscape
            txtDiv1.Text = rSplit1
            txtDiv1.SelStart = Len(txtDiv1.Text)
        Case vbKeyReturn
            If Val(txtDiv1.Text) > 0 Then
                If Val(txtDiv1.Text) > rSplit2 - 10 Then 'set the min & max values
                    txtDiv1.Text = rSplit2 - 10
                ElseIf Val(txtDiv1.Text) < iBottomPlotMargin / 2 And Len(txtDiv1.Text) > 0 Then
                    txtDiv1.Text = iBottomPlotMargin / 2
                End If
                rSplit1 = Val(txtDiv1.Text)
                Call frmMain.SetMargins
                Call frmMain.ChartBoxDraw
            End If
        Case Else
            KeyCode = 0
    End Select
End Sub

Private Sub txtDiv2_Click()
    txtDiv2.SelStart = 0
    txtDiv2.SelLength = Len(txtDiv2.Text)
End Sub
Private Sub txtDiv2_KeyDown(KeyCode As Integer, Shift As Integer)
    'Debug.Print KeyCode
    Select Case KeyCode
        Case 48 To 57, vbKeyDelete, vbKeyLeft, vbKeyRight, vbKeyBack  ' numerical, delete, r-l arrows, backspace
            'do nothing
        Case vbKeyEscape
            txtDiv2.Text = rSplit2
            txtDiv2.SelStart = Len(txtDiv2.Text)
        Case vbKeyReturn
            If Val(txtDiv2.Text) > 0 Then
                If Val(txtDiv2.Text) < rSplit1 + 10 Then 'set the min & max values
                    txtDiv2.Text = rSplit1 + 10
                ElseIf Val(txtDiv2.Text) > iBottomPlotMargin - 50 And Len(txtDiv2.Text) > 0 Then
                    txtDiv2.Text = iBottomPlotMargin - 50
                End If
                rSplit2 = Val(txtDiv2.Text)
                Call frmMain.SetMargins
                Call frmMain.ChartBoxDraw
            End If
        Case Else
            KeyCode = 0
    End Select
End Sub

Private Sub SetUpFontTb()
    
    txtFont.ForeColor = iForeColor
    txtFont.BackColor = iBackColor
    Set txtFont.Font = frmMain.ChartBox.Font
    txtFont.FontBold = iFontBold
    txtFont.FontItalic = iFontItalic
    txtFont.Text = txtFont.FontName & vbCrLf & txtFont.FontSize & " pts " _
                        & IIf(txtFont.FontBold, "Bold ", sEmpty) _
                        & IIf(txtFont.FontItalic, "Italic", sEmpty)
End Sub
Private Sub txtFont_Click()
    Static fIn As Boolean
    If fIn Then Exit Sub 'prevent more than 1 procedure run at a time
    fIn = True
    Dim f As Boolean, fnt As StdFont, clr As Long
    Set fnt = frmMain.ChartBox.Font
    'clr = iForeColor
    fnt.Bold = iFontBold
    fnt.Italic = iFontItalic

    CenterDlgBox 0
    'effects are disabled... don't need underline and strikethough for legion text anyway
    f = VBChooseFont(CurFont:=fnt, _
                         Flags:=CF_BOTH)
    'f = VBChooseFont(CurFont:=fnt, _
                         Color:=clr, _
                         Flags:=CF_EFFECTS Or CF_BOTH)

    If f Then
        Set frmMain.ChartBox.Font = fnt
        sFontName = fnt
        iFontSize = fnt.Size
        'using the font selector only gives 16 colors for text color
        'Note: with black backcolor the text would disappear when changing the font.
        'I twisted my brain trying to figure out why. The font dlg
        'won't show the current color sent to it during init. Always
        'came up 0 (=black) ... eventually I found that it would if one of the 16 colors
        'finally decided to diable the font selector text color color so all colors
        'could be used for the text color
        
        'If clr <> iBackColor Then _
            iForeColor = clr
                
        iFontBold = fnt.Bold
        iFontItalic = fnt.Italic
        Call SetUpFontTb
        Call frmMain.SetColors
        Call frmMain.SetMargins
        Call frmMain.ChartBoxDraw
    End If
    cmdTakeFocus.SetFocus
    fIn = False
End Sub

Private Sub tmrAfterLoad_Timer()
    tmrAfterLoad.Enabled = False
    Call PositionMousePointer(cmdCancel.hWnd, cmdCancel.Width \ 2, cmdCancel.Height / 1.2)
End Sub

Private Sub txtScrollIncrement_Change()
    If Len(txtScrollIncrement.Text) = 2 And txtScrollIncrement.SelLength < 2 Then
        txtScrollIncrement.Text = Right$(txtScrollIncrement.Text, 2) 'don't allow a len value > 2
    End If
    If Val(txtScrollIncrement.Text) > 10 Then
        txtScrollIncrement.Text = 10
    ElseIf Val(txtScrollIncrement.Text) < 10 And Len(txtScrollIncrement.Text) > 0 Then
        txtScrollIncrement.Text = 10
    End If
    
End Sub
Private Sub txtScrollIncrement_Click()
    txtScrollIncrement.SelStart = 0
    txtScrollIncrement.SelLength = Len(txtScrollIncrement.Text)
End Sub
Private Sub txtScrollIncrement_KeyDown(KeyCode As Integer, Shift As Integer)
    
'Debug.Print KeyCode
    Select Case KeyCode
        Case 48 To 57, vbKeyDelete, vbKeyLeft, vbKeyRight, vbKeyBack   ' numerical, delete, r-l arrows, backspace
            'do nothing
        Case vbKeyEscape
            txtScrollIncrement.Text = iCrossHairMode
            txtScrollIncrement.SelStart = Len(txtScrollIncrement.Text)
        Case vbKeyReturn
            If Val(txtScrollIncrement.Text) = 0 Then
                iCrossHairMode = 15
            Else
                iCrossHairMode = Val(txtScrollIncrement.Text)
            End If
        Case Else
            KeyCode = 0
    End Select
End Sub

