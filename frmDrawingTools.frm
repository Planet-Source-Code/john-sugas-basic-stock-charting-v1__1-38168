VERSION 5.00
Begin VB.Form frmDrawingTools 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Drawing Tools"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "frmDrawingTools.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   318
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      Height          =   4425
      Left            =   3240
      ScaleHeight     =   291
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   11
      Top             =   210
      Width           =   2235
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   180
         Picture         =   "frmDrawingTools.frx":0442
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   18
         Top             =   3180
         Width           =   1815
      End
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   180
         Picture         =   "frmDrawingTools.frx":058C
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   17
         Top             =   2580
         Width           =   1815
      End
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   180
         Picture         =   "frmDrawingTools.frx":0796
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   16
         Top             =   1980
         Width           =   1815
      End
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   180
         Picture         =   "frmDrawingTools.frx":09A0
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   15
         Top             =   1380
         Width           =   1815
      End
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   180
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   14
         Top             =   780
         Width           =   1815
         Begin VB.Line Line3 
            BorderWidth     =   2
            X1              =   16
            X2              =   4
            Y1              =   4
            Y2              =   20
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   24
            X2              =   12
            Y1              =   4
            Y2              =   20
         End
      End
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   180
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   13
         Top             =   180
         Width           =   1815
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   24
            X2              =   12
            Y1              =   4
            Y2              =   20
         End
      End
      Begin VB.PictureBox picButton 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   180
         Picture         =   "frmDrawingTools.frx":0BAA
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   12
         Top             =   3780
         Width           =   1815
      End
   End
   Begin VB.Timer tmrAfterLoad 
      Interval        =   100
      Left            =   60
      Top             =   4680
   End
   Begin VB.Frame Frame1 
      Caption         =   "Draw Settings"
      Height          =   4515
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   2835
      Begin VB.OptionButton optCircPt 
         Caption         =   "Opp.Pts"
         Height          =   195
         Index           =   1
         Left            =   1860
         TabIndex        =   31
         Top             =   4140
         Width           =   915
      End
      Begin VB.OptionButton optCircPt 
         Caption         =   "Origin"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   30
         Top             =   4140
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton cmdSetDefault 
         Caption         =   "Set Current as Default"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   60
         TabIndex        =   29
         Top             =   3975
         Width           =   915
      End
      Begin VB.CheckBox chkExtendLeft 
         Height          =   315
         Left            =   1500
         TabIndex        =   27
         Top             =   2700
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox chkCircle 
         Height          =   195
         Left            =   2340
         TabIndex        =   25
         Top             =   3660
         Width           =   195
      End
      Begin VB.ComboBox cboFillStyle 
         Height          =   315
         ItemData        =   "frmDrawingTools.frx":0EB4
         Left            =   1080
         List            =   "frmDrawingTools.frx":0ED0
         TabIndex        =   21
         Top             =   2220
         Width           =   1515
      End
      Begin VB.CheckBox chkSquare 
         Height          =   195
         Left            =   2340
         TabIndex        =   20
         Top             =   3240
         Width           =   195
      End
      Begin VB.ComboBox cboMode 
         Height          =   315
         ItemData        =   "frmDrawingTools.frx":0F3E
         Left            =   1080
         List            =   "frmDrawingTools.frx":0F72
         TabIndex        =   10
         Top             =   1740
         Width           =   1515
      End
      Begin VB.ComboBox cboWidth 
         Height          =   315
         ItemData        =   "frmDrawingTools.frx":101F
         Left            =   1080
         List            =   "frmDrawingTools.frx":1032
         TabIndex        =   9
         Top             =   1260
         Width           =   1515
      End
      Begin VB.ComboBox cboStyle 
         Height          =   315
         ItemData        =   "frmDrawingTools.frx":1045
         Left            =   1080
         List            =   "frmDrawingTools.frx":1058
         TabIndex        =   8
         Top             =   780
         Width           =   1515
      End
      Begin VB.CheckBox chkExtend 
         Height          =   315
         Left            =   2340
         TabIndex        =   7
         Top             =   2700
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Right:"
         Height          =   195
         Index           =   6
         Left            =   1860
         TabIndex        =   32
         Top             =   2760
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Left:"
         Height          =   195
         Index           =   10
         Left            =   1140
         TabIndex        =   28
         Top             =   2760
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Perfect Circle:"
         Height          =   195
         Index           =   9
         Left            =   1260
         TabIndex        =   26
         Top             =   3660
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FillColor:"
         Height          =   195
         Index           =   8
         Left            =   1560
         TabIndex        =   24
         Top             =   420
         Width           =   585
      End
      Begin VB.Label lblFillColor 
         BackColor       =   &H00AE480B&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2220
         TabIndex        =   23
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FillStyle:"
         Height          =   195
         Index           =   7
         Left            =   420
         TabIndex        =   22
         Top             =   2280
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Perfect Square:"
         Height          =   195
         Index           =   5
         Left            =   1140
         TabIndex        =   19
         Top             =   3240
         Width           =   1110
      End
      Begin VB.Label lblColor 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "TL Extend"
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
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DrawMode:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   4
         Top             =   1800
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DrawWidth:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   1320
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DrawStyle:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DrawColor:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   420
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmDrawingTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private tButton() As PbButtonSpecs, fMoved As Boolean, iNumSettings As Long
Private iPbTextHeight As Long, afBorder As Long, afBorderT As Long, afStyle As Long

Private Sub cboStyle_Click()
    If cboStyle.ListIndex = 0 Then
        cboWidth.Enabled = True
    Else
        cboWidth.ListIndex = 0
        cboWidth.Enabled = False
    End If

End Sub

Private Sub Form_Load()
    Dim i As Long
    
    ReDim tButton(0 To picButton.UBound)
    For i = 0 To picButton.UBound
        With picButton(i)
            tButton(i).recButton.Left = .Left
            tButton(i).recButton.Top = .Top
            tButton(i).recButton.Right = .Left + .Width
            tButton(i).recButton.Bottom = .Top + .Height
        End With
        Call InflateRect(tButton(i).recButton, 3, 3)
    Next
    iPbTextHeight = picButton(0).TextHeight("X")
    afBorder = BDR_RAISEDOUTER Or BDR_RAISEDINNER
    afStyle = BF_RECT Or BF_MIDDLE
    
    tButton(0).sCaption = "Cancel & Exit"
    tButton(1).sCaption = "TrendLine"
    tButton(2).sCaption = "Parallel TL"
    tButton(3).sCaption = "Elipse"
    tButton(4).sCaption = "Rectangle"
    tButton(5).sCaption = "Fib Retrace"
    tButton(6).sCaption = "(For Future Use)"
    
    tButton(0).iCaptionX = 35
    tButton(1).iCaptionX = 40
    tButton(2).iCaptionX = 40
    tButton(3).iCaptionX = 50
    tButton(4).iCaptionX = 40
    tButton(5).iCaptionX = 35
    tButton(6).iCaptionX = 20
    
    cboStyle.ListIndex = 0
    cboWidth.ListIndex = 0
    cboMode.ListIndex = 12
    cboFillStyle.ListIndex = 0
    
    iNumSettings = GetNumIniKeys(sINIsetFile$, "DrawingToolDefaults")
    If iNumSettings <> 0 Then
        Call GetDrawToolSettings
    End If

End Sub

Private Sub Form_Paint()
    Call DrawButtons
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmDrawingTools = Nothing
End Sub

Private Sub lblColor_Click()
    lblColor.BackColor = GetColorDlg(lblColor.BackColor)
End Sub

Private Sub lblFillColor_Click()
    lblFillColor.BackColor = GetColorDlg(lblFillColor.BackColor)
End Sub

Private Sub picButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ToggleButton Index, False
    picButton(Index).Move picButton(Index).Left + 1, picButton(Index).Top + 1
    fMoved = True
End Sub

Private Sub picButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ToggleButton Index, True
    'if the pic didn't move on the mousedown then we don't what to move it here
    If fMoved Then picButton(Index).Move picButton(Index).Left - 1, picButton(Index).Top - 1
    fMoved = False
    objDrawingTools.ToolColor = lblColor.BackColor
    objDrawingTools.ToolMode = cboMode.ListIndex + 1
    objDrawingTools.ToolStyle = cboStyle.ListIndex
    objDrawingTools.ToolWidth = cboWidth.ListIndex + 1
    'Extend has 4 possible states, 0= no extend, 1= ext.Right only,
    '2=Ext.Left only, 3=Ext.Both
    objDrawingTools.Extend = Abs(chkExtend.Value) + (Abs(chkExtendLeft.Value) * 2)
    objDrawingTools.ToolFillStyle = cboFillStyle.ListIndex
    objDrawingTools.ToolFillColor = lblFillColor.BackColor
    objDrawingTools.UseOrigin = optCircPt(0).Value
    Me.Hide
    Select Case Index
        Case 0  'exit
            Unload Me
        Case 1  'single trendline
            objDrawingTools.TrendLine
        Case 2  'parallel trendlines
            objDrawingTools.TrendLine (True)
        Case 3  'elipse-circle
            objDrawingTools.CircleElipseTool (chkCircle.Value)
        Case 4  'rect-square
            objDrawingTools.RectAndSquareTool (chkSquare.Value)
        Case 5 'fib retracement
            objDrawingTools.FibRetrace
        Case 6
            
    End Select
    Unload Me
End Sub
Private Sub ToggleButton(Index As Integer, fUp As Boolean)
    'this sub borrowed from hardcore vb... modified a little bit....
    If fUp Then
        afBorder = afBorderT
    Else
        afBorderT = afBorder
        afBorder = (Not afBorder) And &HF
    End If
    Call DrawButtons(Index)
End Sub
Private Sub DrawButtons(Optional Index As Integer = -1)
    Dim i As Long
    For i = 0 To UBound(tButton())
        If Index <> -1 Then i = Index 'only draw one button
        Call DrawEdge(picContainer.hDC, tButton(i).recButton, afBorder, afStyle)
        picContainer.Refresh
        picButton(i).CurrentX = tButton(i).iCaptionX   'picButton(i).Picture.Width \ Screen.TwipsPerPixelX
'Debug.Print picButton(i).Picture.Width \ Screen.TwipsPerPixelX
        picButton(i).CurrentY = (picButton(i).Height - picButton(i).TextHeight(tButton(i).sCaption)) \ 2
        picButton(i).Print tButton(i).sCaption
        If Index <> -1 Then Exit For  'done with the one button so exit
    Next
    
End Sub

Private Sub tmrAfterLoad_Timer()
    tmrAfterLoad.Enabled = False
    Call PositionMousePointer(picButton(0).hWnd, picButton(0).Left, picButton(0).Height / 1.2, True)
End Sub
Private Sub cmdSetDefault_Click()
    iNumSettings = GetNumIniKeys(sINIsetFile$, "DrawingToolDefaults")
    If iNumSettings = 0 Then
        Open sINIsetFile$ For Append Access Write As #1
            Print #1, "[DrawingToolDefaults]"
            Print #1, "DrawColor="
            Print #1, "DrawStyle="
            Print #1, "DrawWidth="
            Print #1, "DrawMode="
            Print #1, "FillColor="
            Print #1, "FillStyle="
            Print #1, "TLExtRight="
            Print #1, "TLExtLeft="
            Print #1, "PerfSqr="
            Print #1, "PerfCirc="
            Print #1, "UseOrigin="
            Print #1, sEmpty
        Close #1
    End If
    WriteIni sINIsetFile, "DrawingToolDefaults", "DrawColor", CStr(lblColor.BackColor)
    WriteIni sINIsetFile, "DrawingToolDefaults", "DrawStyle", CStr(cboStyle.ListIndex)
    WriteIni sINIsetFile, "DrawingToolDefaults", "DrawWidth", CStr(cboWidth.ListIndex)
    WriteIni sINIsetFile, "DrawingToolDefaults", "DrawMode", CStr(cboMode.ListIndex)
    WriteIni sINIsetFile, "DrawingToolDefaults", "FillColor", CStr(lblFillColor.BackColor)
    WriteIni sINIsetFile, "DrawingToolDefaults", "FillStyle", CStr(cboFillStyle.ListIndex)
    WriteIni sINIsetFile, "DrawingToolDefaults", "TLExtRight", CStr(chkExtend.Value)
    WriteIni sINIsetFile, "DrawingToolDefaults", "TLExtLeft", CStr(chkExtendLeft.Value)
    WriteIni sINIsetFile, "DrawingToolDefaults", "PerfSqr", CStr(chkSquare.Value)
    WriteIni sINIsetFile, "DrawingToolDefaults", "PerfCirc", CStr(chkCircle.Value)
    WriteIni sINIsetFile, "DrawingToolDefaults", "UseOrigin", CStr(optCircPt(0).Value)
    MsgBox "Current Settings have been saved as Defaults..." & vbCrLf _
            & "If the main program defaults are ever reset," & vbCrLf _
            & "these settings will be erased....", vbInformation + vbOKOnly, "Successful Save"
End Sub
Private Sub GetDrawToolSettings()
    lblColor.BackColor = Val(GetIni(sINIsetFile, "DrawingToolDefaults", "DrawColor"))
    cboStyle.ListIndex = Val(GetIni(sINIsetFile, "DrawingToolDefaults", "DrawStyle"))
    cboWidth.ListIndex = Val(GetIni(sINIsetFile, "DrawingToolDefaults", "DrawWidth"))
    cboMode.ListIndex = Val(GetIni(sINIsetFile, "DrawingToolDefaults", "DrawMode"))
    lblFillColor.BackColor = Val(GetIni(sINIsetFile, "DrawingToolDefaults", "FillColor"))
    cboFillStyle.ListIndex = Val(GetIni(sINIsetFile, "DrawingToolDefaults", "FillStyle"))
    chkExtend.Value = Val(GetIni(sINIsetFile, "DrawingToolDefaults", "TLExtRight"))
    chkExtendLeft.Value = Val(GetIni(sINIsetFile, "DrawingToolDefaults", "TLExtLeft"))
    chkSquare.Value = Val(GetIni(sINIsetFile, "DrawingToolDefaults", "PerfSqr"))
    chkCircle.Value = Val(GetIni(sINIsetFile, "DrawingToolDefaults", "PerfCirc"))
    optCircPt(0).Value = CBool(GetIni(sINIsetFile, "DrawingToolDefaults", "UseOrigin"))
    If optCircPt(0).Value = False Then optCircPt(1).Value = True
End Sub


