VERSION 5.00
Begin VB.Form frmDownLoad 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data DownLoader"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   Icon            =   "frmDownLoad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BasicStockCharting.DatePicker DatePicker1 
      Height          =   2730
      Left            =   2520
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3660
      Visible         =   0   'False
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   4815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Data Dir"
      ForeColor       =   &H00000000&
      Height          =   530
      Index           =   4
      Left            =   120
      TabIndex        =   25
      Top             =   1560
      Width           =   5535
      Begin VB.CommandButton cmdChangeDir 
         Caption         =   "..."
         Height          =   255
         Left            =   5220
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Change Dir"
         Top             =   180
         Width           =   255
      End
      Begin VB.Label lblDir 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   60
         TabIndex        =   27
         Top             =   180
         Width           =   5160
      End
   End
   Begin VB.Timer tmrProgress 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1920
      Top             =   4020
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Status"
      ForeColor       =   &H00000000&
      Height          =   1395
      Index           =   2
      Left            =   2220
      TabIndex        =   20
      Top             =   2220
      Width           =   1635
      Begin VB.PictureBox picProgressV 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   1035
         Left            =   180
         ScaleHeight     =   975
         ScaleWidth      =   1215
         TabIndex        =   28
         Top             =   240
         Width           =   1275
         Begin VB.Label Label2 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Conn. Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   120
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "LAN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   5
            Left            =   450
            TabIndex        =   31
            Top             =   660
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Modem"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   4
            Left            =   450
            TabIndex        =   30
            Top             =   405
            Width           =   615
         End
         Begin VB.Shape shpLAN 
            BackColor       =   &H00000000&
            BorderColor     =   &H000000FF&
            FillColor       =   &H00404040&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   210
            Shape           =   3  'Circle
            Top             =   690
            Width           =   135
         End
         Begin VB.Shape shpModem 
            BackColor       =   &H00000000&
            BorderColor     =   &H000000FF&
            FillColor       =   &H00404040&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   210
            Shape           =   3  'Circle
            Top             =   450
            Width           =   135
         End
      End
      Begin VB.PictureBox picProgress 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   1035
         Left            =   180
         ScaleHeight     =   975
         ScaleWidth      =   1215
         TabIndex        =   33
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Timer tmrAfterLoad 
      Interval        =   100
      Left            =   0
      Top             =   3900
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Action"
      ForeColor       =   &H00000000&
      Height          =   1395
      Index           =   3
      Left            =   4080
      TabIndex        =   17
      Top             =   2220
      Width           =   1575
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "(future use)"
         Height          =   435
         Left            =   197
         TabIndex        =   24
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdGetTheData 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Get The Data"
         Height          =   435
         Left            =   197
         TabIndex        =   1
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Data Source"
      ForeColor       =   &H00000000&
      Height          =   1395
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   2220
      Width           =   1875
      Begin VB.OptionButton optSource 
         BackColor       =   &H00C0C0C0&
         Caption         =   "..."
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton optSource 
         BackColor       =   &H00C0C0C0&
         Caption         =   "..."
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   660
         Width           =   855
      End
      Begin VB.OptionButton optSource 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Yahoo"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "URL Construction"
      ForeColor       =   &H00000000&
      Height          =   1395
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   5535
      Begin VB.TextBox txtURL 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   995
         Width           =   5295
      End
      Begin VB.OptionButton optPeriod 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Weekly"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   4500
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdSelectDate 
         Caption         =   "ED"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3720
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton cmdSelectDate 
         Caption         =   "BD"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2040
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   300
         Width           =   375
      End
      Begin VB.OptionButton optPeriod 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Daily"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   4500
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   300
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtEndYear 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   540
         Width           =   615
      End
      Begin VB.TextBox txtEndDay 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3180
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   540
         Width           =   375
      End
      Begin VB.TextBox txtEndMonth 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   540
         Width           =   375
      End
      Begin VB.TextBox txtBeginYear 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   540
         Width           =   615
      End
      Begin VB.TextBox txtBeginDay 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   540
         Width           =   375
      End
      Begin VB.TextBox txtBeginMonth 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   540
         Width           =   375
      End
      Begin VB.TextBox txtSymbol 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Top             =   540
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "EndDate:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   2880
         TabIndex        =   6
         Top             =   300
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "BeginDate:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   1140
         TabIndex        =   5
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Symbol:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   5535
   End
End
Attribute VB_Name = "frmDownLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Yahoo format
'"http://table.finance.yahoo.com/table.csv?a=5&b=10&c=2000&d=8&e=11&f=2002&s=msft&y=0&g=d&ignore=.csv"



Private sURLcurrent As String, sData As String, iWhichDate As Long
Private sFileSaveName As String, sURLbase As String, iSource As Long
Private iPeriod As Long, sPeriod As String, fCancel As Boolean


Private Sub cmdChangeDir_Click()
    Dim s As String
    s$ = BrowseForFolder(0, "Select Data Dir", sDataDir$)
    If s$ = sEmpty Then Exit Sub
    sDataDir$ = s$
    Call WriteIni(sINIsetFile, "DLSettings", "DataDir", sDataDir$)
    lblDir.Caption = sDataDir$
End Sub

Private Sub cmdGetTheData_Click()
    '"http://table.finance.yahoo.com/table.csv?a=5&b=10&c=2000&d=8&e=11&f=2002&s=msft&y=0&g=d&ignore=.csv"

    If Not Online() Then _
            Call MsgBox("No Connection", vbCritical + vbOKOnly, "No Connection"): Exit Sub
    If txtSymbol.Text = sEmpty Then lblStatus.Caption = "No Symbol.. Abort Op": Exit Sub

    Screen.MousePointer = vbHourglass
    tmrProgress.Enabled = True
    Call ConstructURL
    sData$ = GetFromInet(sURLcurrent$)
    Call ParseAndSaveData
    tmrProgress.Enabled = False
    Screen.MousePointer = vbDefault
    picProgress.Cls

End Sub



Private Sub cmdSelectDate_Click(Index As Integer)
    
    If Index = 0 Then  'begin date
        iWhichDate = 0
        DatePicker1.InitDate = (Month(Date) - 1 & "/" & Day(Date) & "/" & Year(Date) - 1)
                            '(DatePart("m", Now) & "/" & _
                            DatePart("d", Now) & "/" & DatePart("yyyy", Now) - 1)
    Else
        DatePicker1.InitDate = Date
        iWhichDate = 1
    End If
    DatePicker1.Left = 1560
    DatePicker1.Top = 1050
    DatePicker1.Visible = True
    
End Sub

Private Sub DatePicker1_Cancel()
    DatePicker1.Visible = False
End Sub

Private Sub DatePicker1_OK(ReturnDate As Date)
    DatePicker1.Visible = False
    If iWhichDate = 0 Then 'begin date
        txtBeginMonth.Text = Format(ReturnDate, "mm")
        txtBeginDay.Text = Format(ReturnDate, "dd")
        txtBeginYear.Text = Format(ReturnDate, "yyyy")
    Else
        txtEndMonth.Text = Format(ReturnDate, "mm")
        txtEndDay.Text = Format(ReturnDate, "dd")
        txtEndYear.Text = Format(ReturnDate, "yyyy")
    End If
End Sub

Private Sub Form_Load()

    sDataDir$ = GetIni(sINIsetFile, "DLSettings", "DataDir")
    If Left$(sDataDir$, 1) = "\" Then sDataDir$ = App.Path & sDataDir$
    If Dir(sDataDir$, vbDirectory) = sEmpty$ Then 'not found... make
        MkDir sDataDir$
    End If
    sURLcurrent$ = GetIni(sINIsetFile, "DLSettings", "LastURL")
    lblDir.Caption = sDataDir$
    txtBeginMonth.Text = Format(Now, "mm")
    txtBeginDay.Text = Format(Now, "dd")
    txtBeginYear.Text = Format(Now, "yyyy") - 1
    txtEndMonth.Text = Format(Now, "mm")
    txtEndDay.Text = Format(Now, "dd")
    txtEndYear.Text = Format(Now, "yyyy")
    
    iSource = Val(GetIni(sINIsetFile, "DLSettings", "Source"))
    optSource(iSource).Value = True
    Select Case iSource
        Case 0  'yahoo
            sURLbase$ = "http://table.finance.yahoo.com/table.csv?"
        Case 1
        
    End Select
    

    If ViaLAN() Then shpLAN.FillColor = vbGreen
    If ViaModem() Then shpModem.FillColor = vbGreen

End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrProgress.Enabled = False
    fCancel = True 'get us out of the progress loop if running
    Call WriteIni(sINIsetFile, "DLSettings", "DataDir", sDataDir$)
    Call WriteIni(sINIsetFile, "DLSettings", "LastURL", sURLcurrent$)
    Call WriteIni(sINIsetFile, "DLSettings", "Source", CStr(iSource))
    Set frmDownLoad = Nothing
End Sub
Private Sub ConstructURL()
    Dim sTemp As String
    
    Select Case iSource
        Case 0  'yahoo
            'a=5&b=10&c=2000&d=8&e=11&f=2002&s=msft&y=0&g=d&ignore=.csv"
            sTemp$ = "a=" & txtBeginMonth.Text & "&b=" & txtBeginDay.Text & "&c=" & txtBeginYear.Text
            sTemp$ = sTemp$ & "&d=" & txtEndMonth.Text & "&e=" & txtEndDay.Text & "&f=" & txtEndYear.Text
            sTemp$ = sTemp$ & "&s=" & LCase$(txtSymbol.Text) & "&y=0&g="
            Select Case iPeriod
                Case 0  'daily
                    sPeriod$ = "d"
                Case 1  'weekly
                    sPeriod$ = "w"
            End Select
            sTemp$ = sTemp$ & sPeriod$ & "&ignore=.csv"
    End Select
    sURLcurrent$ = sURLbase$ & sTemp$
    txtURL.Text = sURLcurrent$
End Sub
Private Sub ParseAndSaveData()
    Dim iFile As Integer, iPos As Long, sLine As String, sFirstLine As String
    Dim sTemp As String, iLineCount As Long, sFormat As String, sPath As String
    
    If Len(sData$) < 20 Then
        lblStatus.Caption = "Length of data < 20..."
        Exit Sub
    End If
    sFileSaveName$ = txtSymbol.Text & "~" & sPeriod & "-" & Format(Date, "mmddyyyy") & ".dat"
    sPath$ = sDataDir$ & "\" & sFileSaveName$
    iFile = FreeFile
    
    Select Case iSource
        Case 0  'yahoo
            iPos = InStr(sData$, "Date")  'dump everything before "Date"
            If iPos = 0 Then lblStatus.Caption = "Error with Data, ""Date"" not found": Exit Sub
            sData$ = Mid$(sData$, iPos)
            
            sData$ = Replace(sData$, Chr$(10), vbCrLf) 'give us separate lines
            Open sPath$ For Output Access Write Lock Write As #iFile
                Print #iFile, sData$
            Close #iFile
            
            lblStatus.Caption = "Parsing File..."
            sData$ = ""  'empty data string
            'tested faster to open the file and get each line at a time then to parse
            'the original string when replacing the dates
            Open sPath$ For Input Access Read As #iFile
            Do While Not EOF(iFile)
                DoEvents
                Line Input #iFile, sLine$
                iLineCount = iLineCount + 1
                If Len(sLine$) > 2 Then
                    iPos = InStr(sLine$, ",")
                        If iPos <> 0 Then
                        sTemp$ = Mid$(sLine$, 1, iPos - 1)  'get the first token... it is the date
                        If IsDate(sTemp$) Then  'make sure it is a date
                            sFormat$ = Format(sTemp$, "mm/dd/yyyy")  'better format than original
                            sLine$ = Replace(sLine$, sTemp$, sFormat$) 'replace it
                        End If
                        'build new file with temp string. Reverse the order the so an update
                        'only needs an append.  The chart data loader expects it that way also.
                        If iLineCount = 1 Then 'not the first line
                            sFirstLine$ = sLine$ 'first line is the format header save till later
                        ElseIf iLineCount = 2 Then
                            sData$ = sLine$
                        Else
                            sData$ = sLine$ & vbCrLf & sData$
                        End If
                    End If
                End If
            Loop
            sData$ = sFirstLine$ & vbCrLf & sData$  'put at the head of the file
            Close #iFile
        
        Case 1
        
    End Select
    
    Open sPath$ For Output Access Write Lock Write As #iFile
        Print #iFile, sData$  'save the formatted data
    Close #iFile
    lblStatus.Caption = "Operation Complete"
End Sub

Private Sub optPeriod_Click(Index As Integer)
    iPeriod = Index
End Sub

Private Sub tmrAfterLoad_Timer()
    tmrAfterLoad.Enabled = False
    Call PositionMousePointer(Me.hWnd, Me.Width \ 2, Me.Height / 2, False)
End Sub

Private Sub tmrProgress_Timer()
    Dim i As Long, iColor As Long, x As Long, y As Long, fIn As Boolean, j As Long
    If fIn Then Exit Sub
    fIn = True
    x = picProgress.ScaleWidth \ 2
    y = picProgress.ScaleHeight \ 2
    For i = 1 To 120  '70
        'If i = 120 Then DoEvents
        If fCancel Then Exit For
        j = (i \ 10)
        If j < 1 Then j = 1
        picProgress.DrawWidth = j
        iColor = RGB(0, 255 - i * 2, 0)
        picProgress.FillColor = iColor
        picProgress.Circle (x, y), i * 10, vbGreen
        picProgress.DrawWidth = j
        If i > 24 Then picProgress.Circle (x, y), (i - 25) * 10 + 1, RGB(0, 255 - i, 0)
        If i > 54 Then picProgress.Circle (x, y), (i - 55) * 10 + 1, RGB(0, 255 - i - j * 5, 0)
        'picProgressV.Picture = picProgress.Image
        Call BitBlt(picProgressV.hDC, 0, 0, _
                    picProgressV.ScaleWidth \ Screen.TwipsPerPixelX, _
                    picProgressV.ScaleHeight \ Screen.TwipsPerPixelY, _
                    picProgress.hDC, 0, 0, SRCCOPY)
        
        picProgressV.Refresh
        Delay 0.05
    Next
    fIn = False
End Sub

Private Sub txtSymbol_Change()
    txtSymbol.Text = UCase(txtSymbol.Text)
    txtSymbol.SelStart = Len(txtSymbol.Text)
    Call ConstructURL
End Sub
