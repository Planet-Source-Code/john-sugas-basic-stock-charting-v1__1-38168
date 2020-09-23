VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   5190
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrProgress 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   300
      Top             =   2340
   End
   Begin VB.Timer tmrUnload 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   300
      Top             =   1800
   End
   Begin VB.Label lblTickerTape 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   4815
      Width           =   7635
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private m_lngRegion As Long, iLen As Long, sLoadingString As String, i As Long
Private sUseStr As String, fDoUnload As Boolean

Private Sub Form_DblClick()
    Unload Me
End Sub

Private Sub Form_Load()
    If sINIsetFile$ = sEmpty Then  'program start-up
        Call RegionFromResource(m_lngRegion, 101, "CUSTOM")
        Apply Me.hWnd, m_lngRegion
        tmrUnload.Enabled = True
    Else
        sLoadingString$ = "LOADING DATA..... PLEASE WAIT..... "
        sUseStr$ = Space$(50)
        iLen = Len(sLoadingString$)
        lblTickerTape.BackColor = iBackColor
        lblTickerTape.ForeColor = iForeColor
        Call RegionFromResource(m_lngRegion, 102, "CUSTOM")
        Apply Me.hWnd, m_lngRegion
        tmrProgress.Enabled = True
        tmrUnload.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
            
    If iLen <> 0 Then
        Do While Not fDoUnload And Not fKillSplash 'let timer expire
            DoEvents
        Loop
    End If
    DeleteObject m_lngRegion
    Set frmSplash = Nothing
    fKillSplash = False
End Sub

Private Sub tmrProgress_Timer()
    i = i + 1 'counter
    'subtract 1 from beginning and add 1 to end
    sUseStr$ = Mid$(sUseStr$, 2) & Mid$(sLoadingString$, i, 1)
    'need to grab only amount we need or the label ctrl will draw "chunks"...
    'meaning whole words instead of separate letters
    lblTickerTape.Caption = Left$(sUseStr$, 50)
    If i = iLen Then i = 0 'end of message str. start over
End Sub

Private Sub tmrUnload_Timer()
    'if using as progress ctrl set flag & exit
    fDoUnload = True
    'If iLen = 0 Then Unload Me
    Unload Me
End Sub

