VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl DatePicker 
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2730
   ScaleHeight     =   2730
   ScaleWidth      =   2730
   ToolboxBitmap   =   "DatePicker.ctx":0000
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   375
      Left            =   17
      ScaleHeight     =   315
      ScaleWidth      =   2640
      TabIndex        =   1
      Top             =   2340
      Width           =   2705
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   1500
         TabIndex        =   3
         Top             =   35
         Width           =   915
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   35
         Width           =   915
      End
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   1
      StartOfWeek     =   22806529
      CurrentDate     =   37481
   End
End
Attribute VB_Name = "DatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event OK(ReturnDate As Date)
Event Cancel()


Private Sub cmdcancel_Click()
    RaiseEvent Cancel
End Sub

Private Sub cmdOK_Click()
    RaiseEvent OK(Format(MonthView1.Value(), "mm/dd/yyyy"))
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    Picture1.SetFocus
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 2730
    UserControl.Width = 2730
End Sub


Public Property Let InitDate(ByVal New_InitDate As Date)
Attribute InitDate.VB_Description = "Returns/sets the currently selected date."
    MonthView1.Value() = Format(New_InitDate, "mm/dd/yyyy")
End Property

