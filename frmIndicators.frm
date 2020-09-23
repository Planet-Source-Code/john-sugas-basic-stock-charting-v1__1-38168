VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIndicators 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Indicators & Options"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   Icon            =   "frmIndicators.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrAfterLoad 
      Interval        =   100
      Left            =   900
      Top             =   3660
   End
   Begin VB.CommandButton cmdType 
      Caption         =   "Averages"
      Height          =   435
      Left            =   5880
      TabIndex        =   8
      Top             =   300
      Width           =   1095
   End
   Begin VB.PictureBox picColor 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   3.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   4680
      ScaleHeight     =   135
      ScaleWidth      =   495
      TabIndex        =   7
      ToolTipText     =   "Dbl Click to Edit"
      Top             =   3720
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "&Undo Edit"
      Enabled         =   0   'False
      Height          =   435
      Left            =   5880
      TabIndex        =   6
      Top             =   1620
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel && Exit"
      Height          =   435
      Left            =   5880
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save && E&xit"
      Height          =   435
      Left            =   5880
      TabIndex        =   4
      Top             =   2580
      Width           =   1095
   End
   Begin VB.CommandButton cmdChangeInd 
      Caption         =   "C&hange"
      Height          =   435
      Left            =   5880
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtEdit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Top             =   3720
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Indicator Settings"
      Height          =   3495
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   5535
      Begin MSComctlLib.ListView lstVwSettings 
         Height          =   2955
         Left            =   180
         TabIndex        =   1
         ToolTipText     =   "Dbl Click Edit"
         Top             =   360
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5212
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Parameter"
            Text            =   "Parameter"
            Object.Width           =   5010
         EndProperty
      End
   End
End
Attribute VB_Name = "frmIndicators"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const sIndicators = "Indicators"
Private Const sParameters = "Parameters"
Private Const sAverages = "Averages"

Private oClickedItem As MSComctlLib.ListItem, oLastEditItem As MSComctlLib.ListItem
Private sSection As String, bDirtyFlag As Boolean, bSelectingInd As Boolean
Private sFrameCaption As String, sOrgInd As String, fButtonRightSide As Boolean
Private sPrevInd As String, bBooleanEdit As Boolean

Private Sub cmdCancel_Click()
    Call ResetValues
    Unload Me
End Sub
Private Sub ResetValues()
    'reset to original values when first loaded
    Dim i As Long, p As Long
    If Not bSelectingInd Then  'settings page
        For i = 1 To lstVwSettings.ListItems.Count
            If lstVwSettings.ListItems(i).Text <> sEmpty Then
                p = InStr(lstVwSettings.ListItems(i).Tag, "|") 'dump any pb index data
                If p = 0 Then
                    Call WriteIni(sINIsetFile$, sSection$, lstVwSettings.ListItems(i).Text, lstVwSettings.ListItems(i).Tag)
                Else
                    Call WriteIni(sINIsetFile$, sSection$, lstVwSettings.ListItems(i).Text, Left$(lstVwSettings.ListItems(i).Tag, p - 1))
                End If
            End If
        Next
    End If
    'reset to org indicator
    CurrentIndicator$ = sOrgInd$
    Call WriteIni(sINIsetFile$, "Settings", "CurrentIndicator", CurrentIndicator$)
    bDirtyFlag = False
    Call UpdateChart
End Sub
Private Sub UpdateChart()
    Call GetIndicatorSettings  'reload vars
    Call frmMain.ChartBoxDraw
End Sub
Private Sub cmdType_Click()
    
    If cmdType.Caption = sIndicators$ Then 'bring up the Indicators
        bSelectingInd = False
        Call cmdChangeInd_Click
        cmdType.Caption = sAverages$
        cmdChangeInd.Enabled = True
    ElseIf cmdType.Caption = sAverages$ Then 'bring up the Averages
        Dim i As Long
        Call DirtyFlagRoutine 'do we need to save any changes
        'we need to reload picColor ctrls so unload now
        For i = picColor.LBound + 1 To picColor.UBound
            Unload picColor(i)
        Next
        Call LoadCurrentSettings(sAverages$)
        cmdType.Caption = sIndicators$
        cmdChangeInd.Enabled = False
    End If
End Sub
Private Sub cmdChangeInd_Click()
    Dim i As Integer, iNumInd As Long, sInd As String
    
    If Not bSelectingInd Then
        Call DirtyFlagRoutine 'check if changes need/are wanted to be saved
        Frame1.Caption = sFrameCaption$ & CurrentIndicator$
        'load the indicator types
        iNumInd = GetNumIniKeys(sINIsetFile$, sIndicators$)
        For i = picColor.LBound + 1 To picColor.UBound
            Unload picColor(i)  'don't need the picBxs
        Next
        picColor(0).Visible = False 'can't unload the first one so hide it
        lstVwSettings.ListItems.Clear
        lstVwSettings.ToolTipText = "Dbl Click to Change"
        lstVwSettings.ColumnHeaders(1).Text = sIndicators$
        lstVwSettings.ColumnHeaders.Remove 2 'don't need the header
        If iNumInd <> 0 Then  'get the indicator strings
            For i = 1 To iNumInd
                sInd$ = GetIni(sINIsetFile$, sIndicators$, CStr(i))
                lstVwSettings.ListItems.Add i, sInd$, sInd$
            Next
        End If
        sPrevInd$ = CurrentIndicator$  'save for cancel
        cmdChangeInd.Caption = "OK  |  Esc"  'dual purpose cmdButton
        cmdUndo.Enabled = False
        bSelectingInd = True
    Else
        If fButtonRightSide Then  'Esc.. reset indicator change
            CurrentIndicator$ = sPrevInd$
        Else  'do the change
            Call WriteIni(sINIsetFile$, "Settings", "CurrentIndicator", CurrentIndicator$)
            Call UpdateChart
        End If
        Call LoadCurrentSettings(sIndicators$)
        cmdChangeInd.Caption = "C&hange"
        bSelectingInd = False
    End If
End Sub
Private Sub DirtyFlagRoutine()
    Dim iResult As Long
     'if last indicator was edited then reset if wanted
    If bDirtyFlag Then iResult = MsgBox("Editting for the last Indicator has not been saved..." & _
            vbCrLf & "Do you want to save the changes?", vbQuestion + vbYesNo + vbDefaultButton2, _
            "Changes Not Saved")
    If iResult = vbNo Then Call ResetValues
    bDirtyFlag = False
End Sub
Private Sub cmdChangeInd_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not bSelectingInd Then
        cmdChangeInd.ToolTipText = sEmpty
    Else
        If x < cmdChangeInd.Width / 2 Then  'OK side
            fButtonRightSide = False
            cmdChangeInd.ToolTipText = "OK to Save"
        Else    'Esc side
            fButtonRightSide = True
            cmdChangeInd.ToolTipText = "Cancel Change"
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    Unload Me
End Sub

Private Sub cmdUndo_Click()
    oLastEditItem.SubItems(1) = oLastEditItem.Tag
    Call WriteIni(sINIsetFile$, sSection$, oLastEditItem.Text, oLastEditItem.Tag)
    Call UpdateChart
    cmdUndo.Enabled = False
End Sub

Private Sub Form_Load()
    sOrgInd$ = CurrentIndicator$
    sFrameCaption$ = Frame1.Caption & " -> "
    Call LoadCurrentSettings(sIndicators$)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveIndicatorSettings
    Set frmIndicators = Nothing
End Sub

Private Sub LoadCurrentSettings(sType As String)
    Dim i As Integer, sKey As String, iNumSettings As Long
    Dim sValue As String, iPicCount As Long, iRowCount As Long
    
    lstVwSettings.ListItems.Clear
    lstVwSettings.ToolTipText = "Dbl Click to Edit"
    On Error Resume Next  'if the header is already added proceed anyway
    lstVwSettings.ColumnHeaders.Add (2), "Value", "Value", 1700, vbRightJustify
    On Error GoTo 0

    Select Case sType
        Case sIndicators$
            Frame1.Caption = sFrameCaption$ & CurrentIndicator$
            lstVwSettings.ColumnHeaders(1).Text = sParameters$
            sSection$ = CurrentIndicator$ & "Settings"
            iNumSettings = GetNumIniKeys(sINIsetFile$, sSection$)
        Case sAverages$
            Frame1.Caption = sFrameCaption$ & sAverages$
            lstVwSettings.ColumnHeaders(1).Text = sParameters$
            sSection$ = "AvgSettings"
            iNumSettings = GetNumIniKeys(sINIsetFile$, sSection$)
    End Select
    
    If iNumSettings <> 0 Then
        For i = 1 To iNumSettings
            iRowCount = iRowCount + 1
            sKey$ = GetIniKey(sINIsetFile$, sSection$, CStr(i))
            sValue$ = GetIni(sINIsetFile$, sSection$, sKey$)
            lstVwSettings.ListItems.Add iRowCount, sKey$, sKey$
            lstVwSettings.ListItems(sKey$).SubItems(1) = sValue$
            'save the original value(for cancelling)
            lstVwSettings.ListItems(sKey$).Tag = sValue$
            If InStr(UCase(sKey$), "COLOR") <> 0 Then 'color setting, add picBx
                If iPicCount <> 0 Then
                    Load picColor(iPicCount)
                End If
                picColor(iPicCount).Move lstVwSettings.Left + 20 + Frame1.Left + lstVwSettings.ListItems(iRowCount).Left + lstVwSettings.ColumnHeaders(1).Width, _
                    lstVwSettings.Top + 60 + Frame1.Top + lstVwSettings.ListItems(iRowCount).Top, _
                    lstVwSettings.ColumnHeaders(2).Width, lstVwSettings.ListItems(iRowCount).Height
                picColor(iPicCount).ZOrder 0
                picColor(iPicCount).BackColor = CLng(sValue$)
                'save key in tag so we know item when pic is dblclked
                picColor(iPicCount).Tag = sKey$
                picColor(iPicCount).Visible = True
                'modify the item tag so we know which picBx index goes with the item
                lstVwSettings.ListItems(sKey$).Tag = lstVwSettings.ListItems(sKey$).Tag _
                                                    & "|" & iPicCount
                iPicCount = iPicCount + 1
            End If
            
            If sType = sAverages$ And i Mod 3 = 0 Then
                iRowCount = iRowCount + 1  'adding a blank line between averages
                lstVwSettings.ListItems.Add iRowCount, "blank" & iRowCount, sEmpty
                lstVwSettings.ListItems("blank" & iRowCount).ToolTipText = sEmpty
            End If
        Next
    End If
End Sub
Private Sub lstVwSettings_DblClick()
    If Not bSelectingInd Then
        'if we dblClked then do the edit on current item
        If InStr(UCase(oClickedItem), "COLOR") <> 0 Then 'call color selector
            Dim p As Long
            oClickedItem.SubItems(1) = GetColorDlg(oClickedItem.SubItems(1))
            'fish out the picBx index from the item tag
            p = InStr(oClickedItem.Tag, "|")
            picColor(Mid$(oClickedItem.Tag, p + 1)).BackColor = oClickedItem.SubItems(1)
            Call UpdateColor(oClickedItem)
        ElseIf oClickedItem <> sEmpty Then
            bBooleanEdit = False
            txtEdit.Move lstVwSettings.Left + 4 + Frame1.Left + oClickedItem.Left + lstVwSettings.ColumnHeaders(1).Width, _
                        lstVwSettings.Top + 10 + Frame1.Top + oClickedItem.Top, _
                        lstVwSettings.ColumnHeaders(2).Width, oClickedItem.Height - 10
            
            If UCase(oClickedItem.SubItems(1)) = "TRUE" Or UCase(oClickedItem.SubItems(1)) = "FALSE" Then
                bBooleanEdit = True
            End If
            txtEdit.Text = oClickedItem.SubItems(1)
            txtEdit.Visible = True
            txtEdit.SetFocus
        End If
    Else
        CurrentIndicator$ = oClickedItem.Text
        Frame1.Caption = sFrameCaption$ & CurrentIndicator$
        Call WriteIni(sINIsetFile$, "Settings", "CurrentIndicator", CurrentIndicator$)
        Call UpdateChart
    End If
    
End Sub

Private Sub lstVwSettings_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'save item here since we don't have an "ItemDblClick" event
    Set oClickedItem = Item
End Sub

Private Sub lstVwSettings_KeyPress(KeyAscii As Integer)
'Debug.Print KeyAscii
    Select Case KeyAscii
        Case 13 'enter
            If bSelectingInd Then  'make change happen
                CurrentIndicator$ = lstVwSettings.SelectedItem.Text
                fButtonRightSide = False 'OK button side
                Call cmdChangeInd_Click
            End If
    End Select
End Sub

Private Sub lstVwSettings_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        txtEdit.Visible = False
    ElseIf Button = 2 Then
        
    End If
End Sub

Private Sub lstVwSettings_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'needed... listvw was taking the focus away from the editbox
    If Not bSelectingInd And txtEdit.Visible = True Then txtEdit.SetFocus
End Sub

Private Sub UpdateColor(ByVal Item As MSComctlLib.ListItem)
    Set oLastEditItem = Item
    bDirtyFlag = True 'set edited flag
    cmdUndo.Enabled = True
    'save the new value and redraw chart for dynamic real-time updating
    Call WriteIni(sINIsetFile$, sSection$, Item.Text, Item.SubItems(1))
    Call UpdateChart
End Sub
Private Sub picColor_DblClick(Index As Integer)
    picColor(Index).BackColor = GetColorDlg(picColor(Index).BackColor)
    'key of the item was stored in the pic tag
    lstVwSettings.ListItems(picColor(Index).Tag).SubItems(1) = picColor(Index).BackColor
    Call UpdateColor(lstVwSettings.ListItems(picColor(Index).Tag))
End Sub

Private Sub txtEdit_Change()
    If bBooleanEdit Then
        'filter the edit text for boolean values only
        Select Case UCase(Left$(txtEdit.Text, 1))
            Case "T", "1", "-1"
                txtEdit.Text = "True"
            Case "F", "0"
                txtEdit.Text = "False"
            Case Else
                txtEdit.Text = sEmpty
        End Select
    Else
        'just want numeric values
        If Not IsNumeric(txtEdit.Text) Then txtEdit.Text = sEmpty
    End If
End Sub

Private Sub txtEdit_GotFocus()
    txtEdit.SelStart = 0
    txtEdit.SelLength = Len(txtEdit.Text)
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
'Debug.Print KeyCode
    Select Case KeyCode
        Case vbKeyEscape
            txtEdit.Text = sEmpty
            txtEdit.Visible = False
        Case vbKeyReturn
            oClickedItem.SubItems(1) = txtEdit.Text
            bDirtyFlag = True 'set edited flag
            cmdUndo.Enabled = True
            Set oLastEditItem = oClickedItem
            'save the new value and redraw chart for dynamic real-time updating
            Call WriteIni(sINIsetFile$, sSection$, oClickedItem.Text, oClickedItem.SubItems(1))
            Call UpdateChart
            txtEdit.Visible = False
        Case Else
            'do nothing
    End Select
End Sub

Private Sub tmrAfterLoad_Timer()
    tmrAfterLoad.Enabled = False
    Call PositionMousePointer(cmdCancel.hWnd, cmdCancel.Width \ 2, cmdCancel.Height / 1.2, False)
End Sub
