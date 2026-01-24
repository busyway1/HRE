VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendar 
   Caption         =   "날짜를 선택하세요"
   ClientHeight    =   3240
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   2820
   OleObjectBlob   =   "frmCalendar.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum frmLocation
    xlCenter = 0
    xlNextToCell = 1
    xlNextToCursor = 2
End Enum

Dim returnDate As Date
Dim vLists As Variant
Dim YearBetween As Long
Private formClosed As Boolean


Private Sub cboMonth_Click()
Me.lblMonth.Caption = Me.cboMonth.Value
Me.scrlMonth.Value = Left(Me.lblMonth.Caption, Len(Me.lblMonth.Caption) - 1)
End Sub

Private Sub cboYear_Click()
Me.lblYear.Caption = Me.cboYear.Value
resetDate
End Sub

Private Sub bgNow_Click()
Me.lblMonth.Caption = month(Date) & "월"
Me.lblYear.Caption = Year(Date) & "년"
Me.scrlMonth.Value = Left(Me.lblMonth.Caption, Len(Me.lblMonth.Caption) - 1)
End Sub

Private Sub lblNow_Click()
Me.lblMonth.Caption = month(Date) & "월"
Me.lblYear.Caption = Year(Date) & "년"
Me.scrlMonth.Value = Left(Me.lblMonth.Caption, Len(Me.lblMonth.Caption) - 1)
End Sub

Private Sub lblMonth_Click()
Me.cboMonth.DropDown
End Sub

Private Sub lblYear_Click()
Me.cboYear.DropDown
End Sub

Private Sub scrlMonth_Change()

If scrlMonth.Value > 0 And scrlMonth.Value < 13 Then
Me.lblMonth.Caption = Me.scrlMonth.Value & "월"
ElseIf scrlMonth.Value <= 0 Then
    scrlMonth.Value = 12: Me.lblMonth.Caption = Me.scrlMonth.Value & "월": Me.lblYear.Caption = Left(Me.lblYear.Caption, 4) - 1 & "년"
Else
    scrlMonth.Value = 1: Me.lblMonth.Caption = Me.scrlMonth.Value & "월": Me.lblYear.Caption = Left(Me.lblYear.Caption, 4) + 1 & "년"
End If

resetDate

End Sub
Private Sub UserForm_Initialize()
Dim i As Long
For i = 1 To 42
    With Me.Controls("Label" & i)
        .BackStyle = 0
    End With
Next

For i = 43 To 84
    With Me.Controls("Label" & i)
        .Caption = ""
        .Top = .Top - 2
        .Left = .Left - 2
        .Width = .Width + 3
        .Height = .Height + 2
        .BackStyle = 1
        .Font.Bold = True
    End With
Next

With Me.cboMonth
    For i = 1 To 12: .AddItem i & "월": Next
End With

ReDim vLists(0 To 41)
For i = 0 To 41
    vLists(i) = "Label" & i + 43
Next

resetYear

Me.lblYear.Caption = Year(Date) & "년"
Me.lblMonth.Caption = month(Date) & "월"
Me.scrlMonth.Value = month(Date)

End Sub

Function GetDate(Optional Location As frmLocation = 2, Optional YearGap As Long = 3) As Variant
    Dim Top As Double: Dim Left As Double
    Dim MousePos As POINTAPI

    If Location = 0 Then
        Me.StartUpPosition = 1
    ElseIf Location = 1 Then
        Me.StartUpPosition = 0
        Me.Top = ActiveCell.Top + ActiveCell.Height + Me.Height
        Me.Left = ActiveCell.Offset(0, 1).Left
    Else
        MousePos = convertMouseToForm()
        Me.StartUpPosition = 0
        Me.Top = MousePos.Y
        Me.Left = MousePos.X
    End If

    YearBetween = YearGap
    resetYear
    
    formClosed = False
    Me.Show
    
    If formClosed Then
        GetDate = "-"
    Else
        GetDate = returnDate
    End If
   
    Unload Me
End Function

Sub resetYear()

Dim i As Long

With Me.cboYear
    For i = -YearBetween To YearBetween: .AddItem Year(Date) + i & "년": Next
End With

End Sub
Sub lblClick(lbl As MSForms.Label)
Dim Y As Integer: Dim M As Integer: Dim D As Integer
Y = Left(Me.lblYear.Caption, 4): M = Left(Me.lblMonth.Caption, Len(Me.lblMonth.Caption) - 1): D = lbl.Caption

returnDate = DateSerial(Y, M, D)
Unload Me
End Sub


Sub resetDate()

Dim Y As Integer: Dim M As Integer: Dim D As Integer: Dim w As Integer
Dim i As Integer
Y = Left(Me.lblYear.Caption, 4): M = Left(Me.lblMonth.Caption, Len(Me.lblMonth.Caption) - 1): D = day(DateSerial(Y, M + 1, 1) - 1)
w = Weekday(DateSerial(Y, M, 1))

For i = 1 To 42
    Me.Controls("Label" & i).Enabled = True: Me.Controls("Label" & i + 41).Enabled = True
    Me.Controls("Label" & i).Caption = day(DateSerial(Y, M, i) - w + 1)
    
    If month(DateSerial(Y, M, i) - w + 1) <> M Then
        Me.Controls("Label" & i).ForeColor = RGB(222, 222, 222): Me.Controls("Label" & i).Enabled = False: Me.Controls("Label" & i + 41).Enabled = False
    ElseIf Weekday(DateSerial(Y, M, i) - w + 1) = 1 Then
        Me.Controls("Label" & i).ForeColor = RGB(255, 0, 0):
    ElseIf Weekday(DateSerial(Y, M, i) - w + 1) = 7 Then
        Me.Controls("Label" & i).ForeColor = RGB(0, 0, 255)
    Else
        Me.Controls("Label" & i).ForeColor = RGB(0, 0, 0)
    End If
    
    If DateSerial(Y, M, Me.Controls("Label" & i).Caption) = Date And month(DateSerial(Y, M, i) - w + 1) = M Then
        Me.Controls("Label" & i + 42).BackColor = RGB(51, 51, 51)
        Me.Controls("Label" & i).ForeColor = RGB(255, 255, 255)
    Else
        Me.Controls("Label" & i + 42).BackColor = RGB(255, 255, 255)
    End If
Next

End Sub

Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label43: End Sub
Private Sub Label2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label44: End Sub
Private Sub Label3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label45: End Sub
Private Sub Label4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label46: End Sub
Private Sub Label5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label47: End Sub
Private Sub Label6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label48: End Sub
Private Sub Label7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label49: End Sub
Private Sub Label8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label50: End Sub
Private Sub Label9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label51: End Sub
Private Sub Label10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label52: End Sub
Private Sub Label11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label53: End Sub
Private Sub Label12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label54: End Sub
Private Sub Label13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label55: End Sub
Private Sub Label14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label56: End Sub
Private Sub Label15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label57: End Sub
Private Sub Label16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label58: End Sub
Private Sub Label17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label59: End Sub
Private Sub Label18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label60: End Sub
Private Sub Label19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label61: End Sub
Private Sub Label20_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label62: End Sub
Private Sub Label21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label63: End Sub
Private Sub Label22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label64: End Sub
Private Sub Label23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label65: End Sub
Private Sub Label24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label66: End Sub
Private Sub Label25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label67: End Sub
Private Sub Label26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label68: End Sub
Private Sub Label27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label69: End Sub
Private Sub Label28_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label70: End Sub
Private Sub Label29_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label71: End Sub
Private Sub Label30_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label72: End Sub
Private Sub Label31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label73: End Sub
Private Sub Label32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label74: End Sub
Private Sub Label33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label75: End Sub
Private Sub Label34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label76: End Sub
Private Sub Label35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label77: End Sub
Private Sub Label36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label78: End Sub
Private Sub Label37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label79: End Sub
Private Sub Label38_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label80: End Sub
Private Sub Label39_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label81: End Sub
Private Sub Label40_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label82: End Sub
Private Sub Label41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label83: End Sub
Private Sub Label42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): OnHover_Css Me.Label84: End Sub

Private Sub Label1_Click(): Call lblClick(Me.Label1): End Sub
Private Sub Label2_Click(): Call lblClick(Me.Label2): End Sub
Private Sub Label3_Click(): Call lblClick(Me.Label3): End Sub
Private Sub Label4_Click(): Call lblClick(Me.Label4): End Sub
Private Sub Label5_Click(): Call lblClick(Me.Label5): End Sub
Private Sub Label6_Click(): Call lblClick(Me.Label6): End Sub
Private Sub Label7_Click(): Call lblClick(Me.Label7): End Sub
Private Sub Label8_Click(): Call lblClick(Me.Label8): End Sub
Private Sub Label9_Click(): Call lblClick(Me.Label9): End Sub
Private Sub Label10_Click(): Call lblClick(Me.Label10): End Sub
Private Sub Label11_Click(): Call lblClick(Me.Label11): End Sub
Private Sub Label12_Click(): Call lblClick(Me.Label12): End Sub
Private Sub Label13_Click(): Call lblClick(Me.Label13): End Sub
Private Sub Label14_Click(): Call lblClick(Me.Label14): End Sub
Private Sub Label15_Click(): Call lblClick(Me.Label15): End Sub
Private Sub Label16_Click(): Call lblClick(Me.Label16): End Sub
Private Sub Label17_Click(): Call lblClick(Me.Label17): End Sub
Private Sub Label18_Click(): Call lblClick(Me.Label18): End Sub
Private Sub Label19_Click(): Call lblClick(Me.Label19): End Sub
Private Sub Label20_Click(): Call lblClick(Me.Label20): End Sub
Private Sub Label21_Click(): Call lblClick(Me.Label21): End Sub
Private Sub Label22_Click(): Call lblClick(Me.Label22): End Sub
Private Sub Label23_Click(): Call lblClick(Me.Label23): End Sub
Private Sub Label24_Click(): Call lblClick(Me.Label24): End Sub
Private Sub Label25_Click(): Call lblClick(Me.Label25): End Sub
Private Sub Label26_Click(): Call lblClick(Me.Label26): End Sub
Private Sub Label27_Click(): Call lblClick(Me.Label27): End Sub
Private Sub Label28_Click(): Call lblClick(Me.Label28): End Sub
Private Sub Label29_Click(): Call lblClick(Me.Label29): End Sub
Private Sub Label30_Click(): Call lblClick(Me.Label30): End Sub
Private Sub Label31_Click(): Call lblClick(Me.Label31): End Sub
Private Sub Label32_Click(): Call lblClick(Me.Label32): End Sub
Private Sub Label33_Click(): Call lblClick(Me.Label33): End Sub
Private Sub Label34_Click(): Call lblClick(Me.Label34): End Sub
Private Sub Label35_Click(): Call lblClick(Me.Label35): End Sub
Private Sub Label36_Click(): Call lblClick(Me.Label36): End Sub
Private Sub Label37_Click(): Call lblClick(Me.Label37): End Sub
Private Sub Label38_Click(): Call lblClick(Me.Label38): End Sub
Private Sub Label39_Click(): Call lblClick(Me.Label39): End Sub
Private Sub Label40_Click(): Call lblClick(Me.Label40): End Sub
Private Sub Label41_Click(): Call lblClick(Me.Label41): End Sub
Private Sub Label42_Click(): Call lblClick(Me.Label42): End Sub

Private Sub scrlMonth_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        returnDate = Date
        Unload Me
    ElseIf KeyCode = vbKeyEscape Then
        formClosed = True
        Unload Me
    End If
End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        returnDate = Date
        Unload Me
    ElseIf KeyAscii = 27 Then ' ESC key
        formClosed = True
        Unload Me
    End If
End Sub

Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        returnDate = Date
        Unload Me
    ElseIf KeyCode = vbKeyEscape Then
        formClosed = True
        Unload Me
    End If
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim ctl As control: Dim i As Long
Dim vList As Variant

For Each ctl In Me.Controls
    If ctl.BackColor = RGB(182, 182, 182) Then ctl.BackColor = RGB(255, 255, 255): Exit Sub
Next

End Sub

Private Sub OnHover_Css(lbl As control)
    With lbl
        If .BackColor <> RGB(51, 51, 51) Then .BackColor = RGB(182, 182, 182)
    End With
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        formClosed = True
        Cancel = False
    End If
End Sub

