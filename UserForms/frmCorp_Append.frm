VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCorp_Append 
   Caption         =   "대상법인 추가"
   ClientHeight    =   6705
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9576.001
   OleObjectBlob   =   "frmCorp_Append.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmCorp_Append"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private checkRatio, checkDate_Acq, checkDate_Dis As String
Private Sub Append_Cmd_Click()
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim existingRow As ListRow
    Dim isDuplicate As Boolean
    Dim lastRow As Long
    
    On Error Resume Next
    
    Set tbl = CorpMaster.ListObjects("Corp")
    
    If Me.CorpCode_Txt.Value = "" Or Me.CorpName_Txt.Value = "" Or Me.Hierarchy_Combo.Value = "" _
        Or Me.Ratio_Txt.Value = "" Or Me.AcqDate_Cmd.Caption = "" Or Me.DisDate_Cmd.Caption = "" _
        Or Me.Prior_Combo.Value = "" Or Me.cmbPeople.Value = "" Then
        Msg "모든 값이 채워져 있는 지 확인하세요!", vbCritical
        Exit Sub
    End If
    
    If IsNumeric(Me.CorpName_Txt.Value) Then
        Msg "법인명이 제대로 기입되었는지 확인하세요!", vbCritical
        Exit Sub
    End If
    
    isDuplicate = False
    For Each existingRow In tbl.ListRows
        If Trim(CStr(existingRow.Range(2).Value)) = Trim(CStr(Me.CorpCode_Txt.Value)) Or _
           Trim(CStr(existingRow.Range(3).Value)) = Trim(CStr(Me.CorpName_Txt.Value)) Then
            isDuplicate = True
            Exit For
        End If
    Next existingRow
    
    If isDuplicate Then
        Msg "이미 존재하는 데이터입니다. 회사코드와 회사이름를 Corp에서 확인해주세요.", vbExclamation
        Exit Sub
    End If
    
    CorpMaster.Unprotect PASSWORD
    lastRow = tbl.ListRows.count
    
    Set newRow = tbl.ListRows.Add
    tbl.DataBodyRange(lastRow + 1, 9).FormulaR1C1 = tbl.DataBodyRange(lastRow, 9).FormulaR1C1
    tbl.DataBodyRange(lastRow + 1, 10).FormulaR1C1 = tbl.DataBodyRange(lastRow, 10).FormulaR1C1
    tbl.DataBodyRange(lastRow + 1, 11).FormulaR1C1 = tbl.DataBodyRange(lastRow, 11).FormulaR1C1
    tbl.DataBodyRange(lastRow + 1, 13).FormulaR1C1 = tbl.DataBodyRange(lastRow, 13).FormulaR1C1
    tbl.DataBodyRange(lastRow + 1, 14).FormulaR1C1 = tbl.DataBodyRange(lastRow, 14).FormulaR1C1
    tbl.DataBodyRange(lastRow + 1, 15).FormulaR1C1 = tbl.DataBodyRange(lastRow, 15).FormulaR1C1
    
    newRow.Range(1) = Me.No_Txt.Value
    newRow.Range(2) = Me.CorpCode_Txt.Value
    newRow.Range(3) = Me.CorpName_Txt.Value
    newRow.Range(4) = Me.Hierarchy_Combo.Value
    newRow.Range(5) = Me.Ratio_Txt.Value
    newRow.Range(6) = Me.AcqDate_Cmd.Caption
    newRow.Range(6).NumberFormat = "yyyy-mm-dd"
    newRow.Range(7) = Me.DisDate_Cmd.Caption
    newRow.Range(7).NumberFormat = "yyyy-mm-dd"
    newRow.Range(8) = Me.Prior_Combo.Value
    newRow.Range(12) = Me.cmbPeople.Value
    
    CorpMaster.Protect PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    CorpMaster.Activate: CorpMaster.Range("A1").Select
    
    With Check.Cells(15, 4)
        .Value = "If Any"
        .Interior.Color = RGB(237, 237, 237)
        .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
        .Offset(0, 2).Value = GetUserInfo()
    End With
    
    '로그 전송
    LogData CorpMaster.name, "<대상법인 추가>" & vbNewLine & vbNewLine & _
                     "[추가 전]" & vbNewLine & _
                     "No: " & vbNewLine & _
                     "법인코드: " & vbNewLine & _
                     "법인명: " & vbNewLine & _
                     "Hierarchy: " & vbNewLine & _
                     "유효지분율: " & vbNewLine & _
                     "설립(취득)일: " & vbNewLine & _
                     "매각(청산)일: " & vbNewLine & _
                     "직전사업연도 외감대상: " & vbNewLine & _
                     "담당자명: " & vbNewLine & vbNewLine & _
                     "[추가 후]" & vbNewLine & _
                     "No: " & Me.No_Txt.Value & vbNewLine & _
                     "법인코드: " & Me.CorpCode_Txt.Value & vbNewLine & _
                     "법인명: " & Me.CorpName_Txt.Value & vbNewLine & _
                     "Hierarchy: " & Me.Hierarchy_Combo.Value & vbNewLine & _
                     "유효지분율: " & Me.Ratio_Txt.Value & vbNewLine & _
                     "설립(취득)일: " & Me.AcqDate_Cmd.Caption & vbNewLine & _
                     "매각(청산)일: " & Me.DisDate_Cmd.Caption & vbNewLine & _
                     "직전사업연도 외감대상: " & Me.Prior_Combo.Value & vbNewLine & _
                     "담당자명: " & Me.cmbPeople.Value

    Msg "새로운 법인이 성공적으로 추가되었습니다.", vbInformation
    Unload Me
    Set tbl = Nothing: newRow = Nothing: Set existingRow = Nothing
End Sub
Private Sub UserForm_Initialize()
    Dim tbl As ListObject
    Dim tblPeople As ListObject
    Dim lastRow As Long
    
    On Error Resume Next
    Me.Caption = AppName & " " & AppType
    If Check.Cells(12, 4).Value <> "Complete" Or Check.Cells(13, 4).Value <> "Complete" Or Check.Cells(14, 4).Value <> "Complete" Then
       GoEnd "이전 단계를 완료해주세요!"
    End If
    
    With Check.Cells(15, 4)
       .Value = "If Any"
       .Interior.Color = RGB(237, 237, 237)
       .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
       .Offset(0, 2).Value = GetUserInfo()
    End With
    
    Set tbl = CorpMaster.ListObjects("Corp")
    Set tblPeople = HideSheet.ListObjects("People_Work")
    lastRow = tbl.DataBodyRange.Rows.count
    
    Me.No_Txt.Value = lastRow + 1
    UpdateButtonState
    Me.Hierarchy_Combo.List = Array("본사", "종속회사", "관계회사", "손자회사")
    Me.Prior_Combo.List = Array("O", "X")
    checkRatio = True: checkDate_Acq = True: checkDate_Dis = True
    
    
    cmbPeople.RowSource = tblPeople.DataBodyRange.Address(External:=True)
    
    
    Set tbl = Nothing: Set tblPeople = Nothing
End Sub
Private Sub UpdateButtonState()
    DisDate_Cmd.Enabled = Not Dis_Chk.Value
    If DisDate_Cmd.Enabled Then
        DisDate_Cmd.BackColor = &H80000005
    Else
        DisDate_Cmd.BackColor = &H8000000C
    End If
End Sub
Private Sub Dis_Chk_Click()
    UpdateButtonState
    If Me.Dis_Chk.Value = True Then
        Me.DisDate_Cmd.Caption = "-"
    End If
End Sub
Private Sub DisDate_Cmd_Click()
    If Not Me.Dis_Chk.Value Then
        Me.DisDate_Cmd.Caption = frmCalendar.GetDate(xlNextToCursor)
        Dim dateValue As Variant
        dateValue = Me.DisDate_Cmd.Caption
        
        If dateValue = "-" Then
            checkDate_Dis = True
            Me.Dis_Chk.Value = True
        Else
            If IsDate(dateValue) Then
                If CDate(dateValue) > Date Then
                    Msg "오늘 날짜 이후의 날짜는 선택할 수 없습니다.", vbExclamation
                    checkDate_Dis = False
                Else
                    checkDate_Dis = True
                End If
            Else
                Msg "유효하지 않은 날짜 형식입니다.", vbExclamation
                checkDate_Dis = False
            End If
        End If
    End If
    
    If checkRatio And checkDate_Acq And checkDate_Dis Then
        Me.Append_Cmd.Enabled = True
    Else
        Me.Append_Cmd.Enabled = False
    End If
End Sub

Private Sub AcqDate_Cmd_Click()
    Me.AcqDate_Cmd.Caption = frmCalendar.GetDate(xlNextToCursor)
    Dim dateValue As Variant
    dateValue = Me.AcqDate_Cmd.Caption
    
    If dateValue = "-" Then
        Msg "날짜가 선택되지 않았습니다. 날짜를 선택해주세요.", vbExclamation
        checkDate_Acq = False
        Exit Sub
    End If
    
    If IsDate(dateValue) Then
        If CDate(dateValue) > Date Then
            Msg "오늘 날짜 이후의 날짜는 선택할 수 없습니다.", vbExclamation
            checkDate_Acq = False
        Else
            checkDate_Acq = True
        End If
    Else
        Msg "유효하지 않은 날짜 형식입니다.", vbExclamation
        checkDate_Acq = False
    End If
    
    If checkRatio And checkDate_Acq And checkDate_Dis Then
        Me.Append_Cmd.Enabled = True
    Else
        Me.Append_Cmd.Enabled = False
    End If
End Sub
Private Sub Ratio_Txt_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim inputValue As Double
    If IsNumeric(Me.Ratio_Txt.Value) And Me.Ratio_Txt.Value >= 0 And Me.Ratio_Txt.Value <= 100 Then
        inputValue = CDbl(Me.Ratio_Txt.Value)
        Me.Ratio_Txt.Value = Format(inputValue / 100, "0.00%")
        checkRatio = True
    Else
        checkRatio = False
        Msg "지분율 형식을  확인해주세요!", vbExclamation
    End If
    
    If checkRatio And checkDate_Acq And checkDate_Dis Then
        Me.Append_Cmd.Enabled = True
    Else
        Me.Append_Cmd.Enabled = False
    End If
End Sub
Private Sub Cancel_Cmd_Click()
    Unload Me
End Sub
Private Sub cmbPeople_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.cmbPeople
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    UnhookListBoxScroll
End Sub
