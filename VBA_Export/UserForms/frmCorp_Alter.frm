VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCorp_Alter 
   Caption         =   "대상법인 정보 수정"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9576.001
   OleObjectBlob   =   "frmCorp_Alter.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmCorp_Alter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private isNormalClose As Boolean
Private checkRatio As Boolean
Private checkDate_Acq As Boolean
Private checkDate_Dis As Boolean
Private logVal1, logVal2, logVal3, logVal4, logVal5, logVal6, logVal7, logVal8 As String
Public mRowIndex As Long
Public Property Let rowIndex(Value As Long)
    mRowIndex = Value
End Property
Public Property Get rowIndex() As Long
    rowIndex = mRowIndex
End Property
Public Sub LoadData(cellarray() As Variant, rowIndex As Long)
    Me.No_Txt.Value = cellarray(1).Value
    Me.CorpCode_Txt.Value = cellarray(2).Value
    Me.CorpName_Txt.Value = cellarray(3).Value
    Me.Hierarchy_Combo.Value = cellarray(4).Value
    Me.Ratio_Txt.Value = Format(cellarray(5).Value, "0.00%")
    Me.AcqDate_Cmd.Caption = cellarray(6).Value
    Me.DisDate_Cmd.Caption = cellarray(7).Value
    Me.cmbPeople.Value = cellarray(9).Value
    
    If cellarray(7).Value = "-" Then
        Me.Dis_Chk.Value = True
    End If
    
    Me.Prior_Combo.Value = cellarray(8).Value
    Me.rowIndex = rowIndex
    
    logVal1 = Me.CorpCode_Txt.Value: logVal2 = Me.CorpName_Txt.Value: logVal3 = Me.Hierarchy_Combo.Value: logVal4 = Me.Ratio_Txt.Value
    logVal5 = Me.AcqDate_Cmd.Caption: logVal6 = Me.DisDate_Cmd.Caption: logVal7 = Me.Dis_Chk.Value: logVal8 = Me.cmbPeople.Value

End Sub
Private Sub Alter_Cmd_Click()
    Dim tbl As ListObject
    Dim editRow As ListRow
    
    On Error Resume Next
    
    Set tbl = CorpMaster.ListObjects("Corp")
    Set editRow = tbl.ListRows(Me.rowIndex)

    CorpMaster.Unprotect PASSWORD

    editRow.Range(1) = Me.No_Txt.Value
    editRow.Range(2) = Me.CorpCode_Txt.Value
    editRow.Range(3) = Me.CorpName_Txt.Value
    editRow.Range(4) = Me.Hierarchy_Combo.Value
    editRow.Range(5) = Me.Ratio_Txt.Value
    editRow.Range(6) = Format(Me.AcqDate_Cmd.Caption, "yyyy-mm-dd")
    editRow.Range(6).NumberFormat = "yyyy-mm-dd"
    editRow.Range(7) = Format(Me.DisDate_Cmd.Caption, "yyyy-mm-dd")
    editRow.Range(7).NumberFormat = "yyyy-mm-dd"
    editRow.Range(8) = Me.Prior_Combo.Value
    editRow.Range(12) = Me.cmbPeople.Value
    
    CorpMaster.Protect PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    
    With Check.Cells(15, 4)
            .Value = "If Any"
            .Interior.Color = RGB(237, 237, 237)
            .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
            .Offset(0, 2).Value = GetUserInfo()
    End With
    
    '로그 전송
    LogData CorpMaster.name, "<대상법인 수정>" & vbNewLine & vbNewLine & _
                     "[변경 전]" & vbNewLine & _
                     "No: " & Me.No_Txt.Value & vbNewLine & _
                     "법인코드: " & logVal1 & vbNewLine & _
                     "법인명: " & logVal2 & vbNewLine & _
                     "Hierarchy: " & logVal3 & vbNewLine & _
                     "유효지분율: " & logVal4 & vbNewLine & _
                     "취득(설립)일: " & logVal5 & vbNewLine & _
                     "매각(청산)일: " & logVal6 & vbNewLine & _
                     "직전사업연도 외감대상: " & logVal7 & vbNewLine & _
                     "담당자명: " & logVal8 & vbNewLine & vbNewLine & _
                     "[변경 후]" & vbNewLine & _
                     "No: " & Me.No_Txt.Value & vbNewLine & _
                     "법인코드: " & Me.CorpCode_Txt.Value & vbNewLine & _
                     "법인명: " & Me.CorpName_Txt.Value & vbNewLine & _
                     "Hierarchy: " & Me.Hierarchy_Combo.Value & vbNewLine & _
                     "유효지분율: " & Me.Ratio_Txt.Value & vbNewLine & _
                     "취득(설립)일: " & Me.AcqDate_Cmd.Caption & vbNewLine & _
                     "매각(청산)일: " & Me.DisDate_Cmd.Caption & vbNewLine & _
                     "직전사업연도 외감대상: " & Me.Prior_Combo.Value & vbNewLine & _
                     "담당자명: " & Me.cmbPeople.Value

    Msg "대상법인 데이터가 수정되었습니다.", vbInformation
    isNormalClose = True
    Unload Me
    Set tbl = Nothing: Set editRow = Nothing
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
        Me.Alter_Cmd.Enabled = True
    Else
        Me.Alter_Cmd.Enabled = False
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
        Me.Alter_Cmd.Enabled = True
    Else
        Me.Alter_Cmd.Enabled = False
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
        Me.Alter_Cmd.Enabled = True
    Else
        Me.Alter_Cmd.Enabled = False
    End If
End Sub
Private Sub Cancel_Cmd_Click()
    With Check.Cells(15, 4)
        .Value = "If Any"
        .Interior.Color = RGB(237, 237, 237)
        .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
        .Offset(0, 2).Value = GetUserInfo()
    End With
    End
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    Dim tblPeople As ListObject
    Set tblPeople = HideSheet.ListObjects("People_Work")
    UpdateButtonState
    Me.Caption = AppName & " " & AppType
    Me.Hierarchy_Combo.List = Array("본사", "종속회사", "관계회사")
    Me.Prior_Combo.List = Array("O", "X")
    checkRatio = True: checkDate_Acq = True: checkDate_Dis = True
    cmbPeople.RowSource = tblPeople.DataBodyRange.Address(External:=True)
    Set tblPeople = Nothing
End Sub
Private Sub UpdateButtonState()
    DisDate_Cmd.Enabled = Not Dis_Chk.Value
    If DisDate_Cmd.Enabled Then
        DisDate_Cmd.BackColor = &H80000005
    Else
        DisDate_Cmd.BackColor = &H8000000C
    End If
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
     UnhookListBoxScroll
     If Not isNormalClose Then
        With Check.Cells(15, 4)
        .Value = "If Any"
        .Interior.Color = RGB(237, 237, 237)
        .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
        .Offset(0, 2).Value = GetUserInfo()
        End With
        End
    End If
End Sub
Private Sub cmbPeople_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.cmbPeople
End Sub
