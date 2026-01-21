Attribute VB_Name = "mod_05_PTB_Highlight"
' ============================================================================
' Module: mod_05_PTB_Highlight
' Project: HRE 연결마스터 (Consolidation Master)
' Migrated from: BEP v1.98
' Migration Date: 2026-01-21
'
' Description: Pre-Trial Balance highlighting and query refresh
' Changes: COPY AS-IS - Column references remain same
' ============================================================================
Option Explicit
Sub QueryRefresh() '쿼리 갱신 및 법인별BSPL 피벗테이블 갱신
    Dim tblPTB As ListObject
    Dim tblVerify As ListObject
    Dim pt As PivotTable

    On Error Resume Next

    If Check.Cells(12, 4).Value <> "Complete" Or Check.Cells(13, 4).Value <> "Complete" Or Check.Cells(14, 4).Value <> "Complete" Or _
       Check.Cells(16, 4).Value <> "Complete" Then
        GoEnd "이전 단계를 완료해주세요!"
    End If
    With Check.Cells(18, 4)
        .Value = "In Progress"
        .Interior.Color = RGB(255, 235, 156)
        .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
        .Offset(0, 2).Value = GetUserInfo()
    End With


    Call SpeedUp
    Call OpenProgress("SPO에서 갱신 중...")

    Set tblPTB = BSPL.ListObjects("PTB")
    Set tblVerify = Verify.ListObjects("재무제표")
    Set pt = CorpBSPL.PivotTables("법인별BSPL")

    BSPL.Unprotect PASSWORD: Verify.Unprotect PASSWORD: CorpBSPL.Unprotect PASSWORD

    Call CalculateProgress(0.5, "SPO로부터 합잔 자료 새로고침 중...")
    tblPTB.QueryTable.Refresh BackgroundQuery:=False


    Call CalculateProgress(0.75, "SPO로부터 BSPL 자료 새로고침 중...")
    tblVerify.QueryTable.Refresh BackgroundQuery:=False
    Application.CalculateUntilAsyncQueriesDone

    pt.RefreshTable

    BSPL.Protect PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True: Verify.Protect PASSWORD, UserInterfaceOnly:=True: CorpBSPL.Protect PASSWORD, UserInterfaceOnly:=True
    Call CalculateProgress(1, "작업 완료")

    Call SpeedDown
    Set tblPTB = Nothing: Set tblVerify = Nothing: Set pt = Nothing
End Sub
Sub HighlightPTB() ' 입력되지 않은 합잔 BSPL 노랑 강조
    Dim tblPTB As ListObject
    Dim rng As Range
    Dim lastCol As Long
    Dim i As Long

    On Error Resume Next
    Call SpeedUp

    Set tblPTB = BSPL.ListObjects("PTB")
    tblPTB.AutoFilter.ShowAllData

    Set rng = tblPTB.DataBodyRange
    lastCol = tblPTB.ListColumns.count

    If rng Is Nothing Or rng.Cells.count = 0 Then
        GoEnd "합잔(PTB) 데이터를 새로고침을 확인해주세요!" & vbNewLine & "또는 해당 법인 전체 체크 없이 입력된 내용을 확인하세요!"
    End If


    ' 시트 보호해제
    BSPL.Unprotect PASSWORD

    For i = 1 To rng.Rows.count
        If IsEmpty(rng.Cells(i, 4)) Then
            rng.Cells(i, 1).Resize(1, lastCol).Interior.Color = vbYellow
        Else
            rng.Cells(i, 1).Resize(1, lastCol).Interior.Color = vbWhite
        End If
    Next i

    BSPL.Protect PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True

    Call SpeedDown
    Set tblPTB = Nothing: Set rng = Nothing
End Sub
Sub FilterPTB() ' 강조된 데이터 필터링
    Dim tblPTB As ListObject
    Dim rng As Range
    Dim i As Long
    Dim hasData As Boolean

    On Error Resume Next
    Call SpeedUp

    Set tblPTB = BSPL.ListObjects("PTB")
    Set rng = tblPTB.DataBodyRange
    isYellow = 0
    hasData = False

    BSPL.Unprotect PASSWORD: AddCoA.Unprotect PASSWORD: ThisWorkbook.Unprotect PASSWORD:=PASSWORD_Workbook

    ' 테이블에 실제 데이터가 있는지 확인
    If Not rng Is Nothing Then
        For i = 1 To rng.Rows.count
            ' 첫 번째 열에 값이 있는지 확인
            If Not IsEmpty(rng.Cells(i, 1).Value) Then
                hasData = True
                ' 값이 있고 노란색일 때만 카운트
                If rng.Cells(i, 1).Interior.Color = vbYellow Then
                    isYellow = isYellow + 1
                End If
            End If
        Next i
    End If

    ' 데이터가 없는 경우 자동으로 완료 처리
    If Not hasData Then
        If tblPTB.AutoFilter.FilterMode Then
            tblPTB.AutoFilter.ShowAllData
        End If
        AddCoA.Visible = xlSheetVeryHidden

        ' 진행상황 표시
        With Check.Cells(18, 4)
            .Value = "Complete"
            .Interior.Color = RGB(198, 239, 206)
            .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
            .Offset(0, 2).Value = GetUserInfo()
         End With

        BSPL.Activate
        BSPL.Range("A1").Select
        MsgBox "데이터가 없습니다. 작업이 완료되었습니다.", vbInformation, AppName & " " & AppType
    ElseIf isYellow > 0 Then
        tblPTB.Range.AutoFilter Field:=1, Criteria1:=RGB(255, 255, 0), Operator:=xlFilterCellColor
        BSPL.Activate
        Call Fill_Input_Table
        AddCoA.Visible = xlSheetVisible
        AddCoA.Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
        AddCoA.Activate
        AddCoA.Range("A1").Select

        BSPL.Activate
        BSPL.Range("A1").Select
        MsgBox "CoA를 업데이트 해주세요." & vbNewLine & _
                "계정 시트 또는 CoA 추가 시트에서 CoA를 추가하세요.", vbInformation, AppName & " " & AppType
    Else
        If tblPTB.AutoFilter.FilterMode Then
            tblPTB.AutoFilter.ShowAllData
        End If
        AddCoA.Visible = xlSheetVeryHidden

        ' 진행상황 표시
        With Check.Cells(18, 4)
            .Value = "Complete"
            .Interior.Color = RGB(198, 239, 206)
            .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
            .Offset(0, 2).Value = GetUserInfo()
         End With

        BSPL.Activate
        BSPL.Range("A1").Select
        MsgBox "작업이 완료되었습니다.", vbInformation, AppName & " " & AppType
    End If

    BSPL.Protect PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True: ThisWorkbook.Protect PASSWORD:=PASSWORD_Workbook
    AddCoA.Cells.Locked = True: AddCoA.Range("E5:G1048576").Locked = False: AddCoA.Protect PASSWORD, UserInterfaceOnly:=True

    Call SpeedDown
    Set tblPTB = Nothing: Set rng = Nothing
End Sub
