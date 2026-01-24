Attribute VB_Name = "mod_05_PTB_Highlight"
' ============================================================================
' Module: mod_05_PTB_Highlight
' Project: HRE ���Ḷ���� (Consolidation Master)
' Migrated from: BEP v1.98
' Migration Date: 2026-01-21
'
' Description: Pre-Trial Balance highlighting and query refresh
' Changes: COPY AS-IS - Column references remain same
' ============================================================================
Option Explicit
Sub QueryRefresh() '���� ���� �� ���κ�BSPL �ǹ����̺� ����
    Dim tblPTB As ListObject
    Dim tblVerify As ListObject
    Dim pt As PivotTable

    On Error Resume Next

    If Check.Cells(12, 4).Value <> "Complete" Or Check.Cells(13, 4).Value <> "Complete" Or Check.Cells(14, 4).Value <> "Complete" Or _
       Check.Cells(16, 4).Value <> "Complete" Then
        GoEnd "���� �ܰ踦 �Ϸ����ּ���!"
    End If
    With Check.Cells(18, 4)
        .Value = "In Progress"
        .Interior.Color = RGB(255, 235, 156)
        .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
        .Offset(0, 2).Value = GetUserInfo()
    End With


    Call SpeedUp
    Call OpenProgress("SPO���� ���� ��...")

    Set tblPTB = BSPL.ListObjects("PTB")
    Set tblVerify = Verify.ListObjects("�繫��ǥ")
    Set pt = CorpBSPL.PivotTables("���κ�BSPL")

    BSPL.Unprotect PASSWORD: Verify.Unprotect PASSWORD: CorpBSPL.Unprotect PASSWORD

    Call CalculateProgress(0.5, "SPO�κ��� ���� �ڷ� ���ΰ�ħ ��...")
    tblPTB.QueryTable.Refresh BackgroundQuery:=False


    Call CalculateProgress(0.75, "SPO�κ��� BSPL �ڷ� ���ΰ�ħ ��...")
    tblVerify.QueryTable.Refresh BackgroundQuery:=False
    Application.CalculateUntilAsyncQueriesDone

    pt.RefreshTable

    BSPL.Protect PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True: Verify.Protect PASSWORD, UserInterfaceOnly:=True: CorpBSPL.Protect PASSWORD, UserInterfaceOnly:=True
    Call CalculateProgress(1, "�۾� �Ϸ�")

    Call SpeedDown
    Set tblPTB = Nothing: Set tblVerify = Nothing: Set pt = Nothing
End Sub
Sub HighlightPTB() ' �Էµ��� ���� ���� BSPL ��� ����
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
        GoEnd "����(PTB) �����͸� ���ΰ�ħ�� Ȯ�����ּ���!" & vbNewLine & "�Ǵ� �ش� ���� ��ü üũ ���� �Էµ� ������ Ȯ���ϼ���!"
    End If


    ' ��Ʈ ��ȣ����
    BSPL.Unprotect PASSWORD

    ' HRE v1.00: PwC_CoA는 5번째 컬럼 (법인명 컬럼 추가로 인함)
    ' PTB 컬럼 순서: 법인코드(1), 법인명(2), 법인별CoA(3), 법인별계정과목명(4), PwC_CoA(5), ...
    For i = 1 To rng.Rows.count
        If IsEmpty(rng.Cells(i, 5)) Then
            rng.Cells(i, 1).Resize(1, lastCol).Interior.Color = vbYellow
        Else
            rng.Cells(i, 1).Resize(1, lastCol).Interior.Color = vbWhite
        End If
    Next i

    BSPL.Protect PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True

    Call SpeedDown
    Set tblPTB = Nothing: Set rng = Nothing
End Sub
Sub FilterPTB() ' ������ ������ ���͸�
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

    ' ���̺��� ���� �����Ͱ� �ִ��� Ȯ��
    If Not rng Is Nothing Then
        For i = 1 To rng.Rows.count
            ' ù ��° ���� ���� �ִ��� Ȯ��
            If Not IsEmpty(rng.Cells(i, 1).Value) Then
                hasData = True
                ' ���� �ְ� ������� ���� ī��Ʈ
                If rng.Cells(i, 1).Interior.Color = vbYellow Then
                    isYellow = isYellow + 1
                End If
            End If
        Next i
    End If

    ' �����Ͱ� ���� ��� �ڵ����� �Ϸ� ó��
    If Not hasData Then
        If tblPTB.AutoFilter.FilterMode Then
            tblPTB.AutoFilter.ShowAllData
        End If
        AddCoA.Visible = xlSheetVeryHidden

        ' �����Ȳ ǥ��
        With Check.Cells(18, 4)
            .Value = "Complete"
            .Interior.Color = RGB(198, 239, 206)
            .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
            .Offset(0, 2).Value = GetUserInfo()
         End With

        BSPL.Activate
        BSPL.Range("B1").Select
        MsgBox "�����Ͱ� �����ϴ�. �۾��� �Ϸ�Ǿ����ϴ�.", vbInformation, AppName & " " & AppType
    ElseIf isYellow > 0 Then
        tblPTB.Range.AutoFilter Field:=1, Criteria1:=RGB(255, 255, 0), Operator:=xlFilterCellColor
        BSPL.Activate
        Call Fill_Input_Table
        AddCoA.Visible = xlSheetVisible
        AddCoA.Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
        AddCoA.Activate
        AddCoA.Range("B1").Select

        BSPL.Activate
        BSPL.Range("B1").Select
        MsgBox "CoA�� ������Ʈ ���ּ���." & vbNewLine & _
                "���� ��Ʈ �Ǵ� CoA �߰� ��Ʈ���� CoA�� �߰��ϼ���.", vbInformation, AppName & " " & AppType
    Else
        If tblPTB.AutoFilter.FilterMode Then
            tblPTB.AutoFilter.ShowAllData
        End If
        AddCoA.Visible = xlSheetVeryHidden

        ' �����Ȳ ǥ��
        With Check.Cells(18, 4)
            .Value = "Complete"
            .Interior.Color = RGB(198, 239, 206)
            .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
            .Offset(0, 2).Value = GetUserInfo()
         End With

        BSPL.Activate
        BSPL.Range("B1").Select
        MsgBox "�۾��� �Ϸ�Ǿ����ϴ�.", vbInformation, AppName & " " & AppType
    End If

    BSPL.Protect PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True: ThisWorkbook.Protect PASSWORD:=PASSWORD_Workbook
    AddCoA.Cells.Locked = True: AddCoA.Range("E5:G1048576").Locked = False: AddCoA.Protect PASSWORD, UserInterfaceOnly:=True

    Call SpeedDown
    Set tblPTB = Nothing: Set rng = Nothing
End Sub
