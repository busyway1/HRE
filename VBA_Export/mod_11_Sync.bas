Attribute VB_Name = "mod_11_Sync"
Option Explicit
' ============================================================================
' Module: mod_11_Sync
' Project: HRE 연결마스터 (Consolidation Master)
' Migrated from: BEP v1.98
' Migration Date: 2026-01-21
'
' Description: CoA synchronization between files
' Changes: COPY AS-IS - No HRE-specific adaptations required
' ============================================================================
Sub SyncCoA()
    Dim fd As FileDialog
    Dim strFile As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsHide As Worksheet
    Dim wsCoA As Worksheet
    Dim wsTarget As Worksheet
    Dim tblSource As ListObject
    Dim tblTarget As ListObject
    Dim sourceRow As ListRow
    Dim targetRow As ListRow
    Dim addedCount As Long
    Dim dict As Object
    Dim i As Long

    On Error Resume Next
    Set dict = CreateObject("Scripting.Dictionary")
    Call SpeedUp

    '파일 선택
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "연결마스터에서 파일 선택"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsm"
        .AllowMultiSelect = False
        If .Show = -1 Then
            strFile = .selectedItems(1)
        Else
            Msg "파일 선택이 취소되었습니다.", vbInformation
            Call SpeedDown
            Exit Sub
        End If
    End With

    '현재 열려있는 파일과 같은 파일인지 확인
    Dim selectedFileName As String
    selectedFileName = Right(strFile, Len(strFile) - InStrRev(strFile, "\"))

    For Each wb In Workbooks
        If wb.name = selectedFileName Then
            Msg "현재 열려있는 파일과 같은 파일입니다." & vbNewLine & "다른 파일을 선택해주세요.", vbExclamation
            Call SpeedDown
            Exit Sub
        End If
    Next wb

    '파일 열기
    Set wb = Workbooks.Open(strFile)
    If wb Is Nothing Then
        Msg "파일을 열 수 없습니다.", vbExclamation
        Call SpeedDown
        Exit Sub
    End If

    '필요한 시트 찾기
    For Each ws In wb.Worksheets
        Select Case ws.CodeName
            Case "HideSheet"
                Set wsHide = ws
            Case "CorpCoA"
                Set wsCoA = ws
        End Select

        If Not (wsHide Is Nothing) And Not (wsCoA Is Nothing) Then
            Exit For
        End If
    Next ws

    '시트 존재 여부 확인
    If wsHide Is Nothing Then
        wb.Close SaveChanges:=False
        Msg "올바른 연결마스터 파일을 선택해주세요!", vbExclamation
        Call SpeedDown
        Exit Sub
    End If

    If wsCoA Is Nothing Then
        wb.Close SaveChanges:=False
        Msg "파일에서 법인별 CoA 시트를 찾을 수 없습니다.", vbExclamation
        Call SpeedDown
        Exit Sub
    End If

    Set tblSource = wsCoA.ListObjects("Raw_CoA")
    Set tblTarget = CorpCoA.ListObjects("Raw_CoA")

    '테이블 존재 여부 확인
    If tblSource Is Nothing Or tblTarget Is Nothing Then
        wb.Close SaveChanges:=False
        Msg "Raw_CoA 테이블을 찾을 수 없습니다.", vbExclamation
        Call SpeedDown
        Exit Sub
    End If

    '타겟 테이블의 키를 Dictionary에 저장
    Dim keyVal As String
    For Each targetRow In tblTarget.ListRows
        keyVal = CStr(targetRow.Range.Cells(1, 1).Value) & "|" & CStr(targetRow.Range.Cells(1, 2).Value)
        dict(keyVal) = True
    Next targetRow

    '소스 테이블에서 키 확인 후 없는 경우 추가
    addedCount = 0
    i = 1
    Call OpenProgress("CoA 동기화 중...")

    CorpCoA.Unprotect PASSWORD

    For Each sourceRow In tblSource.ListRows
        Call CalculateProgress(i / tblSource.ListRows.count, "CoA 동기화 중...")
        keyVal = CStr(sourceRow.Range.Cells(1, 1).Value) & "|" & CStr(sourceRow.Range.Cells(1, 2).Value)
        If Not dict.Exists(keyVal) Then
            With tblTarget.ListRows.Add
                .Range.Value = sourceRow.Range.Value
                ' 비고란 가장 마지막 열의 값 & 법 추가
                .Range.Cells(1, .Range.Columns.count).Value = _
                .Range.Cells(1, .Range.Columns.count).Value & "from " & wsHide.Range("U2") & " 추가"
            End With
            dict(keyVal) = True
            addedCount = addedCount + 1
        End If
        i = i + 1
    Next sourceRow

    wb.Close SaveChanges:=False
    Msg addedCount & "개의 항목이 추가되었습니다.", vbInformation

    CorpCoA.Protect PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True

    Set wb = Nothing: Set fd = Nothing: Set ws = Nothing: Set wsHide = Nothing: Set wsCoA = Nothing: Set wsTarget = Nothing
    Set tblSource = Nothing: Set tblTarget = Nothing: Set sourceRow = Nothing: Set targetRow = Nothing: Set dict = Nothing
    Call SpeedDown
End Sub
