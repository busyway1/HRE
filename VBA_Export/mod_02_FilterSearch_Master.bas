Attribute VB_Name = "mod_02_FilterSearch_Master"
' ============================================================================
' Module: mod_02_FilterSearch_Master
' Project: HRE 연결마스터 (Consolidation Master)
' Migrated from: BEP v1.98
' Migration Date: 2026-01-21
'
' Description: Master table filtering functionality
' Changes: COPY AS-IS - No HRE-specific adaptations required
' ============================================================================
Option Explicit
Sub DoFilter_Master()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim visibleRowsCount As Long
    Dim totalRowsCount As Long
    Dim userResponse As VbMsgBoxResult
    On Error Resume Next
    Call SpeedUp

    Set ws = ActiveSheet
    Set tbl = ws.ListObjects(1)

    totalRowsCount = tbl.DataBodyRange.Rows.count
    visibleRowsCount = GetVisibleRowsCount_Master(tbl)

    If tbl.ShowAutoFilter And visibleRowsCount < totalRowsCount Then
        userResponse = MsgBox("이미 필터링이 적용되어 있습니다. 해제 하시겠습니까?" & vbNewLine & vbNewLine & _
                              "예 - 현재 필터에 추가 필터링" & vbNewLine & _
                              "아니요 - 해제 후에 새로 된 후 새로 필터링" & vbNewLine & _
                              "취소 - 작업 취소", _
                              vbYesNoCancel + vbQuestion, "필터 확인")

        Select Case userResponse
            Case vbYes
                ' 현재 필터 유지, 추가 필터링 진행
                frmFilter_Master.Show
            Case vbNo
                ' 필터 해제 진행
                tbl.AutoFilter.ShowAllData
                frmFilter_Master.Show
            Case vbCancel
                GoEnd
        End Select
    Else
        frmFilter_Master.Show
    End If

    Call SpeedDown
    Set ws = Nothing: Set tbl = Nothing
End Sub
Function GetVisibleRowsCount_Master(tbl As ListObject) As Long
    Dim visibleRange As Range
    On Error Resume Next
    Set visibleRange = tbl.DataBodyRange.SpecialCells(xlCellTypeVisible)
    If visibleRange Is Nothing Then
        GetVisibleRowsCount_Master = 0
    Else
        GetVisibleRowsCount_Master = visibleRange.Rows.count
    End If
End Function
Sub UndoFilter_Master()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim visibleRowsCount As Long
    Dim totalRowsCount As Long
    On Error Resume Next
    Call SpeedUp

    Set ws = ActiveSheet
    Set tbl = ws.ListObjects(1)

    ' 전체 데이터 행 수 계산
    totalRowsCount = tbl.DataBodyRange.Rows.count

    If tbl.ShowAutoFilter Then
        ' 보이는 행 수 계산 (필터 적용)
        visibleRowsCount = GetVisibleRowsCount_Master(tbl)

        If visibleRowsCount < totalRowsCount Then
            tbl.AutoFilter.ShowAllData
            MsgBox "필터링이 해제되었습니다.", vbInformation, "완료"
        Else
            MsgBox "필터링이 이미 해제되어 있습니다.", vbExclamation
        End If
    Else
        MsgBox "현재 필터가 적용되어 있지 않습니다.", vbExclamation
    End If

    Call SpeedDown
    Set ws = Nothing: Set tbl = Nothing
End Sub
