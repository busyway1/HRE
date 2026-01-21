Attribute VB_Name = "mod_01_FilterSearch"
' ============================================================================
' Module: mod_01_FilterSearch
' Project: HRE 연결마스터 (Consolidation Master)
' Migrated from: BEP v1.98
' Migration Date: 2026-01-21
'
' Description: Table filtering and search functionality
' Changes: COPY AS-IS - No HRE-specific adaptations required
' ============================================================================
Option Explicit
Sub DoFilter()
    Dim col As ListColumn
    Dim comboBox As MSForms.comboBox
    Dim userResponse As VbMsgBoxResult
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

    ' 보이는 행 수 계산 (필터 적용)
    visibleRowsCount = GetVisibleRowsCount(tbl)

    If tbl.ShowAutoFilter And visibleRowsCount < totalRowsCount Then
        userResponse = MsgBox("이미 필터링이 적용되어 있습니다. 해제 하시겠습니까?" & vbNewLine & vbNewLine & _
                              "예 - 현재 필터에 추가 필터링" & vbNewLine & _
                              "아니요 - 해제 후에 새로 된 후 새로 필터링" & vbNewLine & _
                              "취소 - 작업 취소", _
                              vbYesNoCancel + vbQuestion, AppName & " " & AppType)

        Select Case userResponse
            Case vbYes
                ' 현재 필터 유지, 추가 필터링 진행
            Case vbNo
                ' 필터 해제 진행
                tbl.AutoFilter.ShowAllData
            Case vbCancel
                GoEnd
        End Select
    End If

    Set comboBox = frmFilter.Controls("HeaderName_Combo")
    comboBox.Clear

    For Each col In tbl.ListColumns
        comboBox.AddItem col.name
    Next col

    If comboBox.ListCount > 0 Then
        comboBox.ListIndex = 0
    End If

    frmFilter.Show

    Call SpeedDown
    Set ws = Nothing: Set tbl = Nothing
End Sub
Private Function GetVisibleRowsCount(tbl As ListObject) As Long
    Dim dataRange As Range
    Dim row As Range
    Dim count As Long

    Set dataRange = tbl.DataBodyRange
    count = 0

    For Each row In dataRange.Rows
        If Not row.Hidden Then
            count = count + 1
        End If
    Next row

    GetVisibleRowsCount = count
End Function
Sub FilterTable()
    Dim selectedColumn As String
    Dim keywordText As String
    Dim filterRange As Range
    Dim dataColumn As Range
    Dim dataArr As Variant
    Dim i As Long
    Dim isAllNumeric As Boolean
    Dim numericValue As Double
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim filterEmpty As Boolean

    On Error Resume Next
    Call SpeedUp

    Set ws = ActiveSheet
    Set tbl = ws.ListObjects(1)

    selectedColumn = frmFilter.Controls("HeaderName_Combo").Value
    keywordText = frmFilter.Controls("Keyword_Text").Value
    filterEmpty = frmFilter.Controls("chkEmptyValues").Value

    Set filterRange = tbl.ListColumns(selectedColumn).Range
    Set dataColumn = filterRange.Offset(1, 0).Resize(filterRange.Rows.count - 1)

    dataArr = dataColumn.Value

    isAllNumeric = True
    For i = 1 To UBound(dataArr)
        If Not IsEmpty(dataArr(i, 1)) Then
            If Not IsNumeric(dataArr(i, 1)) Then
                isAllNumeric = False
                Exit For
            End If
        End If
    Next i

    If keywordText <> "" Or filterEmpty Then
        If filterEmpty Then
            If keywordText <> "" Then
                If isAllNumeric And IsNumeric(keywordText) Then
                    numericValue = CDbl(keywordText)
                    filterRange.AutoFilter Field:=filterRange.Column - tbl.Range.Column + 1, _
                                           Criteria1:="=" & numericValue, _
                                           Operator:=xlOr, _
                                           Criteria2:="="
                Else
                    filterRange.AutoFilter Field:=filterRange.Column - tbl.Range.Column + 1, _
                                           Criteria1:="=*" & keywordText & "*", _
                                           Operator:=xlOr, _
                                           Criteria2:="="
                End If
            Else
                filterRange.AutoFilter Field:=filterRange.Column - tbl.Range.Column + 1, _
                                       Criteria1:="="
            End If
        Else
            If isAllNumeric And IsNumeric(keywordText) Then
                numericValue = CDbl(keywordText)
                filterRange.AutoFilter Field:=filterRange.Column - tbl.Range.Column + 1, _
                                       Criteria1:="=" & numericValue, _
                                       Operator:=xlAnd
            Else
                filterRange.AutoFilter Field:=filterRange.Column - tbl.Range.Column + 1, _
                                       Criteria1:="=*" & keywordText & "*", _
                                       Operator:=xlAnd
            End If
        End If
    Else
        If tbl.ShowAutoFilter Then
            tbl.AutoFilter.ShowAllData
        End If
    End If

    MsgBox "필터링이 완료되었습니다.", vbInformation, AppName & " " & AppType
    Unload frmFilter
    Call SpeedDown

    Set ws = Nothing: Set tbl = Nothing
End Sub
Sub UndoFilter()
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
        visibleRowsCount = GetVisibleRowsCount(tbl)

        If visibleRowsCount < totalRowsCount Then
            tbl.AutoFilter.ShowAllData
            MsgBox "필터링이 해제되었습니다.", vbInformation, AppName & " " & AppType
        Else
            MsgBox "필터링이 이미 해제되어 있습니다.", vbExclamation, AppName & " " & AppType
        End If
    Else
        MsgBox "현재 필터가 적용되어 있지 않습니다.", vbExclamation, AppName & " " & AppType
    End If

    Call SpeedDown
    Set ws = Nothing: Set tbl = Nothing
End Sub
