Attribute VB_Name = "mod_09_CheckMaster"
' ============================================================================
' Module: mod_09_CheckMaster
' Project: HRE 연결마스터 (Consolidation Master)
' Migrated from: BEP v1.98
' Migration Date: 2026-01-21
'
' Description: Master data validation
' Changes: COPY AS-IS - No HRE-specific adaptations required
' ============================================================================
Option Explicit
Sub VerifyMaster()
    On Error Resume Next
    Dim tbl As ListObject
    Dim dataRange As Range
    Dim i As Long

    If Check.Cells(12, 4).Value <> "Complete" Or Check.Cells(13, 4).Value <> "Complete" Or Check.Cells(14, 4).Value <> "Complete" Or _
       Check.Cells(16, 4).Value <> "Complete" Or Check.Cells(18, 4).Value <> "Complete" Or Check.Cells(20, 4).Value <> "Complete" Or _
       Check.Cells(21, 4).Value <> "Complete" Or Check.Cells(22, 4).Value <> "Complete" Then
        GoEnd "이전 단계를 완료해주세요!"
    End If
    With Check.Cells(23, 4)
        .Value = "In Progress"
        .Interior.Color = RGB(255, 235, 156)
        .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
        .Offset(0, 2).Value = GetUserInfo()
    End With

    Call SpeedUp

    Set tbl = CoAMaster.ListObjects("Master")
    Set dataRange = tbl.DataBodyRange

    dataRange.Interior.ColorIndex = xlNone

    For i = 1 To dataRange.Rows.count
        ' 금액이 0이 아니면서 계정을 입력하지 않았는지 체크
        If dataRange.Cells(i, 10).Value <> 0 And _
           (IsEmpty(dataRange.Cells(i, 7)) Or _
            IsEmpty(dataRange.Cells(i, 8))) Then
            dataRange.Rows(i).Interior.Color = RGB(255, 254, 0)
        End If
    Next i
    For i = 1 To dataRange.Rows.count
        If dataRange.Rows(i).Interior.Color = RGB(255, 254, 0) Then
            GoEnd "일부 계정이 입력되지 않았습니다!"
        End If
    Next i


    With Check.Cells(23, 4)
        .Value = "Complete"
        .Interior.Color = RGB(198, 239, 206)
        .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
        .Offset(0, 2).Value = GetUserInfo()
    End With


    Msg "CoA 마스터 잔액 검증이 완료되었습니다!", vbInformation

    Call SpeedDown
    Set tbl = Nothing: Set dataRange = Nothing
End Sub
