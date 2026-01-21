Attribute VB_Name = "mod_04_IntializeProgress"
' ============================================================================
' Module: mod_04_IntializeProgress
' Project: HRE 연결마스터 (Consolidation Master)
' Migrated from: BEP v1.98
' Migration Date: 2026-01-21
'
' Description: Progress tracking initialization
' Changes: COPY AS-IS - No HRE-specific adaptations required
' ============================================================================
Option Explicit
Sub DeleteProgress()
    Dim response As VbMsgBoxResult
    On Error Resume Next
    Call SpeedUp

    response = MsgBox("진행현황을 초기화합니다. 진행하시겠습니까?", _
                     vbYesNo + vbQuestion, AppName & " " & AppType)

    Select Case response
        Case vbYes
            Check.Unprotect PASSWORD

            With Range(Check.Cells(12, 4), Check.Cells(14, 4))
                .Value = "Not Started"
                .Interior.Color = RGB(255, 199, 206)
            End With

            With Check.Cells(15, 4)
                .Value = "If Any"
                .Interior.Color = RGB(237, 237, 237)
            End With

            With Check.Cells(16, 4)
                .Value = "Not Started"
                .Interior.Color = RGB(255, 199, 206)
            End With

            With Check.Cells(17, 4)
                .Value = "If Any"
                .Interior.Color = RGB(237, 237, 237)
            End With

             With Check.Cells(18, 4)
                .Value = "Not Started"
                .Interior.Color = RGB(255, 199, 206)
            End With

            With Check.Cells(19, 4)
                .Value = "If Any"
                .Interior.Color = RGB(237, 237, 237)
            End With

            With Check.Cells(20, 4)
                .Value = "Not Started"
                .Interior.Color = RGB(255, 199, 206)
            End With

             With Check.Cells(21, 4)
                .Value = "Not Started"
                .Interior.Color = RGB(255, 199, 206)
            End With

             With Check.Cells(22, 4)
                .Value = "Not Started"
                .Interior.Color = RGB(255, 199, 206)
            End With

            With Check.Cells(23, 4)
                .Value = "Not Started"
                .Interior.Color = RGB(255, 199, 206)
            End With

            With Range(Check.Cells(12, 5), Check.Cells(23, 7))
                .ClearContents
            End With

            Check.Protect PASSWORD, UserInterfaceOnly:=True

            MsgBox "진행현황이 초기화되었습니다.", vbInformation, AppName & " " & AppType

        Case vbNo
            GoEnd
        End Select

        Call SpeedDown
End Sub
