Attribute VB_Name = "mod_Refresh"
' ============================================================================
' Module: mod_Refresh
' Project: HRE 연결마스터 (Consolidation Master)
' Migrated from: BEP v1.98
' Migration Date: 2026-01-21
'
' Description: Refresh all data functionality
' Changes: COPY AS-IS - No HRE-specific adaptations required
' ============================================================================
Option Explicit
Sub RefreshAllData()
    On Error Resume Next
    Call SpeedUp
    ThisWorkbook.RefreshAll

    MsgBox "새로고침이 완료되었습니다.", vbInformation, AppName & " " & AppType

    Call SpeedDown
End Sub
