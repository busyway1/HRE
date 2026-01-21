Attribute VB_Name = "mod_QueryProtection"
' ============================================================================
' Module: mod_QueryProtection
' Project: HRE 연결마스터 (Consolidation Master)
' Migrated from: BEP v1.98
' Migration Date: 2026-01-21
'
' Description: Query editor protection functionality
' Changes: COPY AS-IS - No HRE-specific adaptations required
' ============================================================================
Option Explicit
Private isLocked As Long

Public Sub ProtectQueryEditor()
    On Error Resume Next
    ThisWorkbook.Protect Structure:=True, Windows:=False, PASSWORD:=PASSWORD_Workbook
    DisableQueryEditorButtons

    If isLocked = 0 Then
        isLocked = 2
    ElseIf isLocked = 1 Then
        Msg "쿼리 편집기가 잠겼습니다.", vbInformation
        isLocked = 2
    Else
        Msg "이미 쿼리 편집기가 잠겨있습니다.", vbInformation
    End If
End Sub

Public Sub UnprotectQueryEditor()
    Dim userPassword As String
    On Error Resume Next

    If isLocked <> 1 Then
        userPassword = InputBox("쿼리 편집기 보호를 해제하려면 비밀번호를 입력하세요:", "쿼리 편집기 잠금 해제")
        If userPassword = PASSWORD_Workbook Then
            ThisWorkbook.Unprotect PASSWORD:=PASSWORD_Workbook
            EnableQueryEditorButtons
            Msg "쿼리 편집기 보호가 해제되었습니다.", vbInformation
            isLocked = 1
        Else
            Msg "잘못된 비밀번호입니다.", vbExclamation
        End If
    Else
        Msg "이미 쿼리 편집기 보호가 해제되어있습니다", vbInformation
    End If
End Sub
Private Sub DisableQueryEditorButtons()
    On Error Resume Next
    Application.CommandBars("Queries and Connections").Enabled = False
    Dim ctrl As CommandBarControl
    For Each ctrl In Application.CommandBars("Ribbon").Controls
        If ctrl.ID = "TabGetData" Then
            ctrl.Enabled = False
            Exit For
        End If
    Next ctrl
End Sub

Private Sub EnableQueryEditorButtons()
    On Error Resume Next
    Application.CommandBars("Queries and Connections").Enabled = True
    Dim ctrl As CommandBarControl
    For Each ctrl In Application.CommandBars("Ribbon").Controls
        If ctrl.ID = "TabGetData" Then
            ctrl.Enabled = True
            Exit For
        End If
    Next ctrl
End Sub
Public Sub ToggleQueryEditorProtection()
    On Error Resume Next
    If ThisWorkbook.ProtectStructure Then
        UnprotectQueryEditor
    Else
        ProtectQueryEditor
    End If
End Sub
