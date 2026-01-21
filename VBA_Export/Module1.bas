Attribute VB_Name = "Module1"
' ============================================================================
' Module: Module1
' Project: HRE 연결마스터 (Consolidation Master)
' Migrated from: BEP v1.98
' Migration Date: 2026-01-21
'
' Description: VBA component export utility
' Changes: COPY AS-IS - No HRE-specific adaptations required
' ============================================================================
Option Explicit

Sub ExportAllVbaComponents()
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim exportPath As String
    Dim fileName As String

    ' === 1. Export 경로 설정 ===
    ' 필요시 경로 변경 필요 (맨 끝에 \ 필수!)
    exportPath = "C:\Users\Public\VBA_Export\"

    ' 폴더 없으면 생성
    If dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If

    Set vbProj = ThisWorkbook.VBProject

    ' === 2. 모든 컴포넌트 순회 ===
    For Each vbComp In vbProj.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule        ' 표준 모듈 (.bas)
                fileName = exportPath & vbComp.name & ".bas"

            Case vbext_ct_ClassModule      ' 클래스 모듈 (.cls)
                fileName = exportPath & vbComp.name & ".cls"

            Case vbext_ct_MSForm           ' UserForm (.frm)
                fileName = exportPath & vbComp.name & ".frm"

            Case vbext_ct_Document         ' 워크시트 / ThisWorkbook 코드
                ' 워크시트, ThisWorkbook는 그대로 export 가능 (확장자는 임의)
                fileName = exportPath & vbComp.name & "_code.bas"

            Case Else
                ' 기타 타입은 건너뜀
                fileName = ""
        End Select

        If fileName <> "" Then
            vbComp.Export fileName
        End If
    Next vbComp

    MsgBox "VBA 코드 Export 완료! 경로: " & exportPath, vbInformation
End Sub
