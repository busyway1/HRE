Attribute VB_Name = "mod_Ribbon"
Option Explicit
' ============================================================================
' Module: mod_Ribbon
' Project: HRE 연결마스터 (Consolidation Master)
' Version: 1.00
' Date: 2026-01-21
'
' Description: Ribbon callback functions for custom UI
' Changes from BEP:
'  - Added: GetER_Flow, GetER_Spot (exchange rate functions)
'  - Removed: Sum_MC, Sum_MC_AD (Management Consolidation not needed for HRE)
'  - Updated: Ribbon structure to match HRE workflow
' ============================================================================

' ==================== SETUP GROUP ====================

Sub SetSPO(control As IRibbonControl) ' SPO 홈페이지 설정
    'ValidatePermission
    frmSPO.Show
End Sub

Sub SetDirectory(control As IRibbonControl) ' 조직(부서) 설정
    'ValidatePermission
    frmDirectory.Show
End Sub

Sub SetDate(control As IRibbonControl) ' 결산연월 설정
    'ValidatePermission
    frmDate.Show
End Sub

Sub AppendCorp(control As IRibbonControl) ' 법인명 추가
    'ValidatePermission
    frmCorp_Append.Show
End Sub

Sub SetScope(control As IRibbonControl) ' Scope 설정
    'ValidatePermission
    frmScope.Show
End Sub

' ==================== EXCHANGE RATE GROUP (NEW FOR HRE) ====================

Sub GetER_Flow(control As IRibbonControl) ' 평균환율 조회
    'ValidatePermission
    Call mod_17_ExchangeRate.GetER_Flow
End Sub

Sub GetER_Spot(control As IRibbonControl) ' 기말환율 조회
    'ValidatePermission
    Call mod_17_ExchangeRate.GetER_Spot
End Sub

' ==================== COA GROUP ====================

Sub Synchro_CoA(control As IRibbonControl) ' CoA 동기화
    'ValidatePermission
    Call SyncCoA
End Sub

Sub Update(control As IRibbonControl) ' CoA 확인 및 데이터 합산
    'ValidatePermission
    Call QueryRefresh
    Call HighlightPTB
    Call FilterPTB
End Sub

' ==================== VERIFICATION GROUP ====================

Sub Verify_BSPL(control As IRibbonControl) ' 재무제표 검증
    'ValidatePermission
    Application.DisplayAlerts = False
    Call RefreshPivotVerify
    Call VerifyBS
    Call VerifyIS
    Call ValidateCorpCodes
    Call ValidateSheetColors
    Application.DisplayAlerts = True
End Sub

Sub Update_AD(control As IRibbonControl) ' CoA 확인 및 데이터 합산(취득, 처분)
    'ValidatePermission
    Call QueryRefresh_ADBS
    Call Highlight_ADBS
    Call Filter_ADBS
End Sub

Sub Verify_AD(control As IRibbonControl) ' 검증 실행(취득, 처분)
    'ValidatePermission
    Application.DisplayAlerts = False
    Call RefreshPivot_ADBS
    Call Verify_ADBS_Acq
    Call Verify_ADBS_Dis
    Call Verify_ADPL
    Call ValidateCorp_ADBS
    Call ValidateSheetColors_ADBS
    Application.DisplayAlerts = True
End Sub

Sub Verify_Master(control As IRibbonControl) ' CoA 마스터 검증
    'ValidatePermission
    Call VerifyMaster
End Sub

' ==================== UTILITY GROUP ====================

Sub FilterSheet(control As IRibbonControl) ' 필터링
    Dim ws As Worksheet
    Set ws = ActiveSheet
    'ValidatePermission
    If ws.Name = "법인별 CoA" Or ws.Name = "CoA 마스터" Or ws.Name = "합산 BSPL" Or ws.Name = "취득, 처분 BS" Then
        Call DoFilter
    ElseIf ws.Name = "CoA 마스터" Then
        Call DoFilter_Master
    Else
        GoEnd "해당 시트에서는 필터링 기능을 사용할 수 없습니다."
    End If
    Set ws = Nothing
End Sub

Sub UnfilterSheet(control As IRibbonControl) ' 필터링 해제
    Dim ws As Worksheet
    Set ws = ActiveSheet
    'ValidatePermission
    If ws.Name = "법인별 CoA" Or ws.Name = "CoA 마스터" Or ws.Name = "합산 BSPL" Or ws.Name = "취득, 처분 BS" Then
        Call UndoFilter
    ElseIf ws.Name = "CoA 마스터" Then
        Call UndoFilter_Master
    Else
        GoEnd "해당 시트에서는 필터 해제를 실행할 수 없습니다."
    End If
    Set ws = Nothing
End Sub

Sub ProtectQuery(control As IRibbonControl) ' 쿼리 잠금
    'ValidatePermission
    Call ProtectQueryEditor
End Sub

Sub UnprotectQuery(control As IRibbonControl) ' 쿼리 잠금 해제
    'ValidatePermission
    Call UnprotectQueryEditor
End Sub

Sub Refresh_Data(control As IRibbonControl) ' 쿼리 새로고침
    'ValidatePermission
    Call RefreshAllData
End Sub

Sub Manage_People(control As IRibbonControl) ' 사용자 관리
    'ValidatePermission
    frmPeople.Show
End Sub

' ==================== EXPORT GROUP ====================

Sub Export_File(control As IRibbonControl) ' 파일 내보내기 (엑셀)
    'ValidatePermission
    Call Export_Master
End Sub

' ==================== INFO GROUP ====================

Sub IRVersion(control As IRibbonControl) ' 버전 확인
    Msg "현재 버전: " & AppVersion & vbNewLine & "배포일: " & RelDate & vbNewLine & "만료일: " & ExpDate
End Sub

Sub IRSPO(control As IRibbonControl) ' SPO 사이트 열기
    Call OpenSPO
End Sub

Sub IRBugReport(control As IRibbonControl) ' 버그 리포트
    Call OpenGoogleForm
End Sub

Sub IRManual(control As IRibbonControl) ' 매뉴얼 열기
    Call OpenManual
End Sub
