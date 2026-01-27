Attribute VB_Name = "mod_Ribbon"
Option Explicit
' ============================================================================
' Module: mod_Ribbon
' Project: HRE ���Ḷ���� (Consolidation Master)
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

Sub SetSPO(control As IRibbonControl) ' SPO Ȩ������ ����
    'ValidatePermission
    frmSPO.Show
End Sub

Sub SetDirectory(control As IRibbonControl) ' ����(�μ�) ����
    'ValidatePermission
    frmDirectory.Show
End Sub

Sub SetDate(control As IRibbonControl) ' ��꿬�� ����
    'ValidatePermission
    frmDate.Show
End Sub

Sub AppendCorp(control As IRibbonControl) ' ���θ� �߰�
    'ValidatePermission
    frmCorp_Append.Show
End Sub

Sub SetScope(control As IRibbonControl) ' Scope ����
    'ValidatePermission
    frmScope.Show
End Sub

' ==================== EXCHANGE RATE GROUP (NEW FOR HRE) ====================
' GetER_Flow and GetER_Spot are defined in mod_17_ExchangeRate
' Ribbon calls them directly via onAction

' ==================== COA GROUP ====================

Sub Synchro_CoA(control As IRibbonControl) ' CoA ����ȭ
    'ValidatePermission
    Call SyncCoA
End Sub

Sub Update(control As IRibbonControl) ' CoA Ȯ�� �� ������ �ջ�
    'ValidatePermission
    Call QueryRefresh
    Call HighlightPTB
    Call FilterPTB
End Sub

' ==================== VERIFICATION GROUP ====================

Sub Verify_BSPL(control As IRibbonControl) ' �繫��ǥ ����
    'ValidatePermission
    Application.DisplayAlerts = False
    Call RefreshPivotVerify
    Call VerifyBS
    Call VerifyIS
    Call ValidateCorpCodes
    Call ValidateSheetColors
    Application.DisplayAlerts = True
End Sub

Sub Update_AD(control As IRibbonControl) ' CoA Ȯ�� �� ������ �ջ�(���, ó��)
    'ValidatePermission
    Call QueryRefresh_ADBS
    Call Highlight_ADBS
    Call Filter_ADBS
End Sub

Sub Verify_AD(control As IRibbonControl) ' ���� ����(���, ó��)
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

Sub Verify_Master(control As IRibbonControl) ' CoA ������ ����
    'ValidatePermission
    Call VerifyMaster
End Sub

' ==================== UTILITY GROUP ====================

Sub FilterSheet(control As IRibbonControl) ' ���͸�
    Dim ws As Worksheet
    Set ws = ActiveSheet
    'ValidatePermission
    If ws.Name = "���κ� CoA" Or ws.Name = "CoA ������" Or ws.Name = "�ջ� BSPL" Or ws.Name = "���, ó�� BS" Then
        Call DoFilter
    ElseIf ws.Name = "CoA ������" Then
        Call DoFilter_Master
    Else
        GoEnd "�ش� ��Ʈ������ ���͸� ����� ����� �� �����ϴ�."
    End If
    Set ws = Nothing
End Sub

Sub UnfilterSheet(control As IRibbonControl) ' ���͸� ����
    Dim ws As Worksheet
    Set ws = ActiveSheet
    'ValidatePermission
    If ws.Name = "���κ� CoA" Or ws.Name = "CoA ������" Or ws.Name = "�ջ� BSPL" Or ws.Name = "���, ó�� BS" Then
        Call UndoFilter
    ElseIf ws.Name = "CoA ������" Then
        Call UndoFilter_Master
    Else
        GoEnd "�ش� ��Ʈ������ ���� ������ ������ �� �����ϴ�."
    End If
    Set ws = Nothing
End Sub

Sub ProtectQuery(control As IRibbonControl) ' ���� ���
    'ValidatePermission
    Call ProtectQueryEditor
End Sub

Sub UnprotectQuery(control As IRibbonControl) ' ���� ��� ����
    'ValidatePermission
    Call UnprotectQueryEditor
End Sub

Sub Refresh_Data(control As IRibbonControl) ' ���� ���ΰ�ħ
    'ValidatePermission
    Call RefreshAllData
End Sub

Sub Manage_People(control As IRibbonControl) ' ����� ����
    'ValidatePermission
    frmPeople.Show
End Sub

' ==================== EXPORT GROUP ====================

Sub Export_File(control As IRibbonControl) ' ���� �������� (����)
    'ValidatePermission
    Call Export_Master
End Sub

' ==================== INFO GROUP ====================

Sub IRVersion(control As IRibbonControl) ' ���� Ȯ��
    Msg "���� ����: " & AppVersion & vbNewLine & "������: " & RelDate & vbNewLine & "������: " & ExpDate
End Sub

Sub IRSPO(control As IRibbonControl) ' SPO ����Ʈ ����
    Call OpenSPO
End Sub

Sub IRBugReport(control As IRibbonControl) ' ���� ����Ʈ
    Call OpenGoogleForm
End Sub

Sub IRManual(control As IRibbonControl) ' �Ŵ��� ����
    Call OpenManual
End Sub
