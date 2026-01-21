VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' ============================================================================
' Module: ThisWorkbook (����_����_����)
' Project: HRE ���Ḷ���� (Consolidation Master)
' Version: 1.00
' Date: 2026-01-21
'
' Description: Workbook-level event handlers
' Changes from BEP:
'  - Removed: MC-related sheet protection (MCCoA, AddCoA_MC, etc.)
'  - Added: Exchange rate sheets protection (ȯ������)
'  - Updated: Comments and logging for HRE context
' ============================================================================

Private Const PASSWORD_WS As String = "BEP1234" ' ��ũ��Ʈ ��� PASSWORD
Private Const PASSWORD_WB As String = "PwCDA7529"

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    LogData_Access Me.Name, "����"
    Application.CommandBars("Queries and Connections").Enabled = True
End Sub

Private Sub Workbook_Open()
    On Error Resume Next
    LogData_Access Me.Name, "����"
    Worksheets("HideSheet").Range("N2").Value = AppVersion
'    If Not IsPermittedEmail() Then
'        Msg "�̿� ������ �����ϴ�!", vbCritical
'        Me.Close SaveChanges:=False
'    End If

    ' HRE - Core sheets protection
    Worksheets("CoAMaster").Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
    Worksheets("CorpCoA").Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
    Worksheets("BSPL").Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
    Worksheets("법인별 BSPL").Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
    Worksheets("CorpMaster").Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
    Worksheets("Verify").Protect PASSWORD_WS, UserInterfaceOnly:=True
    Worksheets("Check").Protect PASSWORD_WS, UserInterfaceOnly:=True
    Worksheets("ADBS").Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
    Worksheets("AddCoA_ADBS").Protect PASSWORD_WS, UserInterfaceOnly:=True
    Worksheets("AddCoA").Protect PASSWORD_WS, UserInterfaceOnly:=True

    ' HRE - Optional: Protect exchange rate sheets if they exist
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In Me.Worksheets
        If ws.Name = "ȯ������(���)" Or ws.Name = "ȯ������(����)" Then
            ws.Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
        End If
    Next ws
    On Error GoTo 0

    ProtectQueryEditor
End Sub

Private Sub Workbook_SheetBeforeDelete(ByVal Sh As Object)
    If Not CheckPassword("��Ʈ�� �����Ͻðڽ��ϱ�?") Then
        Msg "��Ʈ ������ ��ҵǾ����ϴ�.", vbInformation
        Application.EnableEvents = False
        Application.Undo
        Application.EnableEvents = True
    End If
End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)
    If Not CheckPassword("�� ��Ʈ�� �߰��Ͻðڽ��ϱ�?") Then
        Msg "�� ��Ʈ �߰��� ��ҵǾ����ϴ�.", vbInformation
        Application.EnableEvents = False
        Me.Sheets(Me.Sheets.Count).Delete
        Application.EnableEvents = True
    End If
End Sub

Private Function CheckPassword(PromptMessage As String) As Boolean
    Dim UserInput As String
    UserInput = InputBox(PromptMessage & vbNewLine & vbNewLine & "�����Ϸ��� ��й�ȣ�� �Է��ϼ���:", "��й�ȣ Ȯ��")

    If UserInput = PASSWORD_WB Then
        CheckPassword = True
    Else
        CheckPassword = False
    End If
End Function
