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
' Module: ThisWorkbook (현재_통합_문서)
' Project: HRE 연결마스터 (Consolidation Master)
' Version: 1.00
' Date: 2026-01-21
'
' Description: Workbook-level event handlers
' Changes from BEP:
'  - Removed: MC-related sheet protection (MCCoA, AddCoA_MC, etc.)
'  - Added: Exchange rate sheets protection (환율정보)
'  - Updated: Comments and logging for HRE context
' ============================================================================

Private Const PASSWORD_WS As String = "BEP1234" ' 워크시트 잠금 PASSWORD
Private Const PASSWORD_WB As String = "PwCDA7529"

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    LogData_Access Me.Name, "종료"
    Application.CommandBars("Queries and Connections").Enabled = True
End Sub

Private Sub Workbook_Open()
    On Error Resume Next
    LogData_Access Me.Name, "실행"
    Worksheets("HideSheet").Range("N2").Value = AppVersion
'    If Not IsPermittedEmail() Then
'        Msg "이용 권한이 없습니다!", vbCritical
'        Me.Close SaveChanges:=False
'    End If

    ' HRE - Core sheets protection (removed MC-related sheets)
    Worksheets("CoAMaster").Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
    Worksheets("CorpCoA").Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
    Worksheets("BSPL").Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
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
        If ws.Name = "환율정보(평균)" Or ws.Name = "환율정보(일자)" Then
            ws.Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
        End If
    Next ws
    On Error GoTo 0

    ProtectQueryEditor
End Sub

Private Sub Workbook_SheetBeforeDelete(ByVal Sh As Object)
    If Not CheckPassword("시트를 삭제하시겠습니까?") Then
        Msg "시트 삭제가 취소되었습니다.", vbInformation
        Application.EnableEvents = False
        Application.Undo
        Application.EnableEvents = True
    End If
End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)
    If Not CheckPassword("새 시트를 추가하시겠습니까?") Then
        Msg "새 시트 추가가 취소되었습니다.", vbInformation
        Application.EnableEvents = False
        Me.Sheets(Me.Sheets.Count).Delete
        Application.EnableEvents = True
    End If
End Sub

Private Function CheckPassword(PromptMessage As String) As Boolean
    Dim UserInput As String
    UserInput = InputBox(PromptMessage & vbNewLine & vbNewLine & "진행하려면 비밀번호를 입력하세요:", "비밀번호 확인")

    If UserInput = PASSWORD_WB Then
        CheckPassword = True
    Else
        CheckPassword = False
    End If
End Function
