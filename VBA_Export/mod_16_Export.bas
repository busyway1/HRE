Attribute VB_Name = "mod_16_Export"
' ============================================================================
' Module: mod_16_Export
' Project: HRE 연결마스터 (Consolidation Master)
' Migrated from: BEP v1.98
' Migration Date: 2026-01-21
'
' Description: Data export functionality
' Changes: Updated sheet names and export filename for HRE
' HRE Note: Export structure adapted for HRE worksheets
' ============================================================================
Option Explicit
Sub Export_Master()
    Dim ws As Worksheet
    Dim newWB As Workbook
    Dim sheetNames As Variant
    Dim i As Long
    Dim cell As Range
    Dim saveFilePath As String
    Dim userResponse As VbMsgBoxResult
    Dim linkArray As Variant
    On Error Resume Next

    If Evaluate(Check.Range("G4").Value) <> 1 Then
        Msg "모든 단계를 완료한 후에 내보낼 수 있습니다!", vbExclamation
        Exit Sub
    End If

    userResponse = MsgBox("연결마스터 작부에 파일을 내보내시겠습니까?", vbYesNo + vbQuestion, AppName & " " & AppType)
    If userResponse = vbNo Then
        Exit Sub
    End If

    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "저장할 위치를 선택하고 파일명을 입력하세요."
        .InitialFileName = "연결마스터" & VBA.Right(GetClosingYear(), 2) & GetClosingMonth() & "_작부에.xlsx"
        If .Show = -1 Then
            saveFilePath = .selectedItems(1)
        Else
            Exit Sub
        End If
    End With

    Call SpeedUp
    Application.DisplayAlerts = False
    sheetNames = Array("계정 마스터", "CoA 마스터", "법인별 CoA", "합계 BSPL", "검증", "취득, 처분 BSPL", "연결관리대장", "연결관리대장(처분)")

    Set newWB = Workbooks.Add

    Call OpenProgress("연결마스터 내보내기 중...")

    For i = LBound(sheetNames) To UBound(sheetNames)
        Call CalculateProgress((i + 1) / (UBound(sheetNames) + 1), sheetNames(i) & " 내보내기 중...")
        ThisWorkbook.Sheets(sheetNames(i)).Copy After:=newWB.Sheets(newWB.Sheets.count)
        If newWB.Sheets(newWB.Sheets.count).ProtectContents Then
            newWB.Sheets(newWB.Sheets.count).Unprotect PASSWORD
        End If
        With newWB.Sheets(newWB.Sheets.count)
            For Each cell In .UsedRange
                If Left(cell.formula, 1) = "=" Then
                    cell.Value = cell.Value
                End If
            Next cell
            .Range("A1").Select
        End With
    Next i

    Call OpenProgress(0.5, "링크 처리 중...")

    With newWB
        For i = .Queries.count To 1 Step -1
           .Queries(i).Delete
        Next i
       For i = .Connections.count To 1 Step -1
           .Connections(i).Delete
       Next i
        linkArray = .LinkSources(xlLinkTypeExcelLinks)
        If IsArray(linkArray) Then
            For i = LBound(linkArray) To UBound(linkArray)
                .BreakLink name:=linkArray(i), Type:=xlLinkTypeExcelLinks
            Next i
        End If

    End With

    Call CalculateProgress(1, "내보내기완료")
    newWB.Sheets(2).Select
    newWB.Sheets(1).Delete
    newWB.SaveAs fileName:=saveFilePath, FileFormat:=xlOpenXMLWorkbook
    newWB.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Call SpeedDown

    Msg "연결마스터 작부에 내보내기 완료되었습니다!", vbInformation

    Set ws = Nothing: Set newWB = Nothing: Set sheetNames = Nothing: Set cell = Nothing: Set linkArray = Nothing
End Sub
