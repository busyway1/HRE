Attribute VB_Name = "mod_17_ExchangeRate"
Option Explicit
' ============================================================================
' Module: mod_17_ExchangeRate
' Project: HRE ���Ḷ���� (Consolidation Master)
' Version: 1.00
' Date: 2026-01-21
'
' Description: Exchange rate fetching from KEB Hana Bank API
' Fetches both average exchange rates (for P&L) and spot rates (for B/S)
' Handles special currencies (JPY, VND, IDR with ȯ��=100)
' Adds KRW baseline automatically
'
' SECURITY NOTE: HTML parsing uses innerHTML on a trusted source
' (KEB Hana Bank official API). The htmlfile COM object is isolated from
' browser security contexts. For production, consider additional validation.
' ============================================================================

' ==================== PUBLIC FUNCTIONS ====================

' GetER_Flow - Fetch average exchange rates for a period (P&L items)
' Called from ribbon button "���ȯ�� ��ȸ"
Sub GetER_Flow(Optional control As Variant = Nothing)
    Dim StartDate As Date, EndDate As Date, newSheetName As String, ws As Worksheet
    Dim html As Object, lastRow As Long, i As Long
    Dim splitValues As Variant
    Dim sheet As Worksheet

    On Error Resume Next

    ' ���� ��¥�� ���� ��¥ ���� (���� ��¥������ ���� ����)
    frmCalendar.Caption = "�������� �����ϼ���."
    StartDate = frmCalendar.GetDate(xlNextToCursor, 3)

    If StartDate = 0 Or StartDate > Date Then
        MsgBox "��ȿ���� ���� ���� ��¥�Դϴ�. ���� �Ǵ� ���� ��¥�� �������ּ���.", vbExclamation
        Exit Sub
    End If

    frmCalendar.Caption = "�������� �����ϼ���."
    EndDate = frmCalendar.GetDate(xlNextToCursor, 3)
    If EndDate = 0 Or EndDate > Date Or EndDate < StartDate Then
        MsgBox "��ȿ���� ���� ���� ��¥�Դϴ�. ���� ��¥ ���ĺ��� ���ñ����� ��¥�� �������ּ���.", vbExclamation
        Exit Sub
    End If

    ' Clear existing sheets with same name
    Application.DisplayAlerts = False
    For Each sheet In ThisWorkbook.Sheets
        If Left(sheet.Name, 9) = "ȯ������(���)" Then
            sheet.Cells.Clear
        End If
    Next sheet
    Application.DisplayAlerts = True

    ' ��Ʈ �غ�
    newSheetName = "ȯ������(���)"
    Set ws = PrepareSheet(newSheetName)

    ws.Range("A1").Value = "��ȸ �Ⱓ : " & Format(StartDate, "yyyy-mm-dd") & " ~ " & Format(EndDate, "yyyy-mm-dd")

    ' 1�� 1�� ������ ó��
    If Format(StartDate, "mm-dd") = "01-01" Then
        StartDate = DateSerial(Year(StartDate), 1, 2)
        ws.Range("A1").Value = "��ȸ �Ⱓ : " & Format(StartDate, "yyyy-mm-dd") & " ~ " & Format(EndDate, "yyyy-mm-dd") & " (1�� 1���� �������̹Ƿ� 1�� 2�Ϻ��ͷ� ��ȸ)"
    End If

    ' HTML ������ �������� �� ó�� (from trusted KEB Hana Bank API)
    Set html = CreateObject("htmlfile")
    Dim htmlResponse As String
    htmlResponse = GetHtmlFlow("https://www.kebhana.com/cms/rate/wpfxd651_06i_01.do", StartDate, EndDate)
    html.body.innerHTML = htmlResponse

    PutIntoClipboard html.getElementsByClassName("tblBasic")(0).outerHTML
    ws.Range("A2").Value = "�� ��ȸ���� ��/��/������ �Ǵ� ���࿵���� 1ȸ�� ���� ���� ���, �� �������ڷ� ��ȸ�˴ϴ�."
    ws.Range("A4").PasteSpecial

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ' �����۸�ũ ����
    RemoveHyperlinks ws

    ' B���� C���� ��ȭ �ڵ�� ȯ�� ���� �߰�
    ws.Columns("B:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Application.DisplayAlerts = False
    ws.Range("A5").Value = "������ �� ��ȭ"
    ws.Range("B5:B7").Value = "��ȭ"
    ws.Range("B5:B7").Merge
    ws.Range("C5:C7").Value = "ȯ��"
    ws.Range("C5:C7").Merge
    Application.DisplayAlerts = True

    For i = 8 To lastRow
        splitValues = Split(ws.Cells(i, 1).Value, " ")
        If UBound(splitValues) >= 1 Then
            ws.Cells(i, 2).Value = splitValues(1)
        Else
            ws.Cells(i, 2).Value = ""
        End If

        ' Special handling for JPY, VND, IDR (ȯ��=100)
        If splitValues(1) = "JPY" Or splitValues(1) = "VND" Or splitValues(1) = "IDR" Then
            ws.Cells(i, 3).Value = 100
        Else
            ws.Cells(i, 3).Value = 1
        End If
    Next i

    ' ���ѹα� KRW ���� �߰� (baseline)
    lastRow = lastRow + 1
    ws.Cells(lastRow, 1).Value = "���ѹα� KRW"
    ws.Cells(lastRow, 2).Value = "KRW"
    ws.Cells(lastRow, 3).Value = 1
    ws.Cells(lastRow, 11).Value = 1

    ' ���� ����
    ApplyFormatting ws, lastRow

    ' ���ݼ� ����
    RemoveGridlines ws

    Application.CutCopyMode = False
    ws.Columns("D:M").AutoFit
    ws.Columns("A:A").ColumnWidth = 12
    ws.Columns("B:B").ColumnWidth = 8
    ws.Columns("C:C").ColumnWidth = 8
    ws.Columns("B:C").HorizontalAlignment = xlHAlignCenter
    ws.Range("A1").Select

    ' Update Check sheet workflow status (Row 20)
    Call UpdateCheckStatus(20, "Complete")

    MsgBox "���ȯ�� ������ ������Ʈ�Ǿ����ϴ�.", vbInformation

    Set html = Nothing
    Set ws = Nothing

End Sub

' GetER_Spot - Fetch spot exchange rates for a specific date (B/S items)
' Called from ribbon button "�⸻ȯ�� ��ȸ"
Sub GetER_Spot(Optional control As Variant = Nothing)
    Dim selectedDate As Date, newSheetName As String, ws As Worksheet
    Dim html As Object, lastRow As Long, i As Long
    Dim splitValues As Variant
    Dim sheet As Worksheet

    On Error Resume Next

    ' ��¥ ���� (���� ��¥������ ���� ����)
    frmCalendar.Caption = "�������� �����ϼ���."
    selectedDate = frmCalendar.GetDate(xlNextToCursor, 3)
    If selectedDate = 0 Or selectedDate > Date Then
        MsgBox "��ȿ���� ���� ��¥�Դϴ�. ���� �Ǵ� ���� ��¥�� �������ּ���.", vbExclamation
        Exit Sub
    End If

    ' Clear existing sheets with same name
    Application.DisplayAlerts = False
    For Each sheet In ThisWorkbook.Sheets
        If Left(sheet.Name, 9) = "ȯ������(����)" Then
            sheet.Cells.Clear
        End If
    Next sheet
    Application.DisplayAlerts = True

    ' ��Ʈ �غ�
    newSheetName = "ȯ������(����)"
    Set ws = PrepareSheet(newSheetName)

    ws.Range("A1").Value = "��ȸ ������ : " & Format(selectedDate, "yyyy-mm-dd")

    ' HTML ������ �������� �� ó�� (from trusted KEB Hana Bank API)
    Set html = CreateObject("htmlfile")
    Dim htmlResponse As String
    htmlResponse = GetHtmlSpot("https://www.kebhana.com/cms/rate/wpfxd651_01i_01.do", selectedDate)
    html.body.innerHTML = htmlResponse

    PutIntoClipboard html.getElementsByClassName("tblBasic")(0).outerHTML
    ws.Range("A2").Value = "�� ��ȸ���� ��/��/������ �Ǵ� ���࿵���� 1ȸ�� ���� ���� ���, �� �������ڷ� ��ȸ�˴ϴ�."
    ws.Range("A4").PasteSpecial

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ' �����۸�ũ ����
    RemoveHyperlinks ws

    ' B���� C���� ��ȭ �ڵ�� ȯ�� ���� �߰�
    ws.Columns("B:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Application.DisplayAlerts = False
    ws.Range("A5").Value = "������ �� ��ȭ"
    ws.Range("B5:B7").Value = "��ȭ"
    ws.Range("B5:B7").Merge
    ws.Range("C5:C7").Value = "ȯ��"
    ws.Range("C5:C7").Merge
    Application.DisplayAlerts = True

    For i = 8 To lastRow
        splitValues = Split(ws.Cells(i, 1).Value, " ")
        If UBound(splitValues) >= 1 Then
            ws.Cells(i, 2).Value = splitValues(1)
        Else
            ws.Cells(i, 2).Value = ""
        End If

        ' For spot rates, ȯ�� value is in parentheses (e.g., "USD (1)")
        If UBound(splitValues) >= 3 Then
            ws.Cells(i, 3).Value = Replace(Replace(splitValues(2), "(", ""), ")", "")
        Else
            ws.Cells(i, 3).Value = 1
        End If
    Next i

    ' ���ѹα� KRW ���� �߰� (baseline)
    lastRow = lastRow + 1
    ws.Cells(lastRow, 1).Value = "���ѹα� KRW"
    ws.Cells(lastRow, 2).Value = "KRW"
    ws.Cells(lastRow, 3).Value = 1
    ws.Cells(lastRow, 11).Value = 1

    ' ���� ����
    ApplyFormatting ws, lastRow

    ' ���ݼ� ����
    RemoveGridlines ws

    Application.CutCopyMode = False
    ws.Columns("D:M").AutoFit
    ws.Columns("A:A").ColumnWidth = 12
    ws.Columns("B:B").ColumnWidth = 8
    ws.Columns("C:C").ColumnWidth = 8
    ws.Columns("B:C").HorizontalAlignment = xlHAlignCenter
    ws.Range("A1").Select

    ' Update Check sheet workflow status (Row 20)
    Call UpdateCheckStatus(20, "Complete")

    MsgBox "�⸻ȯ�� ������ ������Ʈ�Ǿ����ϴ�.", vbInformation

    Set html = Nothing
    Set ws = Nothing

End Sub

' ==================== PRIVATE HELPER FUNCTIONS ====================

' PrepareSheet - Create or clear sheet for exchange rate data
Private Function PrepareSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        ' Create new sheet after HideSheet
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("HideSheet"))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    Set PrepareSheet = ws
End Function

' RemoveGridlines - Hide gridlines for cleaner appearance
Private Sub RemoveGridlines(ws As Worksheet)
    Application.ScreenUpdating = False
    ws.Select
    ActiveWindow.DisplayGridlines = False
    Application.ScreenUpdating = True
End Sub

' RemoveHyperlinks - Remove all hyperlinks from worksheet
Private Sub RemoveHyperlinks(ws As Worksheet)
    Dim hyp As Hyperlink
    For Each hyp In ws.Hyperlinks
        hyp.Delete
    Next hyp
End Sub

' ApplyFormatting - Apply borders, fonts, and row heights
Private Sub ApplyFormatting(ws As Worksheet, lastRow As Long)
    Dim formatRange As Range
    Set formatRange = ws.Range("A5:M" & lastRow)

    ' �׵θ� �߰�
    With formatRange.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    formatRange.BorderAround Weight:=xlMedium

    ' �� ���� ����
    ws.Range("A5:M" & lastRow).RowHeight = 12

    ' �۲� ����
    With formatRange.Font
        .Name = "���� ����"
        .Size = 10
    End With

    ' ���� �� ���� ����
    With ws.Range("A4:M4")
        .Font.Bold = True
        .RowHeight = 15
    End With

    ' ù ��°�� �� ��° �� ���� ����
    ws.Range("A1:M2").RowHeight = 15

    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
    End With
End Sub

' GetHtmlFlow - Fetch HTML from KEB Hana Bank for average exchange rates
Private Function GetHtmlFlow(url As String, StartDate As Date, EndDate As Date) As String
    Dim http As Object, postData As String

    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    postData = "ajax=true" & _
               "&curCd=" & _
               "&pbldDvCd=1" & _
               "&pbldSqn=" & _
               "&hid_key_data=" & _
               "&inqKindCd=1" & _
               "&hid_enc_data=" & _
               "&requestTarget=searchContentDiv" & _
               "&tmpInqStrDt=" & Format(StartDate, "YYYY-MM-DD") & _
               "&inqStrDt=" & Format(StartDate, "YYYYMMDD") & _
               "&tmpInqEndDt=" & Format(EndDate, "YYYY-MM-DD") & _
               "&inqEndDt=" & Format(EndDate, "YYYYMMDD")

    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        .setRequestHeader "Referer", "https://www.kebhana.com/cms/rate/index.do?contentUrl=/cms/rate/wpfxd651_06i_01.do"
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        .setRequestHeader "X-Requested-With", "XMLHttpRequest"
        .send postData
        GetHtmlFlow = .responseText
    End With

    Set http = Nothing
End Function

' GetHtmlSpot - Fetch HTML from KEB Hana Bank for spot exchange rates
Private Function GetHtmlSpot(url As String, selectedDate As Date) As String
    Dim http As Object, postData As String

    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    postData = "ajax=true" & _
               "&curCd=" & _
               "&pbldDvCd=1" & _
               "&pbldSqn=" & _
               "&hid_key_data=" & _
               "&inqKindCd=1" & _
               "&hid_enc_data=" & _
               "&requestTarget=searchContentDiv" & _
               "&tmpInqStrDt=" & Format(selectedDate, "YYYY-MM-DD") & _
               "&inqStrDt=" & Format(selectedDate, "YYYYMMDD")

    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        .setRequestHeader "Referer", "https://www.kebhana.com/cms/rate/index.do?contentUrl=/cms/rate/wpfxd651_01i.do"
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        .setRequestHeader "X-Requested-With", "XMLHttpRequest"
        .send postData
        GetHtmlSpot = .responseText
    End With

    Set http = Nothing
End Function

' PutIntoClipboard - Copy text to clipboard
Private Function PutIntoClipboard(txt As String)
    Dim dataObj As Object
    Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")

    dataObj.SetText txt
    dataObj.PutInClipboard

    Set dataObj = Nothing
End Function

' UpdateCheckStatus - Update workflow status in Check sheet
Private Sub UpdateCheckStatus(stepRow As Long, status As String)
    On Error Resume Next
    With Check
        .Cells(stepRow, 4).Value = status  ' "Complete", "In Progress", or ""
        .Cells(stepRow, 5).Value = Format(Now, "yyyy-mm-dd hh:mm")
        .Cells(stepRow, 6).Value = GetUserInfo()

        ' Color coding
        Select Case status
            Case "Complete"
                .Cells(stepRow, 4).Interior.Color = RGB(198, 239, 206)  ' Light green
            Case "In Progress"
                .Cells(stepRow, 4).Interior.Color = RGB(255, 235, 156)  ' Yellow
            Case Else
                .Cells(stepRow, 4).Interior.ColorIndex = xlNone
        End Select
    End With
End Sub
