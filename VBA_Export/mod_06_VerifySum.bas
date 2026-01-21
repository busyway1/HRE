Attribute VB_Name = "mod_06_VerifySum"
Option Explicit
Sub RefreshPivotVerify()
    Dim pt As PivotTable

    On Error Resume Next

    If Check.Cells(12, 4).Value <> "Complete" Or Check.Cells(13, 4).Value <> "Complete" Or Check.Cells(14, 4).Value <> "Complete" Or _
       Check.Cells(16, 4).Value <> "Complete" Or Check.Cells(18, 4).Value <> "Complete" Then
        GoEnd "선행 단계를 완료하세요!"
    End If
    With Check.Cells(20, 4)
        .Value = "In Progress"
        .Interior.Color = RGB(255, 235, 156)
        .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
        .Offset(0, 2).Value = GetUserInfo()
    End With


    Call SpeedUp
    Call OpenProgress("검증 계산 중...")

    For Each pt In Verify.PivotTables
        pt.RefreshTable
    Next pt

    Call CalculateProgress(1, "계산 완료")
    Call SpeedDown
    Set pt = Nothing
End Sub
Sub VerifyBS()
    Dim pvt As PivotTable
    Dim dataRange As Range
    Dim lastRow As Long, verifyCol As Long, linkCol As Long
    Dim i As Long, j As Long
    Dim formula As String
    Dim linkTable As ListObject
    Dim corpCode As String

    On Error Resume Next
    Call SpeedUp

    Set pvt = Verify.PivotTables("합산검증(BS)")
    Set linkTable = HideSheet.ListObjects("Link")
    Set dataRange = pvt.TableRange1

    If pvt Is Nothing Then
        GoEnd "합산검증 피벗 테이블을 찾을 수 없습니다."
    End If

    Verify.Unprotect PASSWORD

    verifyCol = dataRange.Columns.count + dataRange.Column
    linkCol = verifyCol + 1

    ' 결과값 다시 초기화
    For i = dataRange.row + 2 To 1000
        Verify.Cells(i, verifyCol).Clear
        Verify.Cells(i, linkCol).Clear
    Next i

    With Range(Verify.Cells(dataRange.row, verifyCol), Verify.Cells(dataRange.row + 1, verifyCol))
        .Value = "검증"
        .Merge
        .Font.name = "맑은 고딕 Semilight"
        .Font.Size = 11
        .Font.Color = vbWhite
        .Interior.Color = RGB(192, 0, 0)
    End With
    With Range(Verify.Cells(dataRange.row, linkCol), Verify.Cells(dataRange.row + 1, linkCol))
        .Value = "링크"
        .Merge
        .Font.name = "맑은 고딕 Semilight"
        .Font.Size = 11
        .Font.Color = vbWhite
        .Interior.Color = RGB(192, 0, 0)
    End With

    If linkTable Is Nothing Then
        GoEnd "Hide 시트에서 Link 테이블을 찾을 수 없습니다."
    End If

    For i = dataRange.row + 2 To dataRange.row + dataRange.Rows.count - 1
        formula = "=IF((C" & i & "-D" & i & "-E" & i & ") = 0, ""TRUE"", C" & i & "-D" & i & "-E" & i & ")"

        With Verify.Cells(i, verifyCol)
            .formula = formula
            .Font.name = "맑은 고딕 Semilight"
            .Font.Size = 11
            .NumberFormat = "#,###;[Red](#,###);-"

            .FormatConditions.Delete

            .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""TRUE"""
            .FormatConditions(1).Interior.Color = RGB(198, 239, 206)

            .FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=""TRUE"""
            .FormatConditions(2).Interior.Color = RGB(255, 199, 206)
            .HorizontalAlignment = -4108
        End With

        corpCode = Verify.Cells(i, dataRange.Column).Value

        Dim foundCell As Range
        Set foundCell = linkTable.ListColumns("법인코드").DataBodyRange.Cells.Find(What:=corpCode, LookAt:=xlWhole, MatchCase:=False)
        If Not foundCell Is Nothing Then
            Dim linkColumnIndex As Long
            linkColumnIndex = linkTable.ListColumns("Link").Index
            Dim linkValue As String
            linkValue = foundCell.Offset(0, linkColumnIndex - 1).Value

            If linkValue <> "" Then
                Verify.Hyperlinks.Add Anchor:=Verify.Cells(i, linkCol), Address:=linkValue, TextToDisplay:="Link"
            Else
                Verify.Cells(i, linkCol).Value = ""
            End If
        Else
            Verify.Cells(i, linkCol).Value = ""
        End If

        With Verify.Cells(i, linkCol)
            .Font.name = "맑은 고딕 Semilight"
            .Font.Size = 11
            .HorizontalAlignment = -4108
            If .Value = "" Then
                .Interior.Color = RGB(255, 255, 0)
            End If
        End With
    Next i

    Dim borderRange As Range
    Set borderRange = Verify.Range(Verify.Cells(dataRange.row, verifyCol), Verify.Cells(dataRange.row + dataRange.Rows.count - 1, linkCol))

    ' 테두리를 스타일 적용
    With borderRange.Borders
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .Weight = xlThin
    End With

    Call SpeedDown
    Set pvt = Nothing: Set linkTable = Nothing: Set dataRange = Nothing: Set foundCell = Nothing: Set borderRange = Nothing
End Sub
Sub VerifyIS()
    Dim pvt As PivotTable
    Dim dataRange As Range
    Dim lastRow As Long, periodCol As Long, verifyCol As Long, linkCol As Long
    Dim i As Long
    Dim formula As String
    Dim linkTable As ListObject
    Dim corpTable As ListObject
    Dim corpCode As String
    Dim dateTable As ListObject

    On Error Resume Next
    Call SpeedUp

    Set pvt = Verify.PivotTables("합산검증(IS)")
    Set dataRange = pvt.TableRange1
    Set linkTable = HideSheet.ListObjects("Link")
    Set corpTable = CorpMaster.ListObjects("Corp")
    Set dateTable = HideSheet.ListObjects("결산연월")

    If pvt Is Nothing Then
        GoEnd "합산검증 피벗 테이블을 찾을 수 없습니다."
    End If

    If linkTable Is Nothing Then
        GoEnd "Hide 시트에서 Link 테이블을 찾을 수 없습니다."
    End If

    periodCol = dataRange.Columns.count + dataRange.Column
    verifyCol = dataRange.Columns.count + dataRange.Column + 1
    linkCol = verifyCol + 1

    For i = dataRange.row + 2 To 1000
        Verify.Cells(i, periodCol).Clear
        Verify.Cells(i, periodCol + 7).Clear
        Verify.Cells(i, verifyCol).Clear
        Verify.Cells(i, linkCol).Clear
    Next i

    With Range(Verify.Cells(dataRange.row, periodCol), Verify.Cells(dataRange.row + 1, periodCol))
        .Value = "기간"
        .Merge
        .Font.name = "맑은 고딕 Semilight"
        .Font.Size = 11
        .Font.Color = vbBlack
        .Interior.Color = RGB(217, 217, 217)
    End With

    ' 기간 처분
    With Verify.Cells(dataRange.row + 1, periodCol + 7)
        .Value = "기간"
        .Merge
        .Font.name = "맑은 고딕 Semilight"
        .Font.Size = 11
        .Font.Color = vbBlack
        .Interior.Color = RGB(217, 217, 217)
        .HorizontalAlignment = -4108
    End With

     With Range(Verify.Cells(dataRange.row, verifyCol), Verify.Cells(dataRange.row + 1, verifyCol))
        .Value = "검증"
        .Merge
        .Font.name = "맑은 고딕 Semilight"
        .Font.Size = 11
        .Font.Color = vbWhite
        .Interior.Color = RGB(192, 0, 0)
    End With
    With Range(Verify.Cells(dataRange.row, linkCol), Verify.Cells(dataRange.row + 1, linkCol))
        .Value = "링크"
        .Merge
        .Font.name = "맑은 고딕 Semilight"
        .Font.Size = 11
        .Font.Color = vbWhite
        .Interior.Color = RGB(192, 0, 0)
    End With


    For i = dataRange.row + 2 To dataRange.row + dataRange.Rows.count - 1
        corpCode = Verify.Cells(i, dataRange.Column).Value

        ' 기간 계산
        Dim periodCell As Range
        Dim cell As Range
        For Each cell In corpTable.ListColumns("법인코드").DataBodyRange
            If CStr(cell.Value) = corpCode Then
                Set periodCell = cell
                Exit For
            End If
        Next cell

        If Not periodCell Is Nothing Then
            Dim dateValue As String
            Dim dateValueDisposal As String
            Dim dateYearBegin As String
            dateValue = periodCell.Offset(0, 4).Value
            dateValueDisposal = periodCell.Offset(0, 5).Value
            dateYearBegin = dateTable.DataBodyRange.Cells(1, 1).Value & "-01-01"

            If dateValueDisposal = "-" Then
                ' 취득일 처분일(2000년 1월 1일로 처리)
                If dateValue = "-" Then
                    dateValue = "2000-01-01"
                End If

                If CDate(dateValue) < CDate(dateYearBegin) Then
                    Verify.Cells(i, periodCol).Value = dateTable.DataBodyRange.Cells(1, 1).Value & "-01 ~ " & dateTable.DataBodyRange.Cells(1, 1).Value & "-" & Format(dateTable.DataBodyRange.Cells(1, 2).Value, "00")
                    ' 기간 처분
                    Verify.Cells(i, periodCol + 7).Value = Verify.Cells(i, periodCol).Value
                Else
                    Verify.Cells(i, periodCol).Value = Format(dateValue, "yyyy-mm") & " ~ " & dateTable.DataBodyRange.Cells(1, 1).Value & "-" & Format(dateTable.DataBodyRange.Cells(1, 2).Value, "00")
                    ' 기간 처분
                    Verify.Cells(i, periodCol + 7).Value = Verify.Cells(i, periodCol).Value
                End If

            Else
                If CDate(dateValueDisposal) <= CDate(dateTable.DataBodyRange.Cells(1, 1).Value & "-" & Format(dateTable.DataBodyRange.Cells(1, 2).Value, "00")) Then
                    Verify.Cells(i, periodCol).Value = Format(dateYearBegin, "yyyy-mm") & " ~ " & Format(dateValueDisposal, "yyyy-mm")
                    ' 기간 처분
                    Verify.Cells(i, periodCol + 7).Value = Verify.Cells(i, periodCol).Value
                Else
                    Verify.Cells(i, periodCol).Value = Format(dateYearBegin, "yyyy-mm") & " ~ " & dateTable.DataBodyRange.Cells(1, 1).Value & "-" & Format(dateTable.DataBodyRange.Cells(1, 2).Value, "00")
                    ' 기간 처분
                    Verify.Cells(i, periodCol + 7).Value = Verify.Cells(i, periodCol).Value
                End If

            End If
        Else
            Verify.Cells(i, periodCol).Value = ""
            ' 기간 처분
            Verify.Cells(i, periodCol + 7).Value = ""
        End If

        With Verify.Cells(i, periodCol)
            .Font.name = "맑은 고딕 Semilight"
            .Font.Size = 11
            .HorizontalAlignment = -4108
            If .Value = "" Then
                .Interior.Color = RGB(255, 255, 0)
            End If
        End With

        ' 기간 처분
        With Verify.Cells(i, periodCol + 7)
            .Font.name = "맑은 고딕 Semilight"
            .Font.Size = 11
            .HorizontalAlignment = -4108
            If .Value = "" Then
                .Interior.Color = RGB(255, 255, 0)
            End If
        End With


        ' 검증 계산
        formula = "=IF((J" & i & "-K" & i & "-XLOOKUP(I" & i & ",[법인코드],[당기순이익],,0)) = 0, ""TRUE"", J" & i & "-K" & i & "-XLOOKUP(I" & i & ",[법인코드],[당기순이익],,0))"
        With Verify.Cells(i, verifyCol)
            .formula = formula
            .Font.name = "맑은 고딕 Semilight"
            .Font.Size = 11
            .NumberFormat = "#,###;[Red](#,###);-"

            .FormatConditions.Delete

            .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""TRUE"""
            .FormatConditions(1).Interior.Color = RGB(198, 239, 206)

            .FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=""TRUE"""
            .FormatConditions(2).Interior.Color = RGB(255, 199, 206)
            .HorizontalAlignment = -4108
        End With


        Dim foundCell As Range
        Set foundCell = linkTable.ListColumns("법인코드").DataBodyRange.Cells.Find(What:=corpCode, LookAt:=xlWhole, MatchCase:=False)
        If Not foundCell Is Nothing Then
            Dim linkColumnIndex As Long
            linkColumnIndex = linkTable.ListColumns("Link").Index
            Dim linkValue As String
            linkValue = foundCell.Offset(0, linkColumnIndex - 1).Value

            If linkValue <> "" Then
                Verify.Hyperlinks.Add Anchor:=Verify.Cells(i, linkCol), Address:=linkValue, TextToDisplay:="Link"
            Else
                Verify.Cells(i, linkCol).Value = ""
            End If
        Else
            Verify.Cells(i, linkCol).Value = ""
        End If

        With Verify.Cells(i, linkCol)
            .Font.name = "맑은 고딕 Semilight"
            .Font.Size = 11
            .HorizontalAlignment = -4108
            If .Value = "" Then
                .Interior.Color = RGB(255, 255, 0)
            End If
        End With
    Next i


    Dim borderRange As Range
    Set borderRange = Verify.Range(Verify.Cells(dataRange.row, periodCol), Verify.Cells(dataRange.row + dataRange.Rows.count - 1, linkCol))

    With borderRange.Borders
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .Weight = xlThin
    End With

    ' 처분 기간 Border Line 추가
    With Verify.Range(Verify.Cells(dataRange.row + 1, periodCol + 7), Verify.Cells(dataRange.row + dataRange.Rows.count - 1, periodCol + 7)).Borders
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .Weight = xlThin
    End With

    Call SpeedDown
    Set pvt = Nothing: Set dataRange = Nothing: Set linkTable = Nothing: Set corpTable = Nothing: Set dateTable = Nothing
    Set cell = Nothing: Set periodCell = Nothing: Set foundCell = Nothing: Set borderRange = Nothing
End Sub
Sub ValidateCorpCodes()
    Dim corpTable As ListObject
    Dim linkTable As ListObject
    Dim corpCode As Variant
    Dim lastColumn As Long
    Dim i As Long
    Dim outputCol As Long
    Dim missingBS As Boolean
    Dim missingIS As Boolean
    Dim corpCodeColumn As Range
    Dim isVerified As Boolean
    Dim scopeColumn As Range
    Dim corpName As Variant
    Dim scopeNum As Long ' 전체 Scope 법인 개수

    On Error Resume Next
    Call SpeedUp

    Set corpTable = CorpMaster.ListObjects("Corp")
    Set linkTable = HideSheet.ListObjects("Link")
    Set corpCodeColumn = corpTable.ListColumns("법인코드").DataBodyRange
    Set scopeColumn = corpTable.ListColumns("Scope").DataBodyRange
    isVerified = True

    If corpCodeColumn Is Nothing Then
        GoEnd "법인코드 열을 찾을 수 없습니다"
    End If
    If linkTable Is Nothing Then
        GoEnd "Hide 시트에서 Link 테이블을 찾을 수 없습니다."
    End If

    ' 기존 데이터 초기화 1000열까지
    lastColumn = 1000
    Verify.Range(Verify.Cells(14, 2), Verify.Cells(18, lastColumn)).UnMerge
    Verify.Range(Verify.Cells(14, 2), Verify.Cells(18, lastColumn)).Clear

    outputCol = 3
    scopeNum = 0

    For i = 1 To corpCodeColumn.Rows.count
        ' "Scope" 열이 "O"인 경우만 처리
        If scopeColumn.Cells(i, 1).Value = "O" Then
            corpCode = corpCodeColumn.Cells(i, 1).Value
            corpName = corpCodeColumn.Cells(i, 2).Value

            missingBS = WorksheetFunction.CountIf(Verify.Range("B:B"), corpCode) = 0
            missingIS = WorksheetFunction.CountIf(Verify.Range("I:I"), corpCode) = 0

            If missingBS Or missingIS Then

                Verify.Cells(14, outputCol).Value = corpCode
                Verify.Cells(14, outputCol).HorizontalAlignment = -4108
                Verify.Cells(14, outputCol).Font.name = "맑은 고딕 Semilight"
                Verify.Cells(15, outputCol).Value = corpName
                Verify.Cells(16, outputCol).Value = IIf(missingBS, "", "OK")
                Verify.Cells(17, outputCol).Value = IIf(missingIS, "", "OK")

                Dim foundCell As Range
                Set foundCell = linkTable.ListColumns("법인코드").DataBodyRange.Cells.Find(What:=corpCode, LookAt:=xlWhole, MatchCase:=False)
                If Not foundCell Is Nothing Then
                    Dim linkColumnIndex As Long
                    linkColumnIndex = linkTable.ListColumns("Link").Index
                    Dim linkValue As String
                    linkValue = foundCell.Offset(0, linkColumnIndex - 1).Value

                    If linkValue <> "" Then
                        Verify.Hyperlinks.Add Anchor:=Verify.Cells(18, outputCol), Address:=linkValue, TextToDisplay:="Link"
                    Else
                        Verify.Cells(18, outputCol).Value = ""
                    End If
                Else
                    Verify.Cells(18, outputCol).Value = ""
                End If

                With Range(Verify.Cells(15, outputCol), Verify.Cells(18, outputCol))
                    .HorizontalAlignment = -4108
                    .Font.name = "맑은 고딕 Semilight"
                    .Font.Size = 11
                End With
                outputCol = outputCol + 1
                isVerified = False
            End If

            scopeNum = scopeNum + 1
        End If
    Next i

    ' 법인 개수 확인
    Verify.Cells(4, 5).Value = "개수"
    Verify.Cells(5, 5).Value = "전체 Scope 법인 개수"
    Verify.Cells(6, 5).Value = "BS 법인 개수"
    Verify.Cells(7, 5).Value = "PL 법인 개수"
    Verify.Cells(4, 6).Value = ""

    With Verify.Range("E4:F4")
        .Interior.Color = RGB(217, 217, 217)
        .Font.name = "맑은 고딕 Semilight"
        .Font.Size = 11
    End With
    With Verify.Range("E5:F7")
        .Font.name = "맑은 고딕 Semilight"
        .Font.Size = 11
    End With
    With Verify.Range("F5:F7")
        .NumberFormat = "#,###;[Red](#,###);-"
    End With
    Verify.Range("E4:F7").Borders.LineStyle = xlContinuous

    ' 개수 계산
    With Verify.Cells(5, 6)
        .formula = "=COUNTIF(Corp[Scope],""O"")"
    End With

    With Verify.Cells(6, 6)
        Dim bsRange As String
        bsRange = Verify.PivotTables("합산검증(BS)").RowRange.Address
        .formula = "=COUNTA(" & bsRange & ")-1"
    End With

    With Verify.Cells(7, 6)
        Dim isRange As String
        isRange = Verify.PivotTables("합산검증(IS)").RowRange.Address
        .formula = "=COUNTA(" & isRange & ")-1"
    End With





    If isVerified Then
        With Verify.Cells(14, 2)
            .Value = "단순합산 검증 완료"
            .Font.name = "맑은 고딕 Semilight"
            .Font.Size = 11
            .Font.Bold = True
            .VerticalAlignment = -4108
            .HorizontalAlignment = -4108
            .Interior.Color = RGB(198, 239, 206)
        End With
    Else
        With Verify.Cells(14, 2)
            .Value = "미완료"
            .VerticalAlignment = -4108
            .HorizontalAlignment = -4108
            .Font.name = "맑은 고딕 Semilight"
            .Font.Size = 11
        End With
        Range(Verify.Cells(14, 2), Verify.Cells(15, 2)).Merge

        Verify.Cells(16, 2).Value = "BS"
        Verify.Cells(17, 2).Value = "PL"
        Verify.Cells(18, 2).Value = "Link"
        With Range(Verify.Cells(15, 2), Verify.Cells(18, 2))
            .Font.Size = 11
            .Font.name = "맑은 고딕 Semilight"
            .HorizontalAlignment = -4108
        End With

        lastColumn = Verify.Cells(15, Verify.Columns.count).End(xlToLeft).Column

        With Verify.Range("B15").CurrentRegion
            .Rows(1).Interior.Color = RGB(217, 217, 217)
            .Rows(2).Interior.Color = RGB(217, 217, 217)
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
        End With

        Dim cell As Range
        For Each cell In Verify.Range("B15").CurrentRegion
            If cell.Value = "" Then cell.Interior.Color = RGB(255, 255, 0)
        Next
    End If

    Verify.Columns("B:GG").ColumnWidth = 22
    Verify.Activate: Verify.Range("B1").Select

    If isVerified Then
        MsgBox "단순합산 자료집계가 완료되었습니다." & vbNewLine & "법인별 링크를 확인하세요.", vbInformation, AppName & " " & AppType
    Else
        MsgBox "단순합산 자료집계가 미완료되었습니다." & vbNewLine & "누락 법인을 자세히 확인하세요." & vbNewLine & _
                "법인별 링크를 확인하세요.", vbExclamation, AppName & " " & AppType
    End If

    Call SpeedDown
    Set corpTable = Nothing: Set linkTable = Nothing: Set corpCodeColumn = Nothing
    Set scopeColumn = Nothing: Set foundCell = Nothing: Set cell = Nothing
End Sub
Sub ValidateSheetColors()
    Dim cell As Range
    Dim validColors As Variant
    Dim invalidColorFound As Boolean

    On Error Resume Next
    Call SpeedUp

    validColors = Array( _
        RGB(0, 32, 96), _
        RGB(255, 255, 255), _
        RGB(217, 217, 217), _
        RGB(192, 0, 0), _
        RGB(198, 239, 206), _
        RGB(221, 243, 253))
    invalidColorFound = False

    For Each cell In Verify.UsedRange
        If Not IsValidColor(cell.DisplayFormat.Interior.Color, validColors) Then
            invalidColorFound = True
            Exit For
        End If
    Next cell

    Verify.Unprotect PASSWORD: ThisWorkbook.Unprotect PASSWORD_Workbook

    If invalidColorFound Then
        Verify.Tab.Color = RGB(255, 0, 0)
    Else
        Verify.Tab.Color = RGB(0, 255, 0)
    End If

    Verify.Protect PASSWORD, UserInterfaceOnly:=True: ThisWorkbook.Protect PASSWORD_Workbook

    If invalidColorFound Then
        MsgBox "오류가 발견되었습니다. 색상을 확인하세요.", vbExclamation, AppName & " " & AppType
    Else
        With Check.Cells(20, 4)
            .Value = "Complete"
            .Interior.Color = RGB(198, 239, 206)
            .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
            .Offset(0, 2).Value = GetUserInfo()
        End With
        MsgBox "검증이 완료되었습니다.", vbInformation, AppName & " " & AppType
    End If

    Call SpeedDown
    Set cell = Nothing
End Sub

Function IsValidColor(cellColor As Long, validColors As Variant) As Boolean
    Dim i As Long
    For i = LBound(validColors) To UBound(validColors)
        If cellColor = validColors(i) Then
            IsValidColor = True
            Exit Function
        End If
    Next i
    IsValidColor = False
End Function

