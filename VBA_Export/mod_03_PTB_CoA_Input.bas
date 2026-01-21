Attribute VB_Name = "mod_03_PTB_CoA_Input"
Option Explicit
' ============================================================================
' Module: mod_03_PTB_CoA_Input
' Project: HRE 연결마스터 (Consolidation Master)
' Version: 1.00
' Date: 2026-01-21
'
' Description: Enhanced CoA input with variant detection and 5-digit matching
' Key changes from BEP:
'  - Variant suffix detection (_내부거래, _IC)
'  - 5-digit base code matching (HRE uses 5-digit codes vs BEP's exact match)
'  - Multi-tier Dictionary lookup (variant-specific → BASE fallback)
'  - MC account exclusion retained for consistency
' ============================================================================

' ==================== PUBLIC FUNCTIONS ====================

' Fill_Input_Table - Populate CoA_Input table with First Drafting suggestions
' Enhanced with variant detection and 5-digit matching for HRE
Sub Fill_Input_Table()
   Dim tblPTB As ListObject
   Dim tblCoA As ListObject
   Dim tblRawCoA As ListObject
   Dim visibleRange As Range
   Dim newRows As Long
   Dim coaRow As ListRow
   Dim searchrange As Range
   Dim filteredData As Range
   Dim variantDict As Object
   Dim area As Range
   Dim cell As Range
   Dim accountCode As String
   Dim searchValue As String

   On Error Resume Next
   Call SpeedUp

   Set tblPTB = BSPL.ListObjects("PTB")
   Set tblCoA = AddCoA.ListObjects("CoA_Input")
   Set tblRawCoA = CorpCoA.ListObjects("Raw_CoA")

   BSPL.Unprotect PASSWORD: AddCoA.Unprotect PASSWORD

   With tblPTB.DataBodyRange
       Set visibleRange = .Resize(, 5).SpecialCells(xlCellTypeVisible)
   End With

   newRows = visibleRange.Cells.count / tblPTB.ListColumns.count

   If Not tblCoA.DataBodyRange Is Nothing Then
       tblCoA.DataBodyRange.Delete
   End If

   tblCoA.Resize tblCoA.Range.Resize(newRows + 1)
   visibleRange.Copy
   tblCoA.HeaderRowRange.Offset(1).PasteSpecial xlPasteValues
   DoEvents

   AddCoA.Activate
   If Not tblCoA.DataBodyRange Is Nothing Then
       With tblCoA.DataBodyRange
           .Borders(xlInsideHorizontal).LineStyle = xlDot
           .Borders(xlInsideVertical).LineStyle = xlDot
       End With
   End If
   Application.CutCopyMode = False

   ' ========== HRE ENHANCEMENT: Variant-Aware Dictionary ==========
   ' Build nested dictionary structure: baseCode -> variantType -> targetAccount
   With tblRawCoA.Range
        .AutoFilter Field:=1, Criteria1:="1000"  ' Filter for standard corp code

        Set variantDict = CreateObject("Scripting.Dictionary")
        Set filteredData = tblRawCoA.DataBodyRange.Rows.SpecialCells(xlCellTypeVisible)
        Set searchrange = Intersect(filteredData, tblRawCoA.ListColumns(2).Range)  ' Column 2: 계정코드

        ' Build variant-aware dictionary
        For Each area In searchrange.Areas
            For Each cell In area
                Dim baseCode As String
                Dim variantType As String
                Dim targetAccount As String

                accountCode = cell.Value  ' 계정코드 (e.g., "10300", "11401_내부거래")
                baseCode = GetBaseCode(accountCode)  ' Extract first 5 digits
                variantType = GetVariantType(accountCode)  ' Detect variant suffix
                targetAccount = cell.Offset(0, 3).Value  ' Column 5: Account (PwC consolidated code)

                ' Exclude MC accounts (consolidation accounts handled separately)
                If Left(targetAccount, 2) <> "MC" And Not IsEmpty(targetAccount) Then
                    ' Create nested dictionary if base code doesn't exist
                    If Not variantDict.Exists(baseCode) Then
                        Set variantDict(baseCode) = CreateObject("Scripting.Dictionary")
                    End If

                    ' Add variant-specific mapping
                    If Not variantDict(baseCode).Exists(variantType) Then
                        variantDict(baseCode).Add variantType, Array(targetAccount, cell.Offset(0, 4).Value)  ' Account + Description
                    End If
                End If
            Next cell
        Next area

        ' ========== HRE ENHANCEMENT: Multi-Tier Lookup ==========
        ' Apply First Drafting with variant-aware matching
        For Each coaRow In tblCoA.ListRows
            Dim ptbAccount As String
            Dim ptbBase As String
            Dim ptbVariant As String
            Dim suggestedAccount As String
            Dim suggestedDescription As String

            ptbAccount = coaRow.Range(1, 3).Value  ' 법인별 CoA (from PTB)
            ptbBase = GetBaseCode(ptbAccount)
            ptbVariant = GetVariantType(ptbAccount)

            ' Lookup strategy:
            ' 1. Try exact variant match (e.g., 11401_내부거래 → INTERCO_KR variant)
            ' 2. Fallback to BASE variant (e.g., 11401_내부거래 → 11401 BASE mapping)
            ' 3. If no match, leave empty for manual review

            If variantDict.Exists(ptbBase) Then
                If variantDict(ptbBase).Exists(ptbVariant) Then
                    ' Exact variant match found
                    suggestedAccount = variantDict(ptbBase)(ptbVariant)(0)
                    suggestedDescription = variantDict(ptbBase)(ptbVariant)(1)
                ElseIf variantDict(ptbBase).Exists("BASE") Then
                    ' Fallback to BASE variant
                    suggestedAccount = variantDict(ptbBase)("BASE")(0)
                    suggestedDescription = variantDict(ptbBase)("BASE")(1)
                Else
                    ' No match found
                    suggestedAccount = ""
                    suggestedDescription = ""
                End If
            Else
                ' Base code not found in dictionary
                suggestedAccount = ""
                suggestedDescription = ""
            End If

            ' Populate suggestion columns
            coaRow.Range(1, 4).Value = suggestedAccount  ' PwC_CoA column
            coaRow.Range(1, 5).Value = suggestedDescription  ' PwC_계정과목명 column
        Next coaRow
    End With

   tblRawCoA.AutoFilter.ShowAllData

   BSPL.Protect PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
   AddCoA.Cells.Locked = True
   AddCoA.Range("E5:G1048576").Locked = False
   AddCoA.Protect PASSWORD, UserInterfaceOnly:=True

   Call SpeedDown
   Set tblPTB = Nothing: Set tblCoA = Nothing: Set tblRawCoA = Nothing
   Set searchrange = Nothing: Set variantDict = Nothing: Set area = Nothing: Set cell = Nothing
   Set filteredData = Nothing
End Sub

' Fill_CoA_Table - Finalize CoA mappings and update Raw_CoA table
' (Unchanged from BEP - validation and bulk update logic remains same)
Sub Fill_CoA_Table()
    Dim tblAddCoA As ListObject
    Dim tblRawCoA As ListObject
    Dim tblPTB As ListObject
    Dim tblMaster As ListObject
    Dim inputRow As ListRow
    Dim hasEmptyMapping As Boolean
    Dim isDuplicate As Boolean
    Dim addedCount As Long
    Dim wsCheck As Worksheet
    Dim totalLogString As String

    ' Performance optimization: Array-based operations
    Dim masterData() As Variant
    Dim rawCoaData() As Variant
    Dim inputData() As Variant
    Dim ptbData() As Variant

    On Error Resume Next
    Call SpeedUp

    ' Load tables
    Set tblAddCoA = AddCoA.ListObjects("CoA_Input")
    Set tblRawCoA = CorpCoA.ListObjects("Raw_CoA")
    Set tblPTB = BSPL.ListObjects("PTB")
    Set tblMaster = CoAMaster.ListObjects("Master")

    ' Load data into arrays for performance
    If Not tblAddCoA.DataBodyRange Is Nothing Then
        inputData = tblAddCoA.DataBodyRange.Value
    Else
        GoEnd "입력된 데이터가 없습니다."
    End If

    If Not tblMaster.DataBodyRange Is Nothing Then
        masterData = tblMaster.DataBodyRange.Value
    End If

    If Not tblRawCoA.DataBodyRange Is Nothing Then
        rawCoaData = tblRawCoA.DataBodyRange.Value
    End If

    If Not tblPTB.DataBodyRange Is Nothing Then
        ptbData = tblPTB.DataBodyRange.Value
    End If

    AddCoA.Unprotect PASSWORD
    CorpCoA.Unprotect PASSWORD
    BSPL.Unprotect PASSWORD

    ' Clear previous highlighting
    If Not tblAddCoA.DataBodyRange Is Nothing Then
        tblAddCoA.DataBodyRange.Interior.Color = RGB(255, 255, 255)
    End If

    ' Build Master validation dictionary
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To UBound(masterData)
        dict.Add CStr(masterData(i, 1)) & "|" & CStr(masterData(i, 2)), i
    Next i

    ' Validate empty mappings and Master existence
    hasEmptyMapping = False
    For i = 1 To UBound(inputData)
        If IsEmpty(inputData(i, 3)) Or _
           Len(Trim(CStr(inputData(i, 3)))) = 0 Or _
           IsEmpty(inputData(i, 4)) Or _
           Len(Trim(CStr(inputData(i, 4)))) = 0 Then
            tblAddCoA.DataBodyRange.Cells(i, 3).Interior.Color = RGB(255, 255, 0)
            tblAddCoA.DataBodyRange.Cells(i, 4).Interior.Color = RGB(255, 255, 0)
            hasEmptyMapping = True
        End If

        ' Validate against Master table
        If Not dict.Exists(CStr(inputData(i, 4)) & "|" & CStr(inputData(i, 5))) Then
            tblAddCoA.DataBodyRange.Cells(i, 4).Interior.Color = RGB(255, 255, 0)
            tblAddCoA.DataBodyRange.Cells(i, 5).Interior.Color = RGB(255, 255, 0)
            hasEmptyMapping = True
        End If
    Next i

    If hasEmptyMapping Then
        AddCoA.Protect PASSWORD, UserInterfaceOnly:=True
        CorpCoA.Protect PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
        BSPL.Protect PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
        AddCoA.Activate
        GoEnd "PwC_CoA와 PwC_계정과목명 매칭되지 않은 항목이 있습니다."
    End If

    ' Build Raw_CoA duplicate check dictionary
    Dim dictRawCoA As Object
    Set dictRawCoA = CreateObject("Scripting.Dictionary")

    If Not IsEmpty(rawCoaData) Then
        For i = 1 To UBound(rawCoaData)
            dictRawCoA.Add CStr(rawCoaData(i, 1)) & "|" & CStr(rawCoaData(i, 2)), i
        Next i
    End If

    ' Prepare bulk insert data
    Dim newData() As Variant
    ReDim newData(1 To UBound(inputData), 1 To UBound(inputData, 2))
    Dim newDataCount As Long
    newDataCount = 0

    Call OpenProgress("CoA 추가 작업 진행중")

    ' Process each input row
    Dim key As String
    For i = 1 To UBound(inputData)
        Call CalculateProgress(i / UBound(inputData), "CoA 추가 중...")

        ' Update PTB highlighting
        Dim j As Long
        For j = 1 To UBound(ptbData)
            If ptbData(j, 1) = inputData(i, 1) And _
               ptbData(j, 2) = inputData(i, 2) Then
                tblPTB.DataBodyRange.Rows(j).Interior.Color = RGB(0, 176, 80)
                Exit For
            End If
        Next j

        ' Check for duplicates
        key = CStr(inputData(i, 1)) & "|" & CStr(inputData(i, 2))
        If Not dictRawCoA.Exists(key) Then
            newDataCount = newDataCount + 1
            For j = 1 To UBound(inputData, 2)
                newData(newDataCount, j) = inputData(i, j)
            Next j

            ' Build log string
            totalLogString = totalLogString & Join(Application.Index(Application.Transpose(Application.Index(inputData, i, 0)), 0), " | ") & vbNewLine
        End If
    Next i

    ' Bulk insert into Raw_CoA
    If newDataCount > 0 Then
        ReDim Preserve newData(1 To newDataCount, 1 To UBound(inputData, 2))
        With tblRawCoA.ListRows.Add(AlwaysInsert:=True)
            .Range.Resize(newDataCount, UBound(inputData, 2)).Value = newData
        End With
    End If

    Call CalculateProgress(1, "작업 완료")

    ' Log activity
    LogData CorpCoA.Name, "<CoA 대량추가>" & vbNewLine & vbNewLine & _
            "법인코드 | 법인별 CoA | 법인별 계정과목명 | PwC_CoA | PwC_계정명 | 비고" & vbNewLine & _
            totalLogString

    ' Update Check sheet
    With Check.Cells(19, 4)
        .Value = "If Any"
        .Interior.Color = RGB(237, 237, 237)
        .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
        .Offset(0, 2).Value = GetUserInfo()
    End With

    BSPL.Activate
    BSPL.Range("B1").Select

    Application.CutCopyMode = False
    AddCoA.Cells.Locked = True: AddCoA.Range("E5:G1048576").Locked = False
    ThisWorkbook.Unprotect PASSWORD_Workbook
    AddCoA.Visible = xlSheetVeryHidden
    ThisWorkbook.Protect PASSWORD_Workbook
    AddCoA.Protect PASSWORD, UserInterfaceOnly:=True
    CorpCoA.Protect PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    BSPL.Protect PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True

    Call SpeedDown
    MsgBox "CoA가 법인별CoA에 추가되었습니다.", vbInformation, AppName & " " & AppType
    MsgBox "CoA 확인 및 데이터 합산을 다시 실행" & vbNewLine & "하여 결과를 확인하세요.", vbInformation, AppName & " " & AppType

    Set dict = Nothing
    Set dictRawCoA = Nothing
    Set tblAddCoA = Nothing: Set tblRawCoA = Nothing: Set tblPTB = Nothing: Set tblMaster = Nothing
End Sub

' ==================== PRIVATE HELPER FUNCTIONS ====================

' GetBaseCode - Extract first 5 digits from account code (before variant suffix)
' Example: "11401_내부거래" → "11401"
' Example: "10300" → "10300"
Private Function GetBaseCode(accountCode As String) As String
    Dim baseCode As String
    baseCode = accountCode

    ' Remove variant suffixes
    If InStr(baseCode, "_") > 0 Then
        baseCode = Left(baseCode, InStr(baseCode, "_") - 1)
    End If

    ' Return first 5 digits (HRE standard)
    If Len(baseCode) >= 5 Then
        GetBaseCode = Left(baseCode, 5)
    Else
        GetBaseCode = baseCode
    End If
End Function

' GetVariantType - Detect variant type from account code suffix
' Returns: "BASE", "INTERCO_KR", "INTERCO_IC", or "CONSOLIDATION"
Private Function GetVariantType(accountCode As String) As String
    If InStr(accountCode, "_내부거래") > 0 Then
        GetVariantType = "INTERCO_KR"
    ElseIf InStr(accountCode, "_IC") > 0 Then
        GetVariantType = "INTERCO_IC"
    ElseIf Left(accountCode, 2) = "MC" Then
        GetVariantType = "CONSOLIDATION"
    Else
        GetVariantType = "BASE"
    End If
End Function
