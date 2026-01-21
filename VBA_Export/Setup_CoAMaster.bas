Attribute VB_Name = "Setup_CoAMaster"
Option Explicit

'==============================================================================
' Module: Setup_CoAMaster
' Purpose: Populate CoAMaster table with HRE Chart of Accounts data
' Author: PwC BEP Team
' Created: 2026-01-21
' Version: 1.0
'
' Description:
'   This module reads CoA data from coa.md file and populates the Master table
'   in the CoAMaster worksheet with 178 account records. The data includes
'   account codes, descriptions, classifications, and derived metadata.
'
' Dependencies:
'   - mod_10_Public.bas (SpeedUp, SpeedDown, OpenProgress, CalculateProgress)
'   - CoAMaster worksheet with "Master" ListObject table
'   - coa.md file at /Users/jaewookim/Desktop/Project/HRE/참고/VBA_Export/coa.md
'
' Master Table Structure (10 columns):
'   1. Account - PwC 6-digit consolidated account code (PRIMARY KEY)
'   2. Description - English account description
'   3. 연결계정명 - Korean account name
'   4. 분류 - Korean classification (유동자산, 비유동부채, etc.)
'   5. Category - English classification (Current asset, etc.)
'   6. BSPL - Financial statement type (BS 자산, BS 부채, BS 지분, IS)
'   7. 대분류 - Major classification (자산, 부채, 자본, 수익, 비용)
'   8. Ranking - Sequential ranking (110, 120, 130...)
'   9. 부호 - Sign (D for debit accounts, C for credit accounts)
'   10. 금액 - Amount (initialized to 0)
'==============================================================================

Sub PopulateCoAMaster()
    '==========================================================================
    ' Main procedure to populate CoAMaster table from coa.md file
    '
    ' Process Flow:
    '   1. Validate prerequisites (worksheet and table existence)
    '   2. Read and parse coa.md file
    '   3. Transform data (derive BSPL, 대분류, 부호)
    '   4. Bulk insert data into Master table
    '   5. Format and protect worksheet
    '
    ' Error Handling:
    '   - Validates 6-digit Account codes
    '   - Skips rows with empty Account column
    '   - Reports processing errors to user
    '==========================================================================

    On Error GoTo ErrorHandler
    Call SpeedUp

    ' Variable declarations
    Dim tblMaster As ListObject
    Dim coaFilePath As String
    Dim coaData() As String
    Dim rowData() As String
    Dim masterArray() As Variant
    Dim i As Long, rowCount As Long
    Dim validRows As Long
    Dim fso As Object
    Dim fileStream As Object

    ' Display progress indicator
    Call OpenProgress("CoA Master 데이터 초기화 중...")
    Call CalculateProgress(0.1, "Master 테이블 검증 중...")

    '==========================================================================
    ' STEP 1: Validate CoAMaster worksheet and Master table
    '==========================================================================
    On Error Resume Next
    Set tblMaster = CoAMaster.ListObjects("Master")
    On Error GoTo ErrorHandler

    If tblMaster Is Nothing Then
        GoEnd "CoAMaster 시트의 'Master' 테이블을 찾을 수 없습니다!"
    End If

    '==========================================================================
    ' STEP 2: Read coa.md file
    '==========================================================================
    Call CalculateProgress(0.2, "coa.md 파일 읽는 중...")

    coaFilePath = "/Users/jaewookim/Desktop/Project/HRE/참고/VBA_Export/coa.md"

    ' Create FileSystemObject for UTF-8 file reading
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(coaFilePath) Then
        GoEnd "coa.md 파일을 찾을 수 없습니다: " & coaFilePath
    End If

    ' Read file using ADODB.Stream for UTF-8 support
    Set fileStream = CreateObject("ADODB.Stream")
    With fileStream
        .Type = 2 'adTypeText
        .Charset = "UTF-8"
        .Open
        .LoadFromFile coaFilePath

        Dim fileContent As String
        fileContent = .ReadText
        .Close
    End With

    ' Split content into lines
    coaData = Split(fileContent, vbLf)
    rowCount = UBound(coaData) + 1

    If rowCount < 2 Then
        GoEnd "coa.md 파일이 비어있거나 형식이 올바르지 않습니다!"
    End If

    '==========================================================================
    ' STEP 3: Parse data and build array (rows 2 to end, skip header row 1)
    '==========================================================================
    Call CalculateProgress(0.3, "데이터 파싱 중...")

    ' Allocate array for maximum possible rows (subtract header row)
    ReDim masterArray(1 To rowCount - 1, 1 To 10)
    validRows = 0

    For i = 1 To rowCount - 1 ' Skip row 0 (header)
        If i Mod 20 = 0 Then
            Call CalculateProgress(0.3 + (0.4 * i / rowCount), _
                "데이터 파싱 중... (" & i & "/" & (rowCount - 1) & ")")
        End If

        ' Parse tab-delimited row
        rowData = Split(coaData(i), vbTab)

        ' Skip rows without sufficient columns or empty Account code (column 6 = index 6)
        If UBound(rowData) < 7 Then GoTo NextRow

        Dim accountCode As String
        Dim accountDesc As String
        Dim coaName As String
        Dim classification As String
        Dim category As String

        ' Extract data from coa.md columns (0-based index)
        ' Column 7 (index 6) = Account, Column 8 (index 7) = Description
        ' Column 2 (index 1) = 연결계정명, Column 5 (index 4) = 분류, Column 6 (index 5) = Category
        accountCode = Trim(rowData(6))      ' Column 7: Account
        accountDesc = Trim(rowData(7))      ' Column 8: Description
        coaName = Trim(rowData(1))          ' Column 2: 연결계정명
        classification = Trim(rowData(4))   ' Column 5: 분류
        category = Trim(rowData(5))         ' Column 6: Category

        ' Skip if Account code is empty or invalid
        If accountCode = "" Or accountCode = "#N/A" Then GoTo NextRow

        ' Validate Account code is 6 digits
        If Len(accountCode) <> 6 Then
            ' Allow non-6-digit codes but log warning (some codes may be special)
            ' Skip completely invalid entries
            If Not IsNumeric(accountCode) Then GoTo NextRow
        End If

        ' Increment valid row counter
        validRows = validRows + 1

        '======================================================================
        ' Build Master table row with derived fields
        '======================================================================

        ' Column 1: Account (PwC 6-digit code)
        masterArray(validRows, 1) = accountCode

        ' Column 2: Description (English)
        masterArray(validRows, 2) = accountDesc

        ' Column 3: 연결계정명 (Korean account name)
        masterArray(validRows, 3) = coaName

        ' Column 4: 분류 (Korean classification)
        masterArray(validRows, 4) = classification

        ' Column 5: Category (English classification)
        masterArray(validRows, 5) = category

        ' Column 6: BSPL - Derive from Account code pattern
        ' 1xxxxx = BS 자산 (Balance Sheet Asset)
        ' 2xxxxx = BS 부채 (Balance Sheet Liability)
        ' 3xxxxx = BS 지분 (Balance Sheet Equity)
        ' 4xxxxx, 5xxxxx, 8xxxxx, 9xxxxx = IS (Income Statement)
        masterArray(validRows, 6) = DeriveBSPL(accountCode)

        ' Column 7: 대분류 - Derive from 분류 field
        ' 유동자산, 비유동자산 → 자산
        ' 유동부채, 비유동부채 → 부채
        ' 자본 → 자본
        ' 매출액, 매출원가 → 수익/비용
        ' 판관비, 영업외손익 → 비용
        masterArray(validRows, 7) = Derive대분류(classification)

        ' Column 8: Ranking - Sequential 100-unit increments (110, 120, 130...)
        masterArray(validRows, 8) = (validRows + 1) * 10

        ' Column 9: 부호 - D for Debit accounts (Assets, Expenses), C for Credit accounts (Liabilities, Equity, Revenue)
        ' Based on BSPL field
        masterArray(validRows, 9) = Derive부호(masterArray(validRows, 6), classification)

        ' Column 10: 금액 - Initialize to 0
        masterArray(validRows, 10) = 0

NextRow:
    Next i

    If validRows = 0 Then
        GoEnd "유효한 CoA 데이터가 없습니다!"
    End If

    '==========================================================================
    ' STEP 4: Clear existing Master table and insert new data
    '==========================================================================
    Call CalculateProgress(0.7, "Master 테이블 초기화 중...")

    CoAMaster.Unprotect PASSWORD

    ' Clear existing data
    If Not tblMaster.DataBodyRange Is Nothing Then
        tblMaster.DataBodyRange.Delete
    End If

    ' Resize table to accommodate new data
    tblMaster.Resize tblMaster.Range.Resize(validRows + 1)

    Call CalculateProgress(0.8, "데이터 삽입 중...")

    ' Bulk insert data using array (performance optimization)
    ReDim Preserve masterArray(1 To validRows, 1 To 10)
    tblMaster.DataBodyRange.Value = masterArray

    '==========================================================================
    ' STEP 5: Format table and protect worksheet
    '==========================================================================
    Call CalculateProgress(0.9, "테이블 서식 적용 중...")

    With tblMaster.DataBodyRange
        ' Apply borders
        .Borders(xlInsideHorizontal).LineStyle = xlDot
        .Borders(xlInsideVertical).LineStyle = xlDot

        ' Format Ranking column as number
        .Columns(8).NumberFormat = "0"

        ' Format 금액 column as accounting
        .Columns(10).NumberFormat = "#,##0"
    End With

    ' Protect worksheet with user filtering enabled
    CoAMaster.Protect PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True

    Call CalculateProgress(1, "완료!")

    ' Cleanup
    Set tblMaster = Nothing
    Set fso = Nothing
    Set fileStream = Nothing

    Call SpeedDown

    ' Success message
    Msg "CoA Master 테이블이 성공적으로 초기화되었습니다!" & vbNewLine & vbNewLine & _
        "총 " & validRows & "개의 계정이 등록되었습니다.", vbInformation

    Exit Sub

ErrorHandler:
    Call SpeedDown
    Call CalculateProgress(1) ' Close progress form

    Dim errMsg As String
    errMsg = "CoA Master 초기화 중 오류 발생:" & vbNewLine & vbNewLine & _
             "Error " & Err.Number & ": " & Err.Description

    ' Re-protect worksheet if error occurred
    On Error Resume Next
    CoAMaster.Protect PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True
    On Error GoTo 0

    GoEnd errMsg
End Sub

'==============================================================================
' Helper Functions for Derived Fields
'==============================================================================

Private Function DeriveBSPL(ByVal accountCode As String) As String
    '==========================================================================
    ' Derives BSPL classification from Account code pattern
    '
    ' Parameters:
    '   accountCode - 6-digit PwC account code
    '
    ' Returns:
    '   "BS 자산"  - Account starts with 1
    '   "BS 부채"  - Account starts with 2
    '   "BS 지분"  - Account starts with 3
    '   "IS"       - Account starts with 4, 5, 8, or 9
    '   ""         - Invalid or empty code
    '==========================================================================

    If Len(accountCode) < 1 Then
        DeriveBSPL = ""
        Exit Function
    End If

    Dim firstDigit As String
    firstDigit = Left(accountCode, 1)

    Select Case firstDigit
        Case "1"
            DeriveBSPL = "BS 자산"
        Case "2"
            DeriveBSPL = "BS 부채"
        Case "3"
            DeriveBSPL = "BS 지분"
        Case "4", "5", "8", "9"
            DeriveBSPL = "IS"
        Case Else
            DeriveBSPL = ""
    End Select
End Function

Private Function Derive대분류(ByVal classification As String) As String
    '==========================================================================
    ' Derives 대분류 (major classification) from 분류 field
    '
    ' Parameters:
    '   classification - Korean classification (분류)
    '
    ' Returns:
    '   "자산" - For 유동자산, 비유동자산
    '   "부채" - For 유동부채, 비유동부채
    '   "자본" - For 자본
    '   "수익" - For 매출액
    '   "비용" - For 매출원가, 판관비, 영업외손익
    '   ""     - For empty or unrecognized classification
    '==========================================================================

    classification = Trim(classification)

    If classification = "" Then
        Derive대분류 = ""
        Exit Function
    End If

    ' Asset classifications
    If InStr(1, classification, "유동자산", vbTextCompare) > 0 Or _
       InStr(1, classification, "비유동자산", vbTextCompare) > 0 Then
        Derive대분류 = "자산"
        Exit Function
    End If

    ' Liability classifications
    If InStr(1, classification, "유동부채", vbTextCompare) > 0 Or _
       InStr(1, classification, "비유동부채", vbTextCompare) > 0 Then
        Derive대분류 = "부채"
        Exit Function
    End If

    ' Equity classification
    If InStr(1, classification, "자본", vbTextCompare) > 0 Then
        Derive대분류 = "자본"
        Exit Function
    End If

    ' Revenue classification
    If InStr(1, classification, "매출액", vbTextCompare) > 0 Then
        Derive대분류 = "수익"
        Exit Function
    End If

    ' Expense classifications
    If InStr(1, classification, "매출원가", vbTextCompare) > 0 Or _
       InStr(1, classification, "판관비", vbTextCompare) > 0 Or _
       InStr(1, classification, "영업외손익", vbTextCompare) > 0 Then
        Derive대분류 = "비용"
        Exit Function
    End If

    ' Default for unrecognized classifications
    Derive대분류 = ""
End Function

Private Function Derive부호(ByVal bspl As String, ByVal classification As String) As String
    '==========================================================================
    ' Derives 부호 (account sign) based on BSPL and classification
    '
    ' Parameters:
    '   bspl - BSPL classification (BS 자산, BS 부채, BS 지분, IS)
    '   classification - Korean classification for IS accounts
    '
    ' Returns:
    '   "D" - Debit accounts (Assets, Expenses)
    '   "C" - Credit accounts (Liabilities, Equity, Revenue)
    '
    ' Logic:
    '   - BS 자산 → D (Assets have debit balance)
    '   - BS 부채 → C (Liabilities have credit balance)
    '   - BS 지분 → C (Equity has credit balance)
    '   - IS with 매출액 → C (Revenue has credit balance)
    '   - IS with 매출원가, 판관비, 영업외손익 → D (Expenses have debit balance)
    '   - IS other → D (Default for expenses)
    '==========================================================================

    Select Case bspl
        Case "BS 자산"
            Derive부호 = "D" ' Assets are debit accounts

        Case "BS 부채"
            Derive부호 = "C" ' Liabilities are credit accounts

        Case "BS 지분"
            Derive부호 = "C" ' Equity is credit account

        Case "IS"
            ' For Income Statement accounts, determine by classification
            If InStr(1, classification, "매출액", vbTextCompare) > 0 Then
                Derive부호 = "C" ' Revenue is credit
            Else
                Derive부호 = "D" ' Expenses (COGS, SG&A, Non-operating) are debit
            End If

        Case Else
            Derive부호 = "" ' Unknown or empty BSPL
    End Select
End Function

'==============================================================================
' End of Module: Setup_CoAMaster
'==============================================================================
