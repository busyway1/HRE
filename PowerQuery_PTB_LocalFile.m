/*
  Power Query Template: PTB from Local Excel File (Testing/Offline Mode)
  Version: 2.0
  Date: 2026-01-21

  Purpose:
  - Load PTB data from local Excel file for testing/offline development
  - Alternative to SharePoint connection when SPO unavailable
  - Same transformation logic as PowerQuery_PTB.m
  - Support for .xls and .xlsx formats

  Source File:
  - 12월 말 합계잔액시산표_에이치알이_수정완료(01.13).xls
  - Location: To be specified by user

  Installation Instructions:
  1. Open Power Query Editor (데이터 > 쿼리 및 연결 > 쿼리 편집)
  2. Create new blank query (새 원본 > 기타 원본 > 빈 쿼리)
  3. Open Advanced Editor (고급 편집기)
  4. Paste this M code
  5. Replace [FILE_PATH] with actual file path
  6. Replace [SHEET_NAME] with actual sheet name
  7. Adjust [HEADER_ROW] if needed (default: 1)
  8. Rename query to "PTB_Local"
  9. Close & Load to Table

  Migration Path:
  - Use this for initial setup and testing
  - Switch to PowerQuery_PTB.m when SharePoint connection is ready
  - Table structure identical to ensure VBA code compatibility
*/

let
    // ========== CONNECTION PARAMETERS ==========
    // TODO: Replace with actual file path
    FilePath = "[FILE_PATH]",
    // Example: "C:\Users\Username\Desktop\HRE\12월 말 합계잔액시산표_에이치알이_수정완료(01.13).xls"
    // Example Mac: "/Users/username/Desktop/HRE/12월 말 합계잔액시산표_에이치알이_수정완료(01.13).xls"

    SheetName = "[SHEET_NAME]",  // Example: "Sheet1" or "합계잔액시산표"
    HeaderRow = 1,  // Row number where headers are located

    // ========== STEP 1: Connect to Local Excel File ==========
    // Supports both .xls (Excel 97-2003) and .xlsx (Excel 2007+)
    Source = Excel.Workbook(File.Contents(FilePath), null, true),

    // ========== STEP 2: Select Target Sheet ==========
    SelectedSheet = Source{[Item=SheetName,Kind="Sheet"]}[Data],

    // ========== STEP 3: Promote Headers ==========
    PromotedHeaders = Table.PromoteHeaders(SelectedSheet, [PromoteAllScalars=true]),

    // ========== STEP 4: Select Required Columns ==========
    // Adjust column names based on actual Excel file structure
    // Use exact column names from the Excel file
    SelectedColumns = Table.SelectColumns(PromotedHeaders, {
        "법인코드",           // Corporate Code
        "법인명",             // Corporate Name
        "계정코드",           // Account Code
        "계정과목명",         // Account Name
        "차변",               // Debit Amount
        "대변",               // Credit Amount
        "잔액",               // Balance (Net)
        "통화코드",           // Currency Code
        "결산연월",           // Closing Period
        "비고"                // Remarks
    }),

    // ========== STEP 5: Data Type Enforcement ==========
    TypedColumns = Table.TransformColumnTypes(SelectedColumns, {
        {"법인코드", type text},
        {"법인명", type text},
        {"계정코드", type text},
        {"계정과목명", type text},
        {"차변", Currency.Type},
        {"대변", Currency.Type},
        {"잔액", Currency.Type},
        {"통화코드", type text},
        {"결산연월", type text},
        {"비고", type text}
    }),

    // ========== STEP 6: Clean Data ==========
    // Remove rows where account code is null or empty
    CleanedRows = Table.SelectRows(TypedColumns,
        each [계정코드] <> null and [계정코드] <> "" and [계정코드] <> "계정코드"
    ),

    // ========== STEP 7: Trim Whitespace ==========
    TrimmedText = Table.TransformColumns(CleanedRows, {
        {"법인코드", Text.Trim, type text},
        {"법인명", Text.Trim, type text},
        {"계정코드", Text.Trim, type text},
        {"계정과목명", Text.Trim, type text},
        {"통화코드", Text.Trim, type text},
        {"결산연월", Text.Trim, type text}
    }),

    // ========== STEP 8: Replace Null Amounts with Zero ==========
    ReplacedNulls = Table.ReplaceValue(TrimmedText, null, 0,
        Replacer.ReplaceValue, {"차변", "대변", "잔액"}
    ),

    // ========== STEP 9: Add Composite Key ==========
    AddedKey = Table.AddColumn(ReplacedNulls, "복합키",
        each [법인코드] & "_" & [계정코드],
        type text
    ),

    // ========== STEP 10: Add PwC CoA Placeholder Columns ==========
    // These columns will be populated through CoA mapping workflow
    AddedPwCCoA = Table.AddColumn(AddedKey, "PwC_CoA",
        each null,
        type text
    ),

    AddedPwCCoAName = Table.AddColumn(AddedPwCCoA, "PwC_CoA_Name",
        each null,
        type text
    ),

    AddedPwCLevel1 = Table.AddColumn(AddedPwCCoAName, "PwC_Level1",
        each null,
        type text
    ),

    AddedPwCLevel2 = Table.AddColumn(AddedPwCLevel1, "PwC_Level2",
        each null,
        type text
    ),

    AddedPwCLevel3 = Table.AddColumn(AddedPwCLevel2, "PwC_Level3",
        each null,
        type text
    ),

    // ========== STEP 11: Add Mapping Status Column ==========
    AddedMappingStatus = Table.AddColumn(AddedPwCLevel3, "매핑상태",
        each if [PwC_CoA] = null then "미완료" else "완료",
        type text
    ),

    // ========== STEP 12: Add Verification Columns ==========
    AddedVerificationFlag = Table.AddColumn(AddedMappingStatus, "검증플래그",
        each null,
        type text
    ),

    AddedVerificationNote = Table.AddColumn(AddedVerificationFlag, "검증비고",
        each null,
        type text
    ),

    // ========== STEP 13: Calculate Debit-Credit Balance Check ==========
    // Add helper column to verify 잔액 = 차변 - 대변
    AddedBalanceCheck = Table.AddColumn(AddedVerificationNote, "차대평형확인",
        each if Number.Abs([잔액] - ([차변] - [대변])) < 0.01 then "정상" else "불일치",
        type text
    ),

    // ========== STEP 14: Sort by Corporate Code and Account Code ==========
    SortedTable = Table.Sort(AddedBalanceCheck, {
        {"법인코드", Order.Ascending},
        {"계정코드", Order.Ascending}
    }),

    // ========== STEP 15: Add Index Column for Reference ==========
    AddedIndex = Table.AddIndexColumn(SortedTable, "행번호", 1, 1, Int64.Type),

    // ========== STEP 16: Add Data Source Indicator ==========
    AddedSourceFlag = Table.AddColumn(AddedIndex, "데이터출처",
        each "로컬파일",
        type text
    ),

    // ========== FINAL OUTPUT ==========
    FinalTable = AddedSourceFlag

in
    FinalTable

/*
  Expected Output Schema (Identical to PowerQuery_PTB.m):
  - 행번호 (Index): Int64
  - 법인코드 (Corporate Code): Text
  - 법인명 (Corporate Name): Text
  - 계정코드 (Account Code): Text
  - 계정과목명 (Account Name): Text
  - 차변 (Debit): Currency
  - 대변 (Credit): Currency
  - 잔액 (Balance): Currency
  - 통화코드 (Currency Code): Text
  - 결산연월 (Closing Period): Text
  - 비고 (Remarks): Text
  - 복합키 (Composite Key): Text
  - PwC_CoA: Text (nullable)
  - PwC_CoA_Name: Text (nullable)
  - PwC_Level1: Text (nullable)
  - PwC_Level2: Text (nullable)
  - PwC_Level3: Text (nullable)
  - 매핑상태 (Mapping Status): Text
  - 검증플래그 (Verification Flag): Text (nullable)
  - 검증비고 (Verification Note): Text (nullable)
  - 차대평형확인 (Balance Check): Text ["정상", "불일치"]
  - 데이터출처 (Data Source): Text ["로컬파일"]

  File Path Examples:
  Windows:
    "C:\Users\JaewooKim\Desktop\Project\HRE\참고\12월 말 합계잔액시산표_에이치알이_수정완료(01.13).xls"

  Mac:
    "/Users/jaewookim/Desktop/Project/HRE/참고/12월 말 합계잔액시산표_에이치알이_수정완료(01.13).xls"

  Common Sheet Names:
  - "Sheet1"
  - "합계잔액시산표"
  - "PTB"
  - "Data"

  Column Name Variations to Check:
  - If column names differ, update STEP 4 SelectedColumns list
  - Common variations:
    - "계정코드" vs "Account Code" vs "Acct Code"
    - "계정과목명" vs "Account Name" vs "계정명"
    - "차변" vs "Debit" vs "DR"
    - "대변" vs "Credit" vs "CR"
    - "잔액" vs "Balance" vs "Net"

  Troubleshooting:
  1. "Column not found" error:
     - Open Excel file and verify exact column names
     - Update STEP 4 column list to match
     - Check for extra spaces or special characters

  2. "File not found" error:
     - Verify FilePath is absolute path
     - Check for Korean characters in path (ensure UTF-8 encoding)
     - Try using raw string with forward slashes

  3. Data type errors:
     - Check if amounts are stored as text in Excel
     - Use Text.Clean() or Value.Replace() to clean data
     - Verify currency format matches locale

  4. Performance issues:
     - Large files (>100MB) may be slow
     - Consider filtering by 결산연월 early in the query
     - Use Table.Buffer() for repeated references

  Migration to SharePoint:
  1. Test all VBA procedures with this local file query
  2. Once validated, create PowerQuery_PTB.m with SharePoint connection
  3. Update VBA code to reference "PTB" instead of "PTB_Local"
  4. Keep this query as backup for offline work

  Best Practices:
  - Store source file in consistent location
  - Use version control for source files (e.g., Git)
  - Document any manual adjustments made to source file
  - Maintain parallel SharePoint and local file workflows during transition
*/
