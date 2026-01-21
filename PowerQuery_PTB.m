/*
  Power Query Template: PTB (Pre-Trial Balance) from SharePoint
  Version: 2.0
  Date: 2026-01-21

  Purpose:
  - Connect to HRE SharePoint consolidation site
  - Load PTB (합계잔액시산표) data from SharePoint List
  - Transform and normalize data for CoA mapping
  - Add placeholder columns for PwC standardized CoA

  Installation Instructions:
  1. Open Power Query Editor (데이터 > 쿼리 및 연결 > 쿼리 편집)
  2. Create new blank query (새 원본 > 기타 원본 > 빈 쿼리)
  3. Open Advanced Editor (고급 편집기)
  4. Paste this M code
  5. Replace [SHAREPOINT_SITE_URL] with actual HRE SharePoint site URL
  6. Replace [LIST_NAME] with actual SharePoint list name (e.g., "PTB_Data")
  7. Rename query to "PTB"
  8. Close & Load to Table

  Required Permissions:
  - SharePoint site read access
  - Office 365 authentication

  Dependencies:
  - SharePoint Online connector
  - Office 365 credentials

  Maintenance Notes:
  - Update column mappings if SharePoint list schema changes
  - Verify data types match source system
  - Check for new required columns in PwC CoA standards
*/

let
    // ========== CONNECTION PARAMETERS ==========
    // TODO: Replace with actual SharePoint site URL
    SharePointSiteUrl = "[SHAREPOINT_SITE_URL]",  // Example: "https://pwckorea.sharepoint.com/sites/HRE_Consolidation"
    ListName = "[LIST_NAME]",  // Example: "PTB_Data" or "합계잔액시산표"

    // ========== STEP 1: Connect to SharePoint List ==========
    Source = SharePoint.Tables(SharePointSiteUrl, [ApiVersion = 15]),

    // ========== STEP 2: Select Target List ==========
    SelectedList = Source{[Title=ListName]}[Items],

    // ========== STEP 3: Select Required Columns ==========
    // Adjust column names based on actual SharePoint list schema
    SelectedColumns = Table.SelectColumns(SelectedList, {
        "법인코드",           // Corporate Code
        "법인명",             // Corporate Name
        "계정코드",           // Account Code
        "계정과목명",         // Account Name (from client system)
        "차변",               // Debit Amount
        "대변",               // Credit Amount
        "잔액",               // Balance (Net)
        "통화코드",           // Currency Code
        "결산연월",           // Closing Period (YYYY-MM format)
        "비고"                // Remarks
    }),

    // ========== STEP 4: Data Type Enforcement ==========
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

    // ========== STEP 5: Add Composite Key ==========
    AddedKey = Table.AddColumn(TypedColumns, "복합키",
        each [법인코드] & "_" & [계정코드],
        type text
    ),

    // ========== STEP 6: Add PwC CoA Placeholder Columns ==========
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

    // ========== STEP 7: Add Mapping Status Column ==========
    AddedMappingStatus = Table.AddColumn(AddedPwCLevel3, "매핑상태",
        each if [PwC_CoA] = null then "미완료" else "완료",
        type text
    ),

    // ========== STEP 8: Add Verification Columns ==========
    AddedVerificationFlag = Table.AddColumn(AddedMappingStatus, "검증플래그",
        each null,
        type text
    ),

    AddedVerificationNote = Table.AddColumn(AddedVerificationFlag, "검증비고",
        each null,
        type text
    ),

    // ========== STEP 9: Filter Out Invalid Rows ==========
    // Remove rows with null or empty account codes
    FilteredRows = Table.SelectRows(AddedVerificationNote,
        each [계정코드] <> null and [계정코드] <> ""
    ),

    // ========== STEP 10: Sort by Corporate Code and Account Code ==========
    SortedTable = Table.Sort(FilteredRows, {
        {"법인코드", Order.Ascending},
        {"계정코드", Order.Ascending}
    }),

    // ========== STEP 11: Add Index Column for Reference ==========
    AddedIndex = Table.AddIndexColumn(SortedTable, "행번호", 1, 1, Int64.Type),

    // ========== FINAL OUTPUT ==========
    FinalTable = AddedIndex

in
    FinalTable

/*
  Expected Output Schema:
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

  Usage Notes:
  - This query should be refreshed after SPO connection is configured
  - PwC CoA columns will be populated by VBA procedures (Fill_Input_Table, Fill_CoA_Table)
  - Mapping status automatically updates based on PwC_CoA presence
  - For local file testing, use PowerQuery_PTB_LocalFile.m instead

  Troubleshooting:
  - If connection fails, verify SharePoint site URL and list name
  - Check Office 365 authentication status
  - Ensure SharePoint list has all required columns
  - For column name mismatches, update STEP 3 column selection
*/
