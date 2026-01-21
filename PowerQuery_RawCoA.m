/*
  Power Query Template: Raw_CoA (Historical CoA Mappings) from SharePoint
  Version: 2.0
  Date: 2026-01-21

  Purpose:
  - Connect to HRE SharePoint site for historical CoA mappings
  - Load Raw_CoA data (법인별 계정과목 매핑 이력)
  - Detect and categorize variant types (변동성 유형 감지)
  - Create composite keys for efficient lookup
  - Support synchronization with current PTB data

  Installation Instructions:
  1. Open Power Query Editor (데이터 > 쿼리 및 연결 > 쿼리 편집)
  2. Create new blank query (새 원본 > 기타 원본 > 빈 쿼리)
  3. Open Advanced Editor (고급 편집기)
  4. Paste this M code
  5. Replace [SHAREPOINT_SITE_URL] with actual HRE SharePoint site URL
  6. Replace [LIST_NAME] with actual SharePoint list name (e.g., "Raw_CoA_History")
  7. Rename query to "Raw_CoA"
  8. Close & Load to Table

  Variant Type Detection Logic:
  - "고정" (Fixed): Account code consistently maps to same PwC CoA across periods
  - "변동" (Variable): Account code maps to different PwC CoA in different periods
  - "신규" (New): Account code has no historical mapping
  - "폐기" (Obsolete): Historical mapping exists but no longer in current PTB
*/

let
    // ========== CONNECTION PARAMETERS ==========
    // TODO: Replace with actual SharePoint site URL
    SharePointSiteUrl = "[SHAREPOINT_SITE_URL]",  // Example: "https://pwckorea.sharepoint.com/sites/HRE_Consolidation"
    ListName = "[LIST_NAME]",  // Example: "Raw_CoA_History" or "계정과목_매핑_이력"

    // ========== STEP 1: Connect to SharePoint List ==========
    Source = SharePoint.Tables(SharePointSiteUrl, [ApiVersion = 15]),

    // ========== STEP 2: Select Target List ==========
    SelectedList = Source{[Title=ListName]}[Items],

    // ========== STEP 3: Select Required Columns ==========
    // Adjust column names based on actual SharePoint list schema
    SelectedColumns = Table.SelectColumns(SelectedList, {
        "법인코드",           // Corporate Code
        "법인명",             // Corporate Name
        "계정코드",           // Account Code (from client system)
        "계정과목명",         // Account Name (from client system)
        "PwC_CoA",           // PwC Standard CoA Code (mapped)
        "PwC_CoA_Name",      // PwC Standard CoA Name
        "PwC_Level1",        // PwC Level 1 (재무상태표/손익계산서)
        "PwC_Level2",        // PwC Level 2 (자산/부채/자본/수익/비용)
        "PwC_Level3",        // PwC Level 3 (유동/비유동 등)
        "결산연월",           // Closing Period (YYYY-MM format)
        "매핑일자",           // Mapping Date
        "매핑자",             // Mapper Name
        "비고"                // Remarks
    }),

    // ========== STEP 4: Data Type Enforcement ==========
    TypedColumns = Table.TransformColumnTypes(SelectedColumns, {
        {"법인코드", type text},
        {"법인명", type text},
        {"계정코드", type text},
        {"계정과목명", type text},
        {"PwC_CoA", type text},
        {"PwC_CoA_Name", type text},
        {"PwC_Level1", type text},
        {"PwC_Level2", type text},
        {"PwC_Level3", type text},
        {"결산연월", type text},
        {"매핑일자", type date},
        {"매핑자", type text},
        {"비고", type text}
    }),

    // ========== STEP 5: Filter Out Invalid Rows ==========
    FilteredRows = Table.SelectRows(TypedColumns,
        each [계정코드] <> null and [계정코드] <> "" and
             [PwC_CoA] <> null and [PwC_CoA] <> ""
    ),

    // ========== STEP 6: Add Composite Key (법인코드 + 계정코드) ==========
    AddedCompositeKey = Table.AddColumn(FilteredRows, "복합키",
        each [법인코드] & "_" & [계정코드],
        type text
    ),

    // ========== STEP 7: Detect Variant Type Per Composite Key ==========
    // Group by composite key and count distinct PwC_CoA mappings
    GroupedByKey = Table.Group(AddedCompositeKey, {"복합키"}, {
        {"매핑건수", each Table.RowCount(_), Int64.Type},
        {"고유_PwC_CoA_개수", each List.Count(List.Distinct([PwC_CoA])), Int64.Type},
        {"상세데이터", each _, type table}
    }),

    // ========== STEP 8: Add Variant Type Classification ==========
    AddedVariantType = Table.AddColumn(GroupedByKey, "변동성유형",
        each if [고유_PwC_CoA_개수] = 1 then "고정"      // Fixed: Always maps to same CoA
             else if [고유_PwC_CoA_개수] > 1 then "변동"  // Variable: Maps to different CoA
             else "미분류",                               // Unclassified
        type text
    ),

    // ========== STEP 9: Expand Detail Data with Variant Type ==========
    ExpandedDetails = Table.ExpandTableColumn(AddedVariantType, "상세데이터", {
        "법인코드", "법인명", "계정코드", "계정과목명",
        "PwC_CoA", "PwC_CoA_Name",
        "PwC_Level1", "PwC_Level2", "PwC_Level3",
        "결산연월", "매핑일자", "매핑자", "비고"
    }),

    // ========== STEP 10: Add Full Composite Key (with Variant Type) ==========
    AddedFullKey = Table.AddColumn(ExpandedDetails, "전체복합키",
        each [법인코드] & "_" & [계정코드] & "_" & [변동성유형],
        type text
    ),

    // ========== STEP 11: Rank Mappings by Date (Most Recent = 1) ==========
    AddedRank = Table.AddColumn(AddedFullKey, "매핑순위",
        each Table.PositionOf(
            Table.Sort(
                Table.SelectRows(AddedFullKey, (r) => r[복합키] = [복합키]),
                {{"매핑일자", Order.Descending}, {"결산연월", Order.Descending}}
            ),
            _,
            Occurrence.First
        ) + 1,
        Int64.Type
    ),

    // ========== STEP 12: Flag Most Recent Mapping ==========
    AddedLatestFlag = Table.AddColumn(AddedRank, "최신매핑여부",
        each if [매핑순위] = 1 then "최신" else "이력",
        type text
    ),

    // ========== STEP 13: Add Confidence Score ==========
    // Higher score = more reliable mapping
    // Fixed type + recent mapping = highest confidence
    AddedConfidenceScore = Table.AddColumn(AddedLatestFlag, "신뢰도점수",
        each (
            (if [변동성유형] = "고정" then 50 else 0) +
            (if [최신매핑여부] = "최신" then 30 else 0) +
            (if [매핑건수] >= 3 then 20 else [매핑건수] * 5)
        ),
        Int64.Type
    ),

    // ========== STEP 14: Sort by Confidence Score (Descending) ==========
    SortedByConfidence = Table.Sort(AddedConfidenceScore, {
        {"복합키", Order.Ascending},
        {"신뢰도점수", Order.Descending},
        {"매핑일자", Order.Descending}
    }),

    // ========== STEP 15: Add Index Column ==========
    AddedIndex = Table.AddIndexColumn(SortedByConfidence, "행번호", 1, 1, Int64.Type),

    // ========== FINAL OUTPUT ==========
    FinalTable = AddedIndex

in
    FinalTable

/*
  Expected Output Schema:
  - 행번호 (Index): Int64
  - 복합키 (Composite Key): Text (법인코드_계정코드)
  - 매핑건수 (Mapping Count): Int64
  - 고유_PwC_CoA_개수 (Unique PwC CoA Count): Int64
  - 변동성유형 (Variant Type): Text ["고정", "변동", "미분류"]
  - 법인코드 (Corporate Code): Text
  - 법인명 (Corporate Name): Text
  - 계정코드 (Account Code): Text
  - 계정과목명 (Account Name): Text
  - PwC_CoA: Text
  - PwC_CoA_Name: Text
  - PwC_Level1: Text
  - PwC_Level2: Text
  - PwC_Level3: Text
  - 결산연월 (Closing Period): Text
  - 매핑일자 (Mapping Date): Date
  - 매핑자 (Mapper): Text
  - 비고 (Remarks): Text
  - 전체복합키 (Full Composite Key): Text
  - 매핑순위 (Mapping Rank): Int64 (1 = most recent)
  - 최신매핑여부 (Latest Flag): Text ["최신", "이력"]
  - 신뢰도점수 (Confidence Score): Int64 (0-100)

  Usage in VBA Workflows:
  1. Use "복합키" to lookup historical mappings for current PTB accounts
  2. Filter by "최신매핑여부" = "최신" for auto-fill suggestions
  3. Check "변동성유형":
     - "고정": Auto-fill with confidence
     - "변동": Prompt user to verify/select from history
     - "미분류": Manual mapping required
  4. Use "신뢰도점수" to prioritize mapping suggestions

  Synchronization with PTB:
  - Join PTB.복합키 with Raw_CoA.복합키
  - Identify "신규" accounts: In PTB but not in Raw_CoA
  - Identify "폐기" accounts: In Raw_CoA but not in PTB

  Troubleshooting:
  - If no historical data appears, verify SharePoint list has records
  - For variant type misclassification, check PwC_CoA normalization
  - If confidence scores seem incorrect, adjust STEP 13 weights
  - For performance issues, filter by 결산연월 to limit historical range
*/
