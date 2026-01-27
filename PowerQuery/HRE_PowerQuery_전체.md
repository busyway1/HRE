# HRE SharePoint Online 전체 구조 Power Query 가이드

**버전**: 1.10
**작성일**: 2026-01-25
**용도**: SPO 다중 법인 폴더 순회 및 PTB 테이블 생성

---

## 1. 개요

### 1.1 쿼리 아키텍처 (전체)

```
[Phase 1: 설정 쿼리 - 연결 전용]
  ├── 메인주소 (Path 테이블 참조)
  └── 회사결산_주소 (TempPath 테이블 참조)

[Phase 2: 데이터 수집 - 연결 전용]
  ├── Raw_Data (다중 법인 폴더 순회)
  │     ├── 의존: 메인주소, 회사결산_주소
  │     └── 출력: 법인명_폴더, Name, Content, Folder Path
  └── 디렉터리 (SPO 폴더 구조)
        └── frmDirectory TreeView용

[Phase 3: 데이터 변환 - 연결 전용]
  └── TB
        ├── 의존: Raw_Data, fnProcessSingleFile
        └── 출력: 법인코드, 법인명, 법인별CoA, 법인별계정과목명, 당기

[Phase 4: 최종 출력 - 테이블 로드]
  ├── PTB (BSPL 시트)
  │     ├── 의존: TB
  │     └── 테이블명: PTB (9컬럼)
  ├── Link (HideSheet)  ← 신규 추가
  │     └── SPO 법인별 파일 링크
  └── Link_취득_처분 (HideSheet)  ← 신규 추가
        └── 취득/처분 법인 파일 링크

[Phase 5: 마스터 참조 쿼리 - 연결 전용] (선택사항)
  ├── Master (CoAMaster 테이블 참조)
  └── Corp (CorpMaster.Corp 테이블 참조)
```

### 1.2 쿼리 개수 및 로드 설정 (완전판)

| 쿼리명 | Phase | 로드 설정 | 출력 위치 | 용도 |
|--------|-------|----------|----------|------|
| 메인주소 | 1 | 연결 전용 | - | SPO 사이트 URL |
| 회사결산_주소 | 1 | 연결 전용 | - | 결산 폴더 경로 |
| 디렉터리 | 2 | **테이블 로드** | DirectoryURL!$A$1 | frmDirectory TreeView |
| Raw_Data | 2 | 연결 전용 | - | SPO 파일 목록 |
| fnProcessSingleFile | 3 | 연결 전용 (함수) | - | 단일 파일 처리 |
| TB | 3 | 연결 전용 | - | 시산표 데이터 변환 |
| PTB | 4 | **테이블 로드** | BSPL!$B$4 | 최종 연결시산표 |
| **Link** | 4 | **테이블 로드** | **HideSheet!$A$5** | 법인별 파일 링크 |
| **Link_취득_처분** | 4 | **테이블 로드** | **HideSheet!$E$5** | 취득/처분 파일 링크 |
| Master | 5 | 연결 전용 | - | CoA 마스터 참조 (선택) |
| Corp | 5 | 연결 전용 | - | 법인 마스터 참조 (선택) |

### 1.3 Excel 테이블 구조 (완전판)

#### 1.3.1 Power Query로 채워지는 테이블

| 시트 | 테이블명 | 컬럼 수 | 데이터 소스 |
|------|----------|---------|-------------|
| BSPL | PTB | 9 | Power Query (TB 쿼리) |
| HideSheet | Link | 3 | Power Query (Link 쿼리) |
| HideSheet | Link_취득_처분 | 3 | Power Query (Link_취득_처분 쿼리) |
| DirectoryURL | 디렉터리 | 10 | Power Query (디렉터리 쿼리) |

#### 1.3.2 수동 생성 필요 테이블 (Excel Tables)

| 시트 | 테이블명 | 용도 | VBA 참조 |
|------|----------|------|----------|
| HideSheet | Path | SPO 메인 URL | 메인주소 쿼리 |
| HideSheet | TempPath | 결산 폴더 경로 | 회사결산_주소 쿼리 |
| HideSheet | 결산연월 | 결산연도/월 | VBA 전역 |
| HideSheet | **People_Work** | 담당자 목록 | frmScope, frmAddPerson |
| CorpMaster | Corp | 법인 현황 | VBA 전역 |
| CorpMaster | **Mode** | Scope 모드 | frmScope |
| CorpMaster | **FS** | 재무제표 유형 | frmScope |
| CorpMaster | **Who** | 담당자 선택 | frmScope |
| CoAMaster | Master | PwC 표준 CoA | VBA CoA 매핑 |
| CorpCoA | Raw_CoA | CoA 매핑 이력 | VBA 자동 추가 |
| AddCoA | **CoA_Input** | CoA 입력용 | VBA Fill_Input_Table |
| ADBS | **AD_BS** | 취득/처분 BS | VBA 취득/처분 |

### 1.4 SPO 폴더 구조

```
★시연용폴더★/2512/
├── HRE/                    ← 법인명 (폴더명)
│   └── *시산표*.xlsx       ← 파일 내: 법인코드 + 법인명
├── 예천솔라/
│   └── *시산표*.xlsx
├── 당진솔라/
│   └── *시산표*.xlsx
└── ...
```

---

## 2. Phase 1: 설정 쿼리

### 2.1 메인주소 쿼리

**용도**: HideSheet.Path 테이블에서 SPO 사이트 URL 읽기

```m
// 쿼리명: 메인주소
// 로드: 연결 전용
let
    // HideSheet의 Path 테이블 참조
    Source = Excel.CurrentWorkbook(){[Name="Path"]}[Content],

    // 첫 번째 행 가져오기
    FirstRow = Source{0},

    // 컬럼명에 상관없이 첫 번째 값 추출 (유연성 확보)
    PathValue = Record.FieldValues(FirstRow){0},

    // URL 정리 (공백 제거)
    CleanedPath = Text.Trim(Text.From(PathValue))
in
    CleanedPath
```

> **참고**: `Record.FieldValues(FirstRow){0}` 사용으로 컬럼명이 "Path"가 아니어도 동작합니다.

**출력 예시**: `https://pwckor.sharepoint.com/sites/KR-ASR-HRE_Consolidation`

---

### 2.2 회사결산_주소 쿼리

**용도**: HideSheet.TempPath 테이블에서 결산 폴더 경로 읽기

```m
// 쿼리명: 회사결산_주소
// 로드: 연결 전용
let
    // HideSheet의 TempPath 테이블 참조
    Source = Excel.CurrentWorkbook(){[Name="TempPath"]}[Content],

    // 테이블이 비어있는지 확인
    TempPathValue = if Table.RowCount(Source) = 0 then null
                    else Source{0}[경로],

    // 값 정리
    CleanedValue = if TempPathValue = null then null
                   else Text.Trim(Text.From(TempPathValue)),

    // 상대 경로 추출 (전체 URL인 경우 처리)
    RelativePath = if CleanedValue = null then null
                   else if Text.Contains(CleanedValue, "/sites/") then
                       Text.AfterDelimiter(CleanedValue, "/sites/KR-ASR-HRE_Consolidation")
                   else CleanedValue,

    // 공백 제거
    NoSpacePath = if RelativePath = null then null
                  else Text.Replace(RelativePath, " ", ""),

    // /Shared Documents 접두사 확인 및 추가
    FinalPath = if NoSpacePath = null then null
                else if Text.Contains(NoSpacePath, "Shared Documents") then NoSpacePath
                else "/Shared Documents" & (if Text.StartsWith(NoSpacePath, "/") then NoSpacePath else "/" & NoSpacePath)
in
    FinalPath
```

> **중요**: frmDirectory에서 선택한 경로가 전체 URL로 저장될 수 있으므로, 상대 경로로 변환하는 로직이 포함되어 있습니다.

**출력 예시**: `/Shared Documents/★시연용폴더★/2512`

---

## 3. Phase 2: 데이터 수집

### 3.1 디렉터리 쿼리 (frmDirectory TreeView용)

**용도**: SPO 폴더 구조를 가져와 frmDirectory TreeView에 표시

```m
// 쿼리명: 디렉터리
// 로드: 테이블 (DirectoryURL 시트)
let
    // 1. SPO 연결
    메인주소 = 메인주소,
    Source = SharePoint.Files(메인주소, [ApiVersion = 15]),

    // 2. 폴더 경로만 추출 (중복 제거)
    FolderPaths = Table.Distinct(
        Table.SelectColumns(Source, {"Folder Path"})
    ),

    // 3. 경로를 레벨별로 분리
    SplitPath = Table.AddColumn(FolderPaths, "PathParts", each
        let
            path = [Folder Path],
            // /sites/.../Shared Documents/ 이후 부분만 추출
            afterShared = if Text.Contains(path, "Shared Documents")
                          then Text.AfterDelimiter(path, "Shared Documents/")
                          else path,
            // 빈 부분 제거
            parts = List.Select(Text.Split(afterShared, "/"), each _ <> "")
        in
            parts
    ),

    // 4. 레벨 컬럼 추가 (최대 10레벨)
    AddLevel1 = Table.AddColumn(SplitPath, "Level1", each try [PathParts]{0} otherwise null),
    AddLevel2 = Table.AddColumn(AddLevel1, "Level2", each try [PathParts]{1} otherwise null),
    AddLevel3 = Table.AddColumn(AddLevel2, "Level3", each try [PathParts]{2} otherwise null),
    AddLevel4 = Table.AddColumn(AddLevel3, "Level4", each try [PathParts]{3} otherwise null),
    AddLevel5 = Table.AddColumn(AddLevel4, "Level5", each try [PathParts]{4} otherwise null),
    AddLevel6 = Table.AddColumn(AddLevel5, "Level6", each try [PathParts]{5} otherwise null),
    AddLevel7 = Table.AddColumn(AddLevel6, "Level7", each try [PathParts]{6} otherwise null),
    AddLevel8 = Table.AddColumn(AddLevel7, "Level8", each try [PathParts]{7} otherwise null),
    AddLevel9 = Table.AddColumn(AddLevel8, "Level9", each try [PathParts]{8} otherwise null),
    AddLevel10 = Table.AddColumn(AddLevel9, "Level10", each try [PathParts]{9} otherwise null),

    // 5. 최종 컬럼 선택
    FinalColumns = Table.SelectColumns(AddLevel10,
        {"Level1", "Level2", "Level3", "Level4", "Level5",
         "Level6", "Level7", "Level8", "Level9", "Level10"}
    ),

    // 6. 중복 제거 및 정렬
    Distinct = Table.Distinct(FinalColumns),
    Sorted = Table.Sort(Distinct, {{"Level1", Order.Ascending}, {"Level2", Order.Ascending}})
in
    Sorted
```

**로드 위치**: DirectoryURL 시트 (테이블명: `디렉터리`)

---

### 3.2 Raw_Data 쿼리

**용도**: SPO 다중 법인 폴더 순회, 시산표 파일 목록 수집

```m
// 쿼리명: Raw_Data
// 로드: 연결 전용
// 의존: 메인주소, 회사결산_주소
let
    // ========== 1. 설정 쿼리 참조 ==========
    메인주소 = 메인주소,
    회사결산_주소 = 회사결산_주소,

    // ========== 2. SharePoint 연결 ==========
    Source = SharePoint.Files(메인주소, [ApiVersion = 15]),

    // ========== 3. 결산 폴더 필터링 ==========
    // 회사결산_주소에 해당하는 폴더만 선택
    FilteredFolder = Table.SelectRows(Source,
        each Text.Contains([Folder Path], 회사결산_주소)
    ),

    // ========== 4. 시산표 파일만 필터링 ==========
    // 파일명에 "시산표" 포함된 Excel 파일
    FilteredFiles = Table.SelectRows(FilteredFolder,
        each Text.Contains([Name], "시산표") and
             (Text.EndsWith([Name], ".xlsx") or Text.EndsWith([Name], ".xls"))
    ),

    // ========== 5. 법인명 추출 (폴더명) ==========
    // Folder Path에서 마지막 폴더명 추출 (법인명 fallback용)
    AddCorpFolder = Table.AddColumn(FilteredFiles, "법인명_폴더", each
        let
            path = [Folder Path],
            // 마지막 "/" 이전까지 자름
            trimmed = if Text.EndsWith(path, "/")
                      then Text.Start(path, Text.Length(path) - 1)
                      else path,
            // 마지막 "/" 이후 = 폴더명
            parts = Text.Split(trimmed, "/"),
            lastPart = List.Last(parts)
        in
            lastPart,
        type text
    ),

    // ========== 6. 필요 컬럼만 선택 ==========
    SelectedColumns = Table.SelectColumns(AddCorpFolder,
        {"법인명_폴더", "Name", "Content", "Folder Path"}
    )
in
    SelectedColumns
```

**출력 스키마**:

| 컬럼 | 타입 | 설명 | 예시 |
|------|------|------|------|
| 법인명_폴더 | text | 폴더명 (fallback용) | `HRE` |
| Name | text | 파일명 | `시산표_202512.xlsx` |
| Content | binary | 파일 바이너리 | (binary) |
| Folder Path | text | 전체 경로 | `/sites/.../2512/HRE/` |

---

## 4. Phase 3: 데이터 변환

### 4.1 fnProcessSingleFile 함수

**용도**: 단일 시산표 파일에서 법인코드, 법인명, CoA, 당기 추출

**실제 파일 구조** (2026-01-24 검증):
```
Row 1: null | null | "합계잔액시산표"
Row 4: "회 사 : 1000.에이치알이주..." ← 메타셀
Row 7-8: 헤더 행
Row 9+: 데이터 (Column1=차변, Column3=계정과목, Column5=대변)
```

**Column3 형식**: `[1030000] 보 통 예 금` (코드+계정명 통합)

```m
// 쿼리명: fnProcessSingleFile
// 로드: 연결 전용 (함수)
// 용도: 단일 시산표 파일에서 데이터 추출

let
    // --------------------------------------------------------------------
    // 0) 공통 유틸
    // --------------------------------------------------------------------
    정규화 = (val as any) as text =>
        let
            str = if val = null then "" else Text.From(val),
            cleaned = Text.Trim(Text.Clean(str))
        in
            cleaned,

    법인코드_추출 = (메타셀 as any) as nullable text =>
        let
            t0 = 정규화(메타셀),
            t1 =
                if Text.Contains(t0, ":")
                then Text.Trim(Text.AfterDelimiter(t0, ":"))
                else t0,
            t2 =
                if Text.Contains(t1, ".")
                then Text.Trim(Text.BeforeDelimiter(t1, "."))
                else Text.Trim(t1)
        in
            if t2 = "" then null else t2,

    법인명_추출 = (메타셀 as any) as nullable text =>
        let
            t0 = 정규화(메타셀),
            t1 =
                if Text.Contains(t0, ":")
                then Text.Trim(Text.AfterDelimiter(t0, ":"))
                else t0,
            name =
                if Text.Contains(t1, ".")
                then Text.Trim(Text.AfterDelimiter(t1, "."))
                else null
        in
            if name = null or name = "" then null else name,

    // "[1030000] 보 통 예 금" → "1030000"
    계정코드_추출 = (계정과목 as any) as nullable text =>
        let
            str = 정규화(계정과목),
            hasCode = Text.Contains(str, "[") and Text.Contains(str, "]"),
            code =
                if hasCode
                then try Text.Trim(Text.BetweenDelimiters(str, "[", "]")) otherwise null
                else null
        in
            if code = null or code = "" then null else code,

    // "[1030000] 보 통 예 금" → "보 통 예 금"
    계정명_추출 = (계정과목 as any) as nullable text =>
        let
            str = 정규화(계정과목),
            hasCode = Text.Contains(str, "]"),
            name =
                if hasCode
                then Text.Trim(Text.AfterDelimiter(str, "]"))
                else str,
            out = Text.Trim(name)
        in
            if out = "" then null else out,

    // --------------------------------------------------------------------
    // 1) 단일 파일 처리 함수
    // --------------------------------------------------------------------
    fnProcessSingleFile = (fileContent as binary, 법인명_폴더 as text) as table =>
        let
            // 1) Excel 로드
            Source = Excel.Workbook(fileContent, null, true),
            FirstSheet = Source{0}[Data],

            // 2) 메타셀: Row 4(인덱스 3), Column1
            메타셀 = try FirstSheet{3}[Column1] otherwise null,
            법인코드 = 법인코드_추출(메타셀),
            법인명_메타 = 법인명_추출(메타셀),
            법인명_최종 =
                if 법인명_메타 <> null and 법인명_메타 <> ""
                then 법인명_메타
                else 법인명_폴더,

            // 3) 데이터 시작: Row 9부터 (8행 skip)
            DataRows = Table.Skip(FirstSheet, 8),

            // 4) Column3에 [ ]가 있는 행만 필터링
            FilteredData =
                Table.SelectRows(
                    DataRows,
                    each
                        let
                            col3 = try 정규화(Record.FieldOrDefault(_, "Column3", null)) otherwise ""
                        in
                            Text.Contains(col3, "[") and Text.Contains(col3, "]")
                ),

            // 5) 컬럼 추가
            Added법인코드 =
                Table.AddColumn(FilteredData, "법인코드", each 법인코드, type text),

            Added법인명 =
                Table.AddColumn(Added법인코드, "법인명", each 법인명_최종, type text),

            Added법인별CoA =
                Table.AddColumn(
                    Added법인명,
                    "법인별CoA",
                    each 계정코드_추출(Record.FieldOrDefault(_, "Column3", null)),
                    type text
                ),

            Added계정과목명 =
                Table.AddColumn(
                    Added법인별CoA,
                    "법인별계정과목명",
                    each 계정명_추출(Record.FieldOrDefault(_, "Column3", null)),
                    type text
                ),

            Added당기 =
                Table.AddColumn(
                    Added계정과목명,
                    "당기",
                    each
                        let
                            차변 = try Number.From(Record.FieldOrDefault(_, "Column1", null)) otherwise 0,
                            대변 = try Number.From(Record.FieldOrDefault(_, "Column5", null)) otherwise 0
                        in
                            차변 - 대변,
                    type number
                ),

            // 6) 최종 컬럼
            FinalColumns =
                Table.SelectColumns(
                    Added당기,
                    {"법인코드", "법인명", "법인별CoA", "법인별계정과목명", "당기"}
                ),

            // 7) null/blank 제거
            Result =
                Table.SelectRows(
                    FinalColumns,
                    each [법인별CoA] <> null and [법인별CoA] <> ""
                )
        in
            Result
in
    fnProcessSingleFile
```

---

### 4.2 TB 쿼리

**용도**: Raw_Data의 각 파일에 fnProcessSingleFile 적용, 전체 데이터 병합

```m
// 쿼리명: TB
// 로드: 연결 전용
// 의존: Raw_Data, fnProcessSingleFile
let
    // ========== 1. Raw_Data 참조 ==========
    Source = Raw_Data,

    // ========== 2. fnProcessSingleFile 참조 ==========
    ProcessFile = fnProcessSingleFile,

    // ========== 3. 각 파일에 함수 적용 ==========
    AddProcessedData = Table.AddColumn(Source, "ProcessedData", each
        try ProcessFile([Content], [법인명_폴더])
        otherwise #table(
            {"법인코드", "법인명", "법인별CoA", "법인별계정과목명", "당기"},
            {}
        ),
        type table
    ),

    // ========== 4. 처리된 테이블 확장 ==========
    ExpandedData = Table.ExpandTableColumn(AddProcessedData, "ProcessedData",
        {"법인코드", "법인명", "법인별CoA", "법인별계정과목명", "당기"}
    ),

    // ========== 5. 필요 컬럼만 선택 ==========
    SelectedColumns = Table.SelectColumns(ExpandedData,
        {"법인코드", "법인명", "법인별CoA", "법인별계정과목명", "당기"}
    ),

    // ========== 6. null 행 제거 ==========
    FilteredRows = Table.SelectRows(SelectedColumns,
        each [법인코드] <> null and [법인별CoA] <> null
    ),

    // ========== 7. 데이터 타입 설정 ==========
    TypedColumns = Table.TransformColumnTypes(FilteredRows, {
        {"법인코드", type text},
        {"법인명", type text},
        {"법인별CoA", type text},
        {"법인별계정과목명", type text},
        {"당기", type number}
    })
in
    TypedColumns
```

**출력 스키마**:

| 컬럼 | 타입 | 설명 |
|------|------|------|
| 법인코드 | text | 메타데이터에서 추출 |
| 법인명 | text | 메타데이터 또는 폴더명 |
| 법인별CoA | text | 계정코드 (7자리) |
| 법인별계정과목명 | text | 계정과목명 |
| 당기 | number | 차변 - 대변 |

---

## 5. Phase 4: 최종 출력

### 5.1 PTB 쿼리

**용도**: CoA 매핑 컬럼 추가, BSPL 시트에 테이블 로드

```m
// 쿼리명: PTB
// 로드: 테이블 (BSPL!$B$4)
// 의존: TB
let
    // ========== 1. TB 쿼리 참조 ==========
    Source = TB,

    // ========== 2. PwC CoA 매핑 컬럼 추가 (VBA가 채움) ==========
    Add연결CoA = Table.AddColumn(Source, "연결CoA", each null, type text),
    Add연결계정명 = Table.AddColumn(Add연결CoA, "연결계정명", each null, type text),

    // ========== 3. 분류 컬럼 추가 (VBA가 채움) ==========
    Add대분류 = Table.AddColumn(Add연결계정명, "대분류", each null, type text),
    Add중분류 = Table.AddColumn(Add대분류, "중분류", each null, type text),

    // ========== 4. 금액 컬럼 (당기 복사) ==========
    // 부호 적용은 VBA에서 처리 (Master 테이블의 부호 참조)
    Add금액 = Table.AddColumn(Add중분류, "금액", each [당기], type number),

    // ========== 5. 최종 컬럼 순서 정렬 ==========
    // PTB 테이블 스키마: 9컬럼
    FinalColumns = Table.SelectColumns(Add금액, {
        "법인코드",          // 1
        "법인명",            // 2 ← 신규 추가
        "법인별CoA",         // 3
        "법인별계정과목명",   // 4
        "연결CoA",          // 5
        "연결계정명",       // 6
        "대분류",            // 7
        "중분류",            // 8
        "금액"               // 9
    }),

    // ========== 6. 정렬 ==========
    SortedTable = Table.Sort(FinalColumns, {
        {"법인코드", Order.Ascending},
        {"법인별CoA", Order.Ascending}
    })
in
    SortedTable
```

**최종 PTB 테이블 스키마 (9컬럼)**:

| 인덱스 | 컬럼명 | 타입 | 소스 | VBA 참조 |
|--------|--------|------|------|----------|
| 1 | 법인코드 | text | TB | `Range(1, 1)` |
| **2** | **법인명** | text | TB | `Range(1, 2)` ← **신규** |
| 3 | 법인별CoA | text | TB | `Range(1, 3)` |
| 4 | 법인별계정과목명 | text | TB | `Range(1, 4)` |
| 5 | 연결CoA | text | VBA | `Range(1, 5)` |
| 6 | 연결계정명 | text | VBA | `Range(1, 6)` |
| 7 | 대분류 | text | VBA | `Range(1, 7)` |
| 8 | 중분류 | text | VBA | `Range(1, 8)` |
| 9 | 금액 | number | TB.당기 | `Range(1, 9)` |

---

### 5.2 Link 쿼리 (신규 추가)

**용도**: Verify 시트에서 법인별 SPO 파일 링크 제공

**VBA 참조**: `mod_06_VerifySum.bas` Line 44, 160, 382

**로드 위치**: `HideSheet!$A$5` (테이블명: `Link`)

```m
// 쿼리명: Link
// 로드: 테이블 (HideSheet!$A$5)
// 테이블명: Link
// 의존: Raw_Data, Corp (CorpMaster)
let
    // ========== 1. 설정 ==========
    메인주소 = 메인주소,
    회사결산_주소 = 회사결산_주소,

    // ========== 2. Corp 테이블에서 Scope = "O"인 법인만 ==========
    CorpTable = Excel.CurrentWorkbook(){[Name="Corp"]}[Content],
    ScopedCorp = Table.SelectRows(CorpTable, each [Scope] = "O"),

    // ========== 3. SharePoint 파일 목록 ==========
    Source = SharePoint.Files(메인주소, [ApiVersion = 15]),

    // 결산 폴더 내 시산표 파일만
    FilteredFiles = Table.SelectRows(Source,
        each Text.Contains([Folder Path], 회사결산_주소) and
             Text.Contains([Name], "시산표") and
             (Text.EndsWith([Name], ".xlsx") or Text.EndsWith([Name], ".xls"))
    ),

    // ========== 4. 법인코드 추출 (파일 내 메타데이터) ==========
    AddCorpCode = Table.AddColumn(FilteredFiles, "법인코드", each
        let
            fileContent = [Content],
            // Excel 파일 열기
            wb = try Excel.Workbook(fileContent, null, true) otherwise null,
            sheet = if wb <> null then try wb{0}[Data] otherwise null else null,
            // Row 4에서 메타셀 추출
            metaCell = if sheet <> null then try sheet{3}[Column1] otherwise null else null,
            // 법인코드 파싱: "회사: 1000.에이치알이" → "1000"
            cleanMeta = if metaCell = null then null else Text.Trim(Text.Clean(Text.From(metaCell))),
            afterColon = if cleanMeta = null then null
                         else if Text.Contains(cleanMeta, ":")
                         then Text.Trim(Text.AfterDelimiter(cleanMeta, ":"))
                         else cleanMeta,
            code = if afterColon = null then null
                   else if Text.Contains(afterColon, ".")
                   then Text.Trim(Text.BeforeDelimiter(afterColon, "."))
                   else afterColon
        in
            code,
        type text
    ),

    // ========== 5. SPO 웹 링크 생성 ==========
    AddLink = Table.AddColumn(AddCorpCode, "Link", each
        let
            // 파일 경로에서 웹 URL 생성
            // 예: /sites/.../Shared Documents/폴더/파일.xlsx
            webUrl = 메인주소 & "/" & Text.Replace([Folder Path], "/sites/KR-ASR-HRE_Consolidation/", "") & [Name]
        in
            webUrl,
        type text
    ),

    // ========== 6. 최종 컬럼 ==========
    FinalColumns = Table.SelectColumns(AddLink, {"법인코드", "Link"}),

    // ========== 7. null 제거 및 정렬 ==========
    Filtered = Table.SelectRows(FinalColumns, each [법인코드] <> null),
    Sorted = Table.Sort(Filtered, {{"법인코드", Order.Ascending}})
in
    Sorted
```

**출력 스키마**:

| 컬럼 | 타입 | 설명 |
|------|------|------|
| 법인코드 | text | 시산표 파일 내 메타데이터에서 추출 |
| Link | text | SPO 웹 URL |

> **참고**: frmScope에서 Scope 설정 시 `qt.Refresh BackgroundQuery:=False`로 이 쿼리가 새로고침됩니다.

---

### 5.3 Link_취득_처분 쿼리 (신규 추가)

**용도**: 취득/처분 법인의 SPO 파일 링크

**VBA 참조**: `frmScope.frm` Line 174

**로드 위치**: `HideSheet!$E$5` (테이블명: `Link_취득_처분`)

> **주의**: 기존 테이블 이름이 `Corp_취득_처분`인 경우 `Link_취득_처분`으로 변경 필요

```m
// 쿼리명: Link_취득_처분
// 로드: 테이블 (HideSheet!$E$5)
// 테이블명: Link_취득_처분
// 의존: 메인주소, Corp 테이블
let
    // ========== 1. 설정 ==========
    메인주소 = 메인주소,

    // ========== 2. Corp 테이블에서 취득/처분 법인 필터 ==========
    // 취득일 또는 처분일이 "-"가 아닌 법인
    CorpTable = Excel.CurrentWorkbook(){[Name="Corp"]}[Content],
    ADCorp = Table.SelectRows(CorpTable, each
        ([취득일] <> null and [취득일] <> "-") or
        ([처분일] <> null and [처분일] <> "-")
    ),

    // ========== 3. SharePoint에서 해당 법인 파일 찾기 ==========
    Source = SharePoint.Files(메인주소, [ApiVersion = 15]),

    // 취득/처분 폴더 경로 패턴 (예: /취득처분/2512/)
    FilteredFiles = Table.SelectRows(Source,
        each (Text.Contains([Folder Path], "취득") or Text.Contains([Folder Path], "처분")) and
             Text.Contains([Name], "시산표") and
             (Text.EndsWith([Name], ".xlsx") or Text.EndsWith([Name], ".xls"))
    ),

    // ========== 4. 법인코드 추출 ==========
    AddCorpCode = Table.AddColumn(FilteredFiles, "법인코드", each
        let
            fileContent = [Content],
            wb = try Excel.Workbook(fileContent, null, true) otherwise null,
            sheet = if wb <> null then try wb{0}[Data] otherwise null else null,
            metaCell = if sheet <> null then try sheet{3}[Column1] otherwise null else null,
            cleanMeta = if metaCell = null then null else Text.Trim(Text.Clean(Text.From(metaCell))),
            afterColon = if cleanMeta = null then null
                         else if Text.Contains(cleanMeta, ":")
                         then Text.Trim(Text.AfterDelimiter(cleanMeta, ":"))
                         else cleanMeta,
            code = if afterColon = null then null
                   else if Text.Contains(afterColon, ".")
                   then Text.Trim(Text.BeforeDelimiter(afterColon, "."))
                   else afterColon
        in
            code,
        type text
    ),

    // ========== 5. SPO 웹 링크 생성 ==========
    AddLink = Table.AddColumn(AddCorpCode, "Link", each
        메인주소 & "/" & Text.Replace([Folder Path], "/sites/KR-ASR-HRE_Consolidation/", "") & [Name],
        type text
    ),

    // ========== 6. 최종 ==========
    FinalColumns = Table.SelectColumns(AddLink, {"법인코드", "Link"}),
    Filtered = Table.SelectRows(FinalColumns, each [법인코드] <> null),
    Sorted = Table.Sort(Filtered, {{"법인코드", Order.Ascending}})
in
    Sorted
```

---

## 6. Phase 5: 마스터 참조 쿼리 (선택사항)

### 6.1 Master 쿼리

**용도**: CoAMaster 테이블 참조 (VBA에서 직접 참조하므로 선택사항)

**Master 테이블 실제 컬럼명** (CoAMaster 시트):

| 컬럼명 | 타입 | PTB 매핑 |
|--------|------|----------|
| TB Account | text | → 연결CoA |
| Account Name | text | → 연결계정명 |
| BSPL | text | BS/PL 구분 |
| 대분류 | text | → 대분류 |
| 중분류 | text | → 중분류 |
| 소분류 | text | |
| 공시계정 | text | |
| 그룹사 보고용 | text | |
| 부호 | number | 금액 부호 적용 |
| 금액 | number | |

```m
// 쿼리명: Master
// 로드: 연결 전용 (또는 테이블)
let
    Source = Excel.CurrentWorkbook(){[Name="Master"]}[Content],
    TypedColumns = Table.TransformColumnTypes(Source, {
        {"TB Account", type text},
        {"Account Name", type text},
        {"BSPL", type text},
        {"대분류", type text},
        {"중분류", type text},
        {"소분류", type text},
        {"공시계정", type text},
        {"그룹사 보고용", type text},
        {"부호", Int64.Type},
        {"금액", type number}
    })
in
    TypedColumns
```

> **VBA 매핑 참고**: PTB 테이블의 `연결CoA`는 Master 테이블의 `TB Account`를 참조하여 채워집니다.

### 6.2 Corp 쿼리

**용도**: CorpMaster.Corp 테이블 참조 (VBA에서 직접 참조하므로 선택사항)

```m
// 쿼리명: Corp
// 로드: 연결 전용
let
    Source = Excel.CurrentWorkbook(){[Name="Corp"]}[Content],
    // Scope = "O"인 법인만 필터
    ScopedCorp = Table.SelectRows(Source, each [Scope] = "O")
in
    ScopedCorp
```

---

## 7. 수동 생성 필요 Excel 테이블 상세

### 7.1 HideSheet 테이블들

#### HideSheet 전체 레이아웃

```
     A        B        C        D        E        F        G        H        I        ...    Q        R        S
1  결산연도  결산월                     Path                               이름      경로          담당자이름  ...      ...
2   2025     12                        (URL)                              결산폴더  /Shared...    홍길동     ...      ...
3
4
5  법인코드   Link                      법인코드  Link
6  (PQ)      (PQ)                       (PQ)     (PQ)
   ↑ Link 테이블                        ↑ Link_취득_처분 테이블

   [A1:B2] 결산연월
   [E1:E2] Path
   [H1:I2] TempPath
   [Q1:S?] People_Work
   [A5:B?] Link (Power Query)
   [E5:F?] Link_취득_처분 (Power Query)
```

#### 7.1.1 Path 테이블 (E1:E2)

| E1 | 내용 |
|----|------|
| 헤더 | `Path` 또는 `메인주소` |
| 데이터 | `https://pwckor.sharepoint.com/sites/KR-ASR-HRE_Consolidation` |

**생성 방법**:
1. HideSheet 시트에서 E1:E2 범위 선택
2. **Ctrl + T** → 테이블 생성
3. 테이블 이름: `Path`

#### 7.1.2 TempPath 테이블 (H1:I2)

| H1 | I1 | 내용 |
|----|----|------|
| 이름 | 경로 | 헤더 |
| 결산폴더 | `/Shared Documents/★시연용폴더★/2512` | 데이터 |

**생성 방법**:
1. HideSheet 시트에서 H1:I2 범위 선택
2. **Ctrl + T** → 테이블 생성
3. 테이블 이름: `TempPath`

> **frmDirectory 연동**: 사용자가 폼에서 폴더 선택 시 `경로` 컬럼에 자동 저장

#### 7.1.3 결산연월 테이블 (A1:B2)

| A1 | B1 | 내용 |
|----|----|------|
| 결산연도 | 결산월 | 헤더 |
| 2025 | 12 | 데이터 |

**생성 방법**:
1. HideSheet 시트에서 A1:B2 범위 선택
2. **Ctrl + T** → 테이블 생성
3. 테이블 이름: `결산연월`

#### 7.1.4 People_Work 테이블

**VBA 참조**: `frmAddPerson.frm` Line 26, `frmScope.frm` Line 69

**위치**: HideSheet 시트 `Q1:S2`
**테이블명**: `People_Work`

| Q1 | R1 | S1 |
|----|----|----|
| 담당자 이름 | (추가 컬럼) | (추가 컬럼) |
| 홍길동 | ... | ... |

**생성 방법**:
1. Q1:S1에 헤더 입력
2. Q2부터 담당자 정보 입력
3. 범위 선택 → **Ctrl + T** → 테이블 생성
4. 테이블 이름: `People_Work`

> **중요**: frmScope에서 담당자 드롭다운 목록에 사용됩니다.

#### 7.1.5 Link 테이블

**용도**: Power Query가 채우는 법인별 SPO 파일 링크

**위치**: HideSheet 시트 `A5:B6`
**테이블명**: `Link`

| A5 | B5 |
|----|----|
| 법인코드 | Link |
| (Power Query) | (Power Query) |

#### 7.1.6 Link_취득_처분 테이블

**용도**: Power Query가 채우는 취득/처분 법인 파일 링크

**위치**: HideSheet 시트 `E5:F6`
**테이블명**: `Link_취득_처분`

| E5 | F5 |
|----|----|
| 법인코드 | Link |
| (Power Query) | (Power Query) |

> **주의**: 기존 이름이 `Corp_취득_처분`이면 `Link_취득_처분`으로 변경해야 VBA와 일치합니다.

---

### 7.2 CorpMaster 테이블들

#### 7.2.1 Mode 테이블 (신규 추가)

**VBA 참조**: `frmScope.frm` Line 39, 141, 147-149

| Mode |
|------|
| Full-Scope |

**위치**: CorpMaster 시트 (예: A1:A2)
**테이블명**: `Mode`

**값 옵션**:
- `Full-Scope`: 전체 법인 연결
- `간병반`: 간이 연결

#### 7.2.2 FS 테이블 (신규 추가)

**VBA 참조**: `frmScope.frm` Line 40, 143, 153-155

| FS |
|----|
| 별도 |

**위치**: CorpMaster 시트 (예: C1:C2)
**테이블명**: `FS`

**값 옵션**:
- `별도`: 별도재무제표 기준
- `연결`: 연결재무제표 기준

#### 7.2.3 Who 테이블 (신규 추가)

**VBA 참조**: `frmScope.frm` Line 41, 142, 165

| Who |
|-----|
| 전계 |

**위치**: CorpMaster 시트 (예: E1:E2)
**테이블명**: `Who`

**값**: 현재 담당자 이름 (frmScope에서 선택)

#### 7.2.4 Corp 테이블 (법인 현황)

**VBA 참조**: `CorpMaster_code.bas` Line 13-24, `mod_06_VerifySum.bas` Line 161, 381

**컬럼 구조 (15개)**:

| 인덱스 | 컬럼명 | 타입 | 예시 | 설명 |
|--------|--------|------|------|------|
| 1 | 법인코드 | text | 1000 | 고유 식별자 |
| 2 | 법인명 | text | 에이치알이 | 회사명 |
| 3 | 통화 | text | KRW | 기능통화 |
| 4 | 국가 | text | 한국 | 소재국 |
| 5 | 취득일 | date/text | 2020-01-01 또는 - | 연결 편입일 |
| 6 | 처분일 | date/text | - | 연결 제외일 |
| 7 | 지분율 | % | 100% | 지배기업 지분 |
| 8 | 연결방법 | text | 종속 | 종속/관계/공동 |
| 9 | 업종 | text | 지주회사 | 사업 유형 |
| 10 | 담당자 | text | 홍길동 | 법인 담당자 |
| 11 | Scope | text | O | O=포함, X=제외 |
| 12 | 비고 | text | - | 추가 메모 |
| 13 | 전기이월 | number | 0 | 이익잉여금 조정 |
| 14 | 당기이월 | number | 0 | 당기 조정 |
| 15 | Link | text | (URL) | SPO 파일 링크 |

**위치**: CorpMaster 시트 (예: A10:O50)
**테이블명**: `Corp`

---

### 7.3 AddCoA 시트 테이블

#### 7.3.1 CoA_Input 테이블 (신규 추가)

**VBA 참조**: `mod_03_PTB_CoA_Input.bas` Fill_Input_Table()

**컬럼 구조 (6개)**:

| 인덱스 | 컬럼명 | 타입 | 설명 |
|--------|--------|------|------|
| 1 | 법인코드 | text | PTB에서 복사 |
| 2 | 법인명 | text | PTB에서 복사 |
| 3 | 법인별CoA | text | PTB에서 복사 |
| 4 | 법인별계정과목명 | text | PTB에서 복사 |
| 5 | 연결CoA | text | 사용자 입력/VBA 제안 |
| 6 | 연결계정명 | text | VBA 자동 채움 |

**위치**: AddCoA 시트 B4:G4 (헤더만)
**테이블명**: `CoA_Input`

> **중요**: VBA `Fill_Input_Table()`이 PTB에서 연결CoA가 비어있는 행을 이 테이블로 복사합니다.

---

### 7.4 ADBS 시트 테이블

#### 7.4.1 AD_BS 테이블 (신규 추가)

**용도**: 취득/처분 시점 B/S 데이터

**컬럼 구조**: PTB와 동일 (9컬럼)

**위치**: ADBS 시트 B4:J4
**테이블명**: `AD_BS`

---

## 8. 쿼리 생성 순서

### 8.1 생성 순서 (의존성 기반)

```
1. fnProcessSingleFile (함수, 연결 전용)
   └── 의존: 없음

2. 메인주소 (텍스트 값, 연결 전용)
   └── 의존: HideSheet.Path 테이블

3. 회사결산_주소 (텍스트 값, 연결 전용)
   └── 의존: HideSheet.TempPath 테이블

4. 디렉터리 (테이블, DirectoryURL 시트)
   └── 의존: 메인주소

5. Raw_Data (테이블, 연결 전용)
   └── 의존: 메인주소, 회사결산_주소

6. TB (테이블, 연결 전용)
   └── 의존: Raw_Data, fnProcessSingleFile

7. PTB (테이블, BSPL 시트)
   └── 의존: TB

8. Link (테이블, HideSheet)
   └── 의존: 메인주소, 회사결산_주소, Corp 테이블

9. Link_취득_처분 (테이블, HideSheet)
   └── 의존: 메인주소, Corp 테이블
```

### 8.2 로드 설정 방법

**연결 전용 설정**:
1. Power Query 편집기에서 쿼리 선택
2. **홈** → **닫기 및 로드** → **닫기 및 로드 위치...**
3. **연결만 만들기** 선택
4. **확인**

**테이블 로드 설정**:
1. 쿼리 선택
2. **홈** → **닫기 및 로드** → **닫기 및 로드 위치...**
3. **테이블** 선택
4. **기존 워크시트**: 셀 주소 지정 (예: `BSPL!$B$4`)
5. **확인**

---

## 9. VBA 컬럼 인덱스 매핑

### 9.1 mod_03_PTB_CoA_Input.bas 수정 사항

| 위치 | 기존 | 변경 | 설명 |
|------|------|------|------|
| Line 46 | `.Resize(, 5)` | `.Resize(, 6)` | PTB 복사 범위 확대 |
| Line 114 | `coaRow.Range(1, 3)` | `coaRow.Range(1, 4)` | 법인별CoA |
| Line 144 | `coaRow.Range(1, 4)` | `coaRow.Range(1, 5)` | 연결CoA |
| Line 145 | `coaRow.Range(1, 5)` | `coaRow.Range(1, 6)` | 연결계정명 |

### 9.2 mod_05_PTB_Highlight.bas 수정 사항

| 위치 | 기존 | 변경 | 설명 |
|------|------|------|------|
| Line 80 | `rng.Cells(i, 4)` | `rng.Cells(i, 5)` | 연결CoA 체크 컬럼 |

---

## 10. 테스트 체크리스트

### 10.1 쿼리별 테스트

- [ ] **메인주소**: HideSheet.Path 값 정상 반환
- [ ] **회사결산_주소**: HideSheet.TempPath 값 정상 반환
- [ ] **디렉터리**: DirectoryURL 시트에 폴더 구조 로드
- [ ] **Raw_Data**: SPO 시산표 파일 목록 표시 (법인명_폴더 포함)
- [ ] **TB**: 모든 법인 데이터 병합 (법인코드, 법인명 정상 추출)
- [ ] **PTB**: BSPL 시트 B4부터 9컬럼 테이블 로드
- [ ] **Link**: HideSheet에 법인코드, Link 컬럼 로드
- [ ] **Link_취득_처분**: HideSheet에 취득/처분 법인 링크 로드

### 10.2 테이블 검증

- [ ] **Path**: 테이블 존재, URL 값 확인
- [ ] **TempPath**: 테이블 존재, 경로 값 확인
- [ ] **People_Work**: 테이블 존재, 담당자 목록 확인
- [ ] **Mode/FS/Who**: CorpMaster에 테이블 존재
- [ ] **Corp**: 15개 컬럼 구조 확인
- [ ] **CoA_Input**: AddCoA 시트에 테이블 존재 (빈 상태)
- [ ] **AD_BS**: ADBS 시트에 테이블 존재

### 10.3 데이터 검증

- [ ] 법인코드: 메타데이터 점 이전 값과 일치
- [ ] 법인명: 메타데이터 점 이후 값 (없으면 폴더명)
- [ ] 법인별CoA: 7자리 숫자
- [ ] 당기: 차변 - 대변 계산 정확

### 10.4 VBA 연동 테스트

- [ ] **QueryRefresh()**: Power Query 새로고침 정상 동작
- [ ] **Fill_Input_Table()**: CoA_Input 테이블에 6컬럼 복사
- [ ] **HighlightPTB()**: 연결CoA 비어있는 행 노란색 하이라이트
- [ ] **frmScope**: Link, Link_취득_처분 쿼리 새로고침

---

## 11. 문제 해결

### 11.1 SharePoint 인증 오류

**증상**: "인증이 필요합니다" 오류

**해결**:
1. Power Query 편집기 → **홈** → **데이터 원본 설정**
2. SharePoint URL 선택 → **권한 편집**
3. **조직 계정** → **로그인**
4. Office 365 계정으로 인증

### 11.2 경로 오류

**증상**: Raw_Data 결과가 비어있음

**해결**:
1. HideSheet.TempPath 값 확인
2. SharePoint에서 경로 정확히 복사
3. 경로 구분자 `/` 확인 (역슬래시 `\` 아님)
4. 공백 문자 확인 (`/Shared Documents/` vs `/SharedDocuments/`)

### 11.3 법인코드 추출 실패

**증상**: 법인코드가 null 또는 Error

**해결**:
1. 시산표 파일 **Row 4 (A4 셀)** 확인 ← 주의: A1이 아님!
2. 형식: `회 사 : 1000.에이치알이` (콜론, 점 필수)
3. Column1에 메타데이터가 있어야 함

**실제 파일 구조 확인**:
```
Row 1: null | null | "합계잔액시산표"
Row 4: "회 사 : 1000.에이치알이주..." ← 여기서 추출
Row 9+: 데이터 시작
```

### 11.4 테이블 누락 오류

**증상**: "Subscript out of range" 런타임 오류 9

**해결**:
1. 오류 발생 VBA 코드에서 참조하는 테이블명 확인
2. Excel에서 해당 테이블 존재 여부 확인
3. 테이블 이름 정확히 일치하는지 확인 (대소문자 구분)

---

## 12. 참조

- **HRE_SPO_M함수_전체.md**: 단일 파일 처리 M 함수 상세
- **완벽한_구현_가이드.md**: 섹션 10 Power Query 연결 설정
- **mod_03_PTB_CoA_Input.bas**: VBA CoA 매핑 로직
- **mod_05_PTB_Highlight.bas**: VBA 하이라이트 로직
- **mod_06_VerifySum.bas**: Link 테이블 참조 로직
- **frmScope.frm**: Mode, FS, Who, Link 쿼리 새로고침

---

## 13. 변경 이력

| 버전 | 날짜 | 변경 내용 |
|------|------|-----------|
| 1.00 | 2026-01-24 | 초기 버전 |
| 1.10 | 2026-01-25 | 누락 항목 추가: Link 쿼리, Link_취득_처분 쿼리, Mode/FS/Who/People_Work/CoA_Input/AD_BS 테이블 |

---

**작성일**: 2026-01-25
**버전**: 1.10
