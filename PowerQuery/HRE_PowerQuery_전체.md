# HRE SharePoint Online 전체 구조 Power Query 가이드

**버전**: 1.00
**작성일**: 2026-01-24
**용도**: SPO 다중 법인 폴더 순회 및 PTB 테이블 생성

---

## 1. 개요

### 1.1 쿼리 아키텍처

```
[Phase 1: 설정 쿼리 - 연결 전용]
  ├── 메인주소 (Path 테이블 참조)
  └── 회사결산_주소 (TempPath 테이블 참조)

[Phase 2: 데이터 수집 - 연결 전용]
  └── Raw_Data
        ├── 의존: 메인주소, 회사결산_주소
        └── 출력: 법인명_폴더, Name, Content, Folder Path

[Phase 3: 데이터 변환 - 연결 전용]
  └── TB
        ├── 의존: Raw_Data, fnProcessSingleFile
        └── 출력: 법인코드, 법인명, 법인별CoA, 법인별계정과목명, 당기

[Phase 4: 최종 출력 - 테이블 로드]
  └── PTB
        ├── 의존: TB, Raw_CoA(선택), Master
        ├── 로드 위치: BSPL 시트 B4
        └── 테이블명: PTB
```

### 1.2 쿼리 개수 및 로드 설정

| 쿼리명 | Phase | 로드 설정 | 출력 위치 |
|--------|-------|----------|----------|
| 메인주소 | 1 | 연결 전용 | - |
| 회사결산_주소 | 1 | 연결 전용 | - |
| Raw_Data | 2 | 연결 전용 | - |
| fnProcessSingleFile | 3 | 연결 전용 (함수) | - |
| TB | 3 | 연결 전용 | - |
| PTB | 4 | **테이블 로드** | BSPL!$B$4 |

### 1.3 HideSheet 테이블 참조

| 테이블명 | 위치 | 용도 | 초기값 |
|----------|------|------|--------|
| Path | HideSheet E1:E2 | SPO 메인 URL | `https://pwckor.sharepoint.com/sites/KR-ASR-HRE_Consolidation` |
| TempPath | HideSheet H1:H2 | 회사결산 폴더 경로 | `/Shared Documents/★시연용폴더★/2512` |

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

    // 첫 번째 행 가져오기
    FirstRow = Source{0},

    // 컬럼명에 상관없이 첫 번째 값 추출 (유연성 확보)
    PathValue = Record.FieldValues(FirstRow){0},

    // 경로 정리
    CleanedPath = Text.Trim(Text.From(PathValue))
in
    CleanedPath
```

> **참고**: `Record.FieldValues(FirstRow){0}` 사용으로 컬럼명이 "경로"가 아니어도 동작합니다.

**출력 예시**: `/Shared Documents/★시연용폴더★/2512`

---

## 3. Phase 2: 데이터 수집

### 3.1 Raw_Data 쿼리

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

### 4.1 fnProcessSingleFile 함수 (참조)

**상세 내용**: `HRE_SPO_M함수_전체.md` 참조

```m
// 쿼리명: fnProcessSingleFile
// 로드: 연결 전용 (함수)
// 용도: 단일 시산표 파일에서 데이터 추출

let
    // 헬퍼 함수들
    정규화 = (val) =>
        let
            str = if val = null then "" else Text.From(val),
            cleaned = Text.Trim(Text.Clean(str))
        in cleaned,

    법인코드_추출 = (메타셀) =>
        let
            t0 = 정규화(메타셀),
            t1 = if t0 = "" then null else Text.AfterDelimiter(t0, ":"),
            t2 = if t1 = null then null else Text.BeforeDelimiter(t1, ".")
        in t2,

    법인명_추출 = (메타셀) =>
        let
            t0 = 정규화(메타셀),
            t1 = if t0 = "" then null else Text.AfterDelimiter(t0, ":"),
            name = if t1 = null then null else Text.AfterDelimiter(t1, ".")
        in name,

    법인별CoA_파싱 = (rawCode) =>
        let
            str = 정규화(rawCode),
            hasBracket = Text.Contains(str, "["),
            code = if hasBracket then
                       Text.BeforeDelimiter(Text.AfterDelimiter(str, "["), "]")
                   else if Text.Length(str) >= 7 then
                       Text.Start(str, 7)
                   else str
        in code,

    계정과목명_파싱 = (rawName) =>
        let
            str = 정규화(rawName),
            hasBracket = Text.Contains(str, "]"),
            name = if hasBracket then Text.AfterDelimiter(str, "]") else str
        in name,

    당기_계산 = (차변, 대변) =>
        let
            debit = if 차변 = null then 0 else Number.From(차변),
            credit = if 대변 = null then 0 else Number.From(대변),
            net = debit - credit
        in net,

    // 메인 함수
    fnProcessSingleFile = (fileContent as binary, 법인명_폴더 as text) =>
        let
            Source = Excel.Workbook(fileContent, null, true),
            FirstSheet = Source{0}[Data],

            메타셀 = FirstSheet{0}[Column1],
            법인코드 = 법인코드_추출(메타셀),
            법인명_메타 = 법인명_추출(메타셀),
            법인명_최종 = if 법인명_메타 <> null and 법인명_메타 <> ""
                         then 법인명_메타 else 법인명_폴더,

            DataRows = Table.Skip(FirstSheet, 4),
            PromotedHeaders = Table.PromoteHeaders(DataRows, [PromoteAllScalars=true]),

            // 컬럼명 자동 매핑 (유연성)
            colNames = Table.ColumnNames(PromotedHeaders),
            차변컬럼 = List.First(List.Select(colNames, each Text.Contains(_, "차변")), "차변잔액"),
            대변컬럼 = List.First(List.Select(colNames, each Text.Contains(_, "대변")), "대변잔액"),
            계정코드컬럼 = List.First(List.Select(colNames, each Text.Contains(_, "계정코드") or Text.Contains(_, "코드")), "계정코드"),
            계정명컬럼 = List.First(List.Select(colNames, each Text.Contains(_, "계정과목") or Text.Contains(_, "계정명")), "계정과목명"),

            AddCols = Table.AddColumn(
                Table.AddColumn(
                    Table.AddColumn(
                        Table.AddColumn(
                            Table.AddColumn(
                                PromotedHeaders,
                                "법인코드", each 법인코드, type text
                            ),
                            "법인명", each 법인명_최종, type text
                        ),
                        "법인별CoA", each try 법인별CoA_파싱(Record.Field(_, 계정코드컬럼)) otherwise null, type text
                    ),
                    "법인별계정과목명", each try 계정과목명_파싱(Record.Field(_, 계정명컬럼)) otherwise null, type text
                ),
                "당기", each try 당기_계산(Record.Field(_, 차변컬럼), Record.Field(_, 대변컬럼)) otherwise 0, type number
            ),

            FinalColumns = Table.SelectColumns(AddCols,
                {"법인코드", "법인명", "법인별CoA", "법인별계정과목명", "당기"}
            ),

            FilteredRows = Table.SelectRows(FinalColumns,
                each [법인별CoA] <> null and [법인별CoA] <> ""
            )
        in
            FilteredRows
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
    AddPwC_CoA = Table.AddColumn(Source, "PwC_CoA", each null, type text),
    AddPwC_계정명 = Table.AddColumn(AddPwC_CoA, "PwC_계정명", each null, type text),

    // ========== 3. 분류 컬럼 추가 (VBA가 채움) ==========
    Add대분류 = Table.AddColumn(AddPwC_계정명, "대분류", each null, type text),
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
        "PwC_CoA",          // 5
        "PwC_계정명",       // 6
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
| 5 | PwC_CoA | text | VBA | `Range(1, 5)` |
| 6 | PwC_계정명 | text | VBA | `Range(1, 6)` |
| 7 | 대분류 | text | VBA | `Range(1, 7)` |
| 8 | 중분류 | text | VBA | `Range(1, 8)` |
| 9 | 금액 | number | TB.당기 | `Range(1, 9)` |

---

## 6. 쿼리 생성 순서

### 6.1 생성 순서 (의존성 기반)

```
1. fnProcessSingleFile (함수, 연결 전용)
   └── 의존: 없음

2. 메인주소 (텍스트 값, 연결 전용)
   └── 의존: HideSheet.Path 테이블

3. 회사결산_주소 (텍스트 값, 연결 전용)
   └── 의존: HideSheet.TempPath 테이블

4. Raw_Data (테이블, 연결 전용)
   └── 의존: 메인주소, 회사결산_주소

5. TB (테이블, 연결 전용)
   └── 의존: Raw_Data, fnProcessSingleFile

6. PTB (테이블, 테이블 로드)
   └── 의존: TB
```

### 6.2 로드 설정 방법

**연결 전용 설정**:
1. Power Query 편집기에서 쿼리 선택
2. **홈** → **닫기 및 로드** → **닫기 및 로드 위치...**
3. **연결만 만들기** 선택
4. **확인**

**테이블 로드 설정 (PTB만)**:
1. PTB 쿼리 선택
2. **홈** → **닫기 및 로드** → **닫기 및 로드 위치...**
3. **테이블** 선택
4. **기존 워크시트**: `BSPL!$B$4`
5. **확인**

---

## 7. HideSheet 테이블 설정

### 7.1 Path 테이블 (E1:E2)

| E1 | 내용 |
|----|------|
| 헤더 | `Path` |
| 데이터 | `https://pwckor.sharepoint.com/sites/KR-ASR-HRE_Consolidation` |

**생성 방법**:
1. HideSheet 시트에서 E1:E2 범위 선택
2. **Ctrl + T** → 테이블 생성
3. 테이블 이름: `Path`

### 7.2 TempPath 테이블 (H1:H2)

| H1 | 내용 |
|----|------|
| 헤더 | `경로` |
| 데이터 | `/Shared Documents/★시연용폴더★/2512` |

**생성 방법**:
1. HideSheet 시트에서 H1:H2 범위 선택
2. **Ctrl + T** → 테이블 생성
3. 테이블 이름: `TempPath`

**경로 설정 방법**:
- **frmDirectory 폼**에서 사용자가 SPO 폴더 선택 시 자동 업데이트
- 또는 직접 H2 셀에 경로 입력

---

## 8. VBA 컬럼 인덱스 매핑

### 8.1 mod_03_PTB_CoA_Input.bas 수정 사항

| 위치 | 기존 | 변경 | 설명 |
|------|------|------|------|
| Line 46 | `.Resize(, 5)` | `.Resize(, 6)` | PTB 복사 범위 확대 |
| Line 114 | `coaRow.Range(1, 3)` | `coaRow.Range(1, 4)` | 법인별CoA |
| Line 144 | `coaRow.Range(1, 4)` | `coaRow.Range(1, 5)` | PwC_CoA |
| Line 145 | `coaRow.Range(1, 5)` | `coaRow.Range(1, 6)` | PwC_계정명 |

### 8.2 mod_05_PTB_Highlight.bas 수정 사항

| 위치 | 기존 | 변경 | 설명 |
|------|------|------|------|
| Line 80 | `rng.Cells(i, 4)` | `rng.Cells(i, 5)` | PwC_CoA 체크 컬럼 |

---

## 9. 테스트 체크리스트

### 9.1 쿼리별 테스트

- [ ] **메인주소**: HideSheet.Path 값 정상 반환
- [ ] **회사결산_주소**: HideSheet.TempPath 값 정상 반환
- [ ] **Raw_Data**: SPO 시산표 파일 목록 표시 (법인명_폴더 포함)
- [ ] **TB**: 모든 법인 데이터 병합 (법인코드, 법인명 정상 추출)
- [ ] **PTB**: BSPL 시트 B4부터 9컬럼 테이블 로드

### 9.2 데이터 검증

- [ ] 법인코드: 메타데이터 점 이전 값과 일치
- [ ] 법인명: 메타데이터 점 이후 값 (없으면 폴더명)
- [ ] 법인별CoA: 7자리 숫자
- [ ] 당기: 차변 - 대변 계산 정확

### 9.3 VBA 연동 테스트

- [ ] **QueryRefresh()**: Power Query 새로고침 정상 동작
- [ ] **Fill_Input_Table()**: CoA_Input 테이블에 6컬럼 복사
- [ ] **HighlightPTB()**: PwC_CoA 비어있는 행 노란색 하이라이트

---

## 10. 문제 해결

### 10.1 SharePoint 인증 오류

**증상**: "인증이 필요합니다" 오류

**해결**:
1. Power Query 편집기 → **홈** → **데이터 원본 설정**
2. SharePoint URL 선택 → **권한 편집**
3. **조직 계정** → **로그인**
4. Office 365 계정으로 인증

### 10.2 경로 오류

**증상**: Raw_Data 결과가 비어있음

**해결**:
1. HideSheet.TempPath 값 확인
2. SharePoint에서 경로 정확히 복사
3. 경로 구분자 `/` 확인 (역슬래시 `\` 아님)

### 10.3 법인코드 추출 실패

**증상**: 법인코드가 null

**해결**:
1. 시산표 파일 A1 셀 확인
2. 형식: `회사: 1000.에이치알이` (콜론, 점 필수)
3. 공백 없이 입력

---

## 11. 참조

- **HRE_SPO_M함수_전체.md**: 단일 파일 처리 M 함수 상세
- **완벽한_구현_가이드.md**: 섹션 10 Power Query 연결 설정
- **mod_03_PTB_CoA_Input.bas**: VBA CoA 매핑 로직
- **mod_05_PTB_Highlight.bas**: VBA 하이라이트 로직

---

**작성일**: 2026-01-24
**버전**: 1.00
