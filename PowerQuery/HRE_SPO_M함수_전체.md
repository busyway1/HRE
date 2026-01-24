# HRE 단일 시산표 파일 처리 M 함수

**버전**: 1.00
**작성일**: 2026-01-24
**용도**: 개별 시산표 Excel 파일에서 법인코드, 법인명, CoA, 금액 추출

---

## 1. 개요

### 1.1 파일 역할

이 문서는 **단일 시산표 파일**을 처리하는 M 함수를 정의합니다.

```
개별 시산표 파일 (*.xlsx)
    ↓
[단일 파일 M 함수]
    ↓
법인코드, 법인명, 법인별CoA, 법인별계정과목명, 당기
```

### 1.2 다른 문서와의 관계

| 문서 | 역할 | 이 문서와의 관계 |
|------|------|------------------|
| `HRE_PowerQuery_전체.md` | SPO 전체 구조 | **TB 쿼리에서 이 함수 호출** |
| `완벽한_구현_가이드.md` | 전체 매뉴얼 | 섹션 10에서 이 문서 참조 |

### 1.3 BEP vs HRE 차이점

| 항목 | BEP | HRE |
|------|-----|-----|
| 법인코드 추출 | 파일명 패턴 `[HRE-001]` | **메타데이터 점 이전**: `1000.에이치알이` → `1000` |
| 법인명 추출 | 없음 | **메타데이터 점 이후**: `1000.에이치알이` → `에이치알이` |
| 출력 컬럼 | 법인코드, CoA, 금액 | 법인코드, **법인명**, CoA, 금액 |

---

## 2. 데이터 추출 위치

### 2.1 시산표 파일 구조 (예시)

```
┌───────────────────────────────────────────────────┐
│  A1: "회사: 1000.에이치알이주식회사"    ← 메타셀   │
│  A2: "기간: 2025-12-01 ~ 2025-12-31"              │
│  A3: (빈 행)                                      │
│  A4: [계정코드] | [계정과목명] | [차변] | [대변]   │
│  A5: 1030000   | 보통예금     | 100,000 | 0       │
│  A6: 1110000   | 외상매출금   | 50,000  | 0       │
│  ...                                              │
└───────────────────────────────────────────────────┘
```

### 2.2 추출 데이터 매핑

| 추출 데이터 | 소스 위치 | 추출 로직 | 예시 |
|------------|----------|----------|------|
| **법인코드** | A1 메타셀 (점 이전) | `회사: 1000.xxx` → `1000` | `1000` |
| **법인명** | A1 메타셀 (점 이후) | `회사: 1000.에이치알이` → `에이치알이` | `에이치알이주식회사` |
| **법인별CoA** | 계정코드 컬럼 | 대괄호 제거 + 7자리 추출 | `[1030000]` → `1030000` |
| **법인별계정과목명** | 계정과목명 컬럼 | 그대로 사용 | `보통예금` |
| **당기** | 차변 - 대변 계산 | 차변잔액 - 대변잔액 | `100,000` |

---

## 3. 핵심 M 함수

### 3.1 정규화 함수 (공백 및 null 처리)

```m
// 정규화 함수 - 공백 제거 및 null 처리
정규화 = (val) =>
    let
        str = if val = null then "" else Text.From(val),
        cleaned = Text.Trim(Text.Clean(str))
    in
        cleaned,
```

### 3.2 법인코드 추출 (메타데이터 점 이전)

```m
// 법인코드 추출: "회사: 1000.에이치알이주식회사" → "1000"
법인코드_추출 = (메타셀 as any) =>
    let
        t0 = 정규화(메타셀),                                    // "회사:1000.에이치알이주식회사"
        t1 = if t0 = "" then null
             else Text.AfterDelimiter(t0, ":"),                 // "1000.에이치알이주식회사"
        t2 = if t1 = null then null
             else Text.BeforeDelimiter(t1, ".")                 // "1000"
    in
        t2,
```

### 3.3 법인명 추출 (메타데이터 점 이후)

```m
// 법인명 추출: "회사: 1000.에이치알이주식회사" → "에이치알이주식회사"
법인명_추출 = (메타셀 as any) =>
    let
        t0 = 정규화(메타셀),
        t1 = if t0 = "" then null
             else Text.AfterDelimiter(t0, ":"),                 // "1000.에이치알이주식회사"
        name = if t1 = null then null
               else Text.AfterDelimiter(t1, ".")                // "에이치알이주식회사"
    in
        name,
```

### 3.4 법인별 CoA 파싱

```m
// 법인별 CoA 파싱: "[1030000]보통예금" → "1030000"
// 또는: "1030000" → "1030000" (대괄호 없는 경우)
법인별CoA_파싱 = (rawCode as any) =>
    let
        str = 정규화(rawCode),
        // 대괄호 패턴 확인
        hasBracket = Text.Contains(str, "["),
        // 대괄호 내용 추출 또는 7자리 추출
        code = if hasBracket then
                   let
                       afterOpen = Text.AfterDelimiter(str, "["),
                       beforeClose = Text.BeforeDelimiter(afterOpen, "]")
                   in beforeClose
               else if Text.Length(str) >= 7 then
                   Text.Start(str, 7)
               else
                   str
    in
        code,
```

### 3.5 계정과목명 파싱

```m
// 계정과목명 파싱: "[1030000]보통예금" → "보통예금"
// 또는: "보통예금" → "보통예금" (대괄호 없는 경우)
계정과목명_파싱 = (rawName as any) =>
    let
        str = 정규화(rawName),
        hasBracket = Text.Contains(str, "]"),
        name = if hasBracket then
                   Text.AfterDelimiter(str, "]")
               else
                   str
    in
        name,
```

### 3.6 당기 계산 (차변 - 대변)

```m
// 당기 계산: 차변잔액 - 대변잔액
당기_계산 = (차변 as any, 대변 as any) =>
    let
        debit = if 차변 = null then 0 else Number.From(차변),
        credit = if 대변 = null then 0 else Number.From(대변),
        net = debit - credit
    in
        net,
```

---

## 4. 완전한 단일 파일 처리 함수

### 4.1 fnProcessSingleFile 함수

이 함수는 `HRE_PowerQuery_전체.md`의 **TB 쿼리에서 호출**됩니다.

```m
// ============================================================
// fnProcessSingleFile - 단일 시산표 파일 처리 함수
//
// 입력:
//   - fileContent: 시산표 Excel 파일 Binary
//   - 법인명_폴더: Raw_Data에서 전달받은 폴더명 (fallback용)
//
// 출력:
//   - Table: {법인코드, 법인명, 법인별CoA, 법인별계정과목명, 당기}
// ============================================================

let
    fnProcessSingleFile = (fileContent as binary, 법인명_폴더 as text) =>
    let
        // ========== 1. Excel 파일 로드 ==========
        Source = Excel.Workbook(fileContent, null, true),

        // 첫 번째 시트 선택 (시산표 시트)
        FirstSheet = Source{0}[Data],

        // ========== 2. 메타데이터 추출 (A1 셀) ==========
        메타셀 = FirstSheet{0}[Column1],

        // ========== 3. 법인코드/법인명 추출 ==========
        법인코드 = 법인코드_추출(메타셀),
        법인명_메타 = 법인명_추출(메타셀),

        // 법인명 결정: 메타데이터 우선, 없으면 폴더명
        법인명_최종 = if 법인명_메타 <> null and 법인명_메타 <> ""
                     then 법인명_메타
                     else 법인명_폴더,

        // ========== 4. 데이터 행 추출 (헤더 스킵) ==========
        // 첫 4행은 메타데이터/헤더, 5행부터 데이터
        DataRows = Table.Skip(FirstSheet, 4),
        PromotedHeaders = Table.PromoteHeaders(DataRows, [PromoteAllScalars=true]),

        // ========== 5. 필수 컬럼 선택 ==========
        SelectedColumns = Table.SelectColumns(PromotedHeaders,
            {"계정코드", "계정과목명", "차변잔액", "대변잔액"},
            MissingField.UseNull),

        // ========== 6. 데이터 변환 ==========
        TransformedData = Table.AddColumn(
            Table.AddColumn(
                Table.AddColumn(
                    Table.AddColumn(
                        Table.AddColumn(
                            SelectedColumns,
                            "법인코드", each 법인코드, type text
                        ),
                        "법인명", each 법인명_최종, type text
                    ),
                    "법인별CoA", each 법인별CoA_파싱([계정코드]), type text
                ),
                "법인별계정과목명", each 계정과목명_파싱([계정과목명]), type text
            ),
            "당기", each 당기_계산([차변잔액], [대변잔액]), type number
        ),

        // ========== 7. 최종 컬럼 선택 ==========
        FinalColumns = Table.SelectColumns(TransformedData,
            {"법인코드", "법인명", "법인별CoA", "법인별계정과목명", "당기"}
        ),

        // ========== 8. 빈 행 제거 ==========
        FilteredRows = Table.SelectRows(FinalColumns,
            each [법인별CoA] <> null and [법인별CoA] <> ""
        )
    in
        FilteredRows
in
    fnProcessSingleFile
```

### 4.2 함수 출력 스키마

| 컬럼 | 타입 | 소스 | 예시 |
|------|------|------|------|
| 법인코드 | text | 메타셀 점 이전 | `1000` |
| 법인명 | text | 메타셀 점 이후 (또는 폴더명) | `에이치알이주식회사` |
| 법인별CoA | text | 계정코드 컬럼 파싱 | `1030000` |
| 법인별계정과목명 | text | 계정과목명 컬럼 파싱 | `보통예금` |
| 당기 | number | 차변잔액 - 대변잔액 | `100000` |

---

## 5. 식별코드 생성 (선택사항)

PTB 쿼리에서 중복 검사용 식별코드 생성:

```m
// 식별코드 생성: 법인코드_법인별CoA
식별코드_생성 = (법인코드 as text, CoA as text) =>
    let
        id = 법인코드 & "_" & CoA
    in
        id,
```

**용도**: CoA 매핑 이력 조회 시 복합키로 사용

---

## 6. 예외 처리

### 6.1 메타셀 형식 오류

```m
// 안전한 법인코드 추출 (예외 처리 포함)
법인코드_추출_안전 = (메타셀 as any) =>
    let
        result = try 법인코드_추출(메타셀) otherwise null
    in
        result,
```

### 6.2 컬럼 누락 대응

```m
// 컬럼 존재 여부 확인
컬럼_존재 = (tbl as table, colName as text) =>
    List.Contains(Table.ColumnNames(tbl), colName),

// 안전한 컬럼 선택
안전한_컬럼_선택 = (tbl as table, colNames as list) =>
    let
        existingCols = List.Select(colNames, each 컬럼_존재(tbl, _)),
        selected = Table.SelectColumns(tbl, existingCols, MissingField.UseNull)
    in
        selected,
```

---

## 7. 테스트 가이드

### 7.1 로컬 파일 테스트

```m
let
    // 테스트 파일 경로
    TestFilePath = "C:\HRE\시산표_HRE_202512.xlsx",
    TestContent = File.Contents(TestFilePath),

    // 함수 호출
    Result = fnProcessSingleFile(TestContent, "HRE"),

    // 결과 확인
    RowCount = Table.RowCount(Result),
    FirstRow = Result{0}
in
    Result
```

### 7.2 예상 출력

```
법인코드    법인명              법인별CoA   법인별계정과목명    당기
--------    ----------------    ---------   ----------------    --------
1000        에이치알이주식회사   1030000     보통예금            100,000
1000        에이치알이주식회사   1110000     외상매출금          50,000
1000        에이치알이주식회사   2010000     외상매입금          -30,000
...
```

---

## 8. 참조

- **HRE_PowerQuery_전체.md**: TB 쿼리에서 이 함수 호출
- **완벽한_구현_가이드.md**: 섹션 10 Power Query 연결 설정
- **mod_03_PTB_CoA_Input.bas**: VBA에서 PTB 데이터 처리

---

**작성일**: 2026-01-24
**버전**: 1.00
