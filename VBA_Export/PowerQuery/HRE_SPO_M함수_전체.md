# HRE SPO Power Query M 함수 (전체 법인 처리)

## 1. 워크플로우 분석

### 1.1 데이터 흐름

```
[HideSheet 테이블]
       │
       ├─→ Path 테이블 (메인주소)
       │       │
       │       └─→ 메인주소 쿼리
       │
       └─→ TempPath 테이블 (경로: /★시연용폴더★/2512)
               │
               └─→ 회사결산 주소 쿼리
                       │
                       └─→ Raw_Data 쿼리 (SPO 파일 목록)
                               │
                               ├─→ 법인명: 폴더명에서 추출 (HRE, 예천솔라)
                               ├─→ 시산표 파일 필터링
                               └─→ 파일 수 검증 (법인당 1개)
                                       │
                                       └─→ TB 쿼리 (파일 파싱)
                                               │
                                               ├─→ 법인코드: 파일 내 메타데이터에서 추출 (1000)
                                               ├─→ 계정과목 파싱
                                               └─→ 당기 계산
                                                       │
                                                       └─→ PTB 쿼리 (CoA 매핑)
```

### 1.2 SPO 폴더 구조

```
★시연용폴더★/
└── 2512/                        ← TempPath.경로에 저장
    ├── HRE/                     ← 법인명 (폴더명)
    │   └── *시산표*.xlsx        ← 시산표 파일
    └── 예천솔라/                ← 법인명 (폴더명)
        └── *시산표*.xlsx
```

### 1.3 컬럼 매핑

| 소스 | 추출 위치 | 예시 |
|------|----------|------|
| 법인명 | 폴더명 | `HRE`, `예천솔라` |
| 법인코드 | 파일 내 메타데이터 (회사:/회계단위:) | `1000` |
| 법인별 CoA | 파일 내 계정과목 파싱 | `1030000` |
| 법인별 계정과목명 | 파일 내 계정과목 파싱 | `보통예금` |
| 당기 | 차변잔액 - 대변잔액 | `1854605784` |

---

## 2. 기본 설정 쿼리

### 2.1 메인주소

```m
let
    경로 = Excel.CurrentWorkbook(){[Name="Path"]}[Content],
    메인주소 = 경로{0}[메인주소]
in
    메인주소
```

### 2.2 회사결산 주소

```m
let
    경로 = Excel.CurrentWorkbook(){[Name="TempPath"]}[Content],
    경로텍스트 = Text.From(경로{0}[경로])
in
    경로텍스트
```

---

## 3. Raw_Data 쿼리 (SPO 파일 목록)

```m
let
    // ========== 설정 ==========
    기준경로 = 회사결산_주소,  // 예: "/★시연용폴더★/2512"

    // ========== SharePoint 파일 목록 ==========
    원본 = SharePoint.Files(메인주소, [ApiVersion = 15]),

    // ========== 경로 필터링 ==========
    경로필터 = Table.SelectRows(원본, each Text.Contains([Folder Path], 기준경로)),

    // ========== 시산표 파일만 필터 ==========
    시산표필터 = Table.SelectRows(경로필터, each Text.Contains([Name], "시산표")),

    // ========== 법인명 추출 (폴더명) ==========
    // 경로 예: .../★시연용폴더★/2512/HRE/
    법인명추출 = Table.AddColumn(시산표필터, "법인명", each
        let
            fullPath = [Folder Path],
            // 기준경로 이후 부분 추출
            afterBase = Text.AfterDelimiter(fullPath, 기준경로),
            // "/" 로 분할하여 첫 번째 비어있지 않은 세그먼트 = 법인명
            segments = Text.Split(afterBase, "/"),
            cleanSegments = List.Select(segments, each _ <> null and _ <> ""),
            법인명 = if List.Count(cleanSegments) > 0 then cleanSegments{0} else null
        in
            법인명,
        type text
    ),

    // ========== 법인명이 있는 것만 유지 ==========
    법인명필터 = Table.SelectRows(법인명추출, each [법인명] <> null),

    // ========== 법인당 파일 수 검증 ==========
    // 그룹화하여 파일 수 확인
    법인별그룹 = Table.Group(법인명필터, {"법인명"}, {
        {"파일수", each Table.RowCount(_), Int64.Type},
        {"파일목록", each _, type table}
    }),

    // 파일이 2개 이상인 법인 체크
    다중파일법인 = Table.SelectRows(법인별그룹, each [파일수] > 1),
    다중파일체크 = if Table.RowCount(다중파일법인) > 0 then
        error "시산표 파일이 여러 개인 법인이 있습니다: " &
              Text.Combine(다중파일법인[법인명], ", ") &
              " - 각 법인 폴더에 시산표 파일은 1개만 있어야 합니다."
    else
        법인명필터,

    // ========== 필요 컬럼만 선택 ==========
    최종결과 = Table.SelectColumns(다중파일체크, {"법인명", "Name", "Content", "Folder Path"})
in
    최종결과
```

---

## 4. TB 쿼리 (파일 파싱 - 핵심)

```m
let
    // ========== Raw_Data에서 파일 목록 가져오기 ==========
    파일목록 = Raw_Data,

    // ========== 공통 함수 정의 ==========
    // 공백 제거 함수
    정규화 = (x as any) as text =>
        if x = null then "" else Text.Replace(Text.From(x), " ", ""),

    // ========== 각 파일 처리 함수 ==========
    파일처리함수 = (fileContent as binary, 법인명_param as text) as table =>
        let
            // 1) Excel 로드
            엑셀 = Excel.Workbook(fileContent, null, true),
            시트 = 엑셀{0}[Data],  // 첫 번째 시트

            // 2) 법인코드 추출 (회사: 또는 회계단위:)
            회사_메타셀 =
                let
                    메타행 = Table.SelectRows(시트, each
                        List.AnyTrue({
                            Text.Contains(정규화([Column1]), "회사:"),
                            Text.Contains(정규화([Column2]), "회사:"),
                            Text.Contains(정규화([Column3]), "회사:")
                        })
                    ),
                    첫셀 = if Table.RowCount(메타행) > 0
                        then List.First(List.RemoveNulls(Record.ToList(메타행{0})))
                        else null
                in 첫셀,

            회계단위_메타셀 =
                let
                    메타행 = Table.SelectRows(시트, each
                        List.AnyTrue({
                            Text.Contains(정규화([Column1]), "회계단위:"),
                            Text.Contains(정규화([Column2]), "회계단위:"),
                            Text.Contains(정규화([Column3]), "회계단위:")
                        })
                    ),
                    첫셀 = if Table.RowCount(메타행) > 0
                        then List.First(List.RemoveNulls(Record.ToList(메타행{0})))
                        else null
                in 첫셀,

            메타셀 = if 회사_메타셀 <> null then 회사_메타셀 else 회계단위_메타셀,

            // 법인코드 파싱: "1000.에이치알이" → "1000"
            법인코드 =
                let
                    정규화된 = 정규화(메타셀),
                    콜론이후 = Text.AfterDelimiter(정규화된, ":"),
                    점이전 = Text.BeforeDelimiter(콜론이후, ".")
                in
                    if 점이전 = "" then 콜론이후 else 점이전,

            // 3) 헤더 처리 (상위 7행 제거 후 헤더 승격)
            상위행제거 = Table.Skip(시트, 7),
            헤더승격 = Table.PromoteHeaders(상위행제거, [PromoteAllScalars=true]),

            // 4) 컬럼명 확인 및 선택
            // 차변잔액 = "잔    액", 대변잔액 = "잔    액_1" 또는 "잔    액_2"
            컬럼목록 = Table.ColumnNames(헤더승격),
            차변잔액컬럼 = "잔    액",
            대변잔액컬럼 =
                if List.Contains(컬럼목록, "잔    액_1") then "잔    액_1"
                else if List.Contains(컬럼목록, "잔    액_2") then "잔    액_2"
                else null,
            계정과목컬럼 = "Column3",

            // 5) 계정과목 파싱 ([숫자...] 패턴만)
            계정필터 = Table.SelectRows(헤더승격, each
                let
                    raw = Record.Field(_, 계정과목컬럼),
                    정규화된 = if raw = null then "" else Text.Replace(Text.From(raw), " ", "")
                in
                    Text.StartsWith(정규화된, "[") and
                    let
                        afterBracket = Text.AfterDelimiter(정규화된, "["),
                        firstChar = if Text.Length(afterBracket) > 0 then Text.At(afterBracket, 0) else ""
                    in firstChar >= "0" and firstChar <= "9"
            ),

            // 6) 컬럼 추가: 법인코드, 법인명, 법인별 CoA, 법인별 계정과목명
            법인코드추가 = Table.AddColumn(계정필터, "법인코드", each 법인코드, type text),
            법인명추가 = Table.AddColumn(법인코드추가, "법인명", each 법인명_param, type text),

            CoA추가 = Table.AddColumn(법인명추가, "법인별 CoA", each
                let
                    raw = Record.Field(_, 계정과목컬럼),
                    정규화된 = if raw = null then "" else Text.Replace(Text.From(raw), " ", "")
                in
                    Text.BetweenDelimiters(정규화된, "[", "]"),
                type text
            ),

            계정과목명추가 = Table.AddColumn(CoA추가, "법인별 계정과목명", each
                let
                    raw = Record.Field(_, 계정과목컬럼),
                    정규화된 = if raw = null then "" else Text.Replace(Text.From(raw), " ", ""),
                    이름 = Text.AfterDelimiter(정규화된, "]")
                in
                    Text.Trim(이름),
                type text
            ),

            // 7) 당기 계산 (차변잔액 - 대변잔액)
            당기추가 = Table.AddColumn(계정과목명추가, "당기", each
                let
                    차변 = try Number.From(Record.Field(_, 차변잔액컬럼)) otherwise 0,
                    대변 = if 대변잔액컬럼 = null then 0
                           else try Number.From(Record.Field(_, 대변잔액컬럼)) otherwise 0
                in
                    (if 차변 = null then 0 else 차변) - (if 대변 = null then 0 else 대변),
                type number
            ),

            // 8) 식별코드 생성
            식별코드추가 = Table.AddColumn(당기추가, "식별코드", each
                [법인코드] & "_" & [법인별 CoA],
                type text
            ),

            // 9) 최종 컬럼 선택
            최종 = Table.SelectColumns(식별코드추가, {
                "법인코드", "법인명", "법인별 CoA", "법인별 계정과목명", "식별코드", "당기"
            })
        in
            최종,

    // ========== 모든 파일 처리 ==========
    파일처리 = Table.AddColumn(파일목록, "데이터", each
        try 파일처리함수([Content], [법인명]) otherwise #table(
            {"법인코드", "법인명", "법인별 CoA", "법인별 계정과목명", "식별코드", "당기"},
            {}
        )
    ),

    // ========== 모든 데이터 합치기 ==========
    데이터확장 = Table.ExpandTableColumn(파일처리, "데이터",
        {"법인코드", "법인명", "법인별 CoA", "법인별 계정과목명", "식별코드", "당기"}
    ),

    // ========== 필요 컬럼만 선택 및 타입 지정 ==========
    최종선택 = Table.SelectColumns(데이터확장, {
        "법인코드", "법인명", "법인별 CoA", "법인별 계정과목명", "식별코드", "당기"
    }),

    타입지정 = Table.TransformColumnTypes(최종선택, {
        {"법인코드", type text},
        {"법인명", type text},
        {"법인별 CoA", type text},
        {"법인별 계정과목명", type text},
        {"식별코드", type text},
        {"당기", Int64.Type}
    })
in
    타입지정
```

---

## 5. PTB 쿼리 (CoA 매핑)

TB 출력 컬럼에 맞춰 PTB도 업데이트 필요:

```m
let
    원본 = Table.NestedJoin(TB, {"식별코드"}, CoA_Processed, {"식별코드"}, "CoA_Processed", JoinKind.LeftOuter),
    확장 = Table.ExpandTableColumn(원본, "CoA_Processed", {"PwC_CoA"}, {"PwC_CoA"}),

    // Master 조인 (부호 적용)
    Master조인 = Table.NestedJoin(확장, {"PwC_CoA"}, Master, {"TB Account"}, "Master", JoinKind.LeftOuter),
    Master확장 = Table.ExpandTableColumn(Master조인, "Master", {"Account Name", "대분류", "중분류", "부호"}, {"PwC_계정명", "대분류", "중분류", "부호"}),

    // 부호 적용: 당기 * 부호
    부호적용 = Table.AddColumn(Master확장, "금액", each
        if [부호] = null then [당기] else [당기] * [부호],
        Int64.Type
    ),

    // 최종 컬럼 선택
    최종 = Table.SelectColumns(부호적용, {
        "법인코드", "법인명", "법인별 CoA", "법인별 계정과목명",
        "PwC_CoA", "PwC_계정명", "대분류", "중분류", "금액"
    })
in
    최종
```

---

## 6. 쿼리 생성 순서

| 순서 | 쿼리명 | 로드 방식 | 비고 |
|------|--------|----------|------|
| 1 | 메인주소 | 연결 전용 | Path 테이블 참조 |
| 2 | 회사결산_주소 | 연결 전용 | TempPath 테이블 참조 |
| 3 | Raw_Data | 연결 전용 | SPO 파일 목록 + 법인명 |
| 4 | TB | 연결 전용 | 파일 파싱 + 법인코드 |
| 5 | PTB | 테이블 로드 | BSPL 시트에 출력 |

---

## 7. 출력 컬럼 (최종)

### TB 출력

| 컬럼 | 타입 | 예시 |
|------|------|------|
| 법인코드 | text | `1000` |
| 법인명 | text | `HRE` |
| 법인별 CoA | text | `1030000` |
| 법인별 계정과목명 | text | `보통예금` |
| 식별코드 | text | `1000_1030000` |
| 당기 | Int64 | `1854605784` |

### PTB 출력 (CoA 매핑 후)

| 컬럼 | 타입 | 예시 |
|------|------|------|
| 법인코드 | text | `1000` |
| 법인명 | text | `HRE` |
| 법인별 CoA | text | `1030000` |
| 법인별 계정과목명 | text | `보통예금` |
| PwC_CoA | text | `110110` |
| PwC_계정명 | text | `Cash and cash equivalents` |
| 대분류 | text | `자산` |
| 중분류 | text | `유동자산` |
| 금액 | Int64 | `1854605784` |

---

## 8. 오류 처리

### 8.1 시산표 파일 여러 개

Power Query에서 error 발생 → VBA에서 처리:

```vba
' VBA에서 오류 메시지 확인
If InStr(Err.Description, "시산표 파일이 여러 개") > 0 Then
    MsgBox "해당 법인 폴더에 시산표 파일이 여러 개 있습니다." & vbCrLf & _
           "각 법인 폴더에 시산표 파일은 1개만 남겨주세요.", vbExclamation
End If
```

### 8.2 시산표 파일 없음

```vba
If InStr(Err.Description, "파일을 찾을 수 없습니다") > 0 Then
    MsgBox "시산표 파일이 없는 법인이 있습니다.", vbExclamation
End If
```

---

## 9. 법인명 활용 방안

### 9.1 현재 설계

- **법인코드** (1000): CoA_Processed, Master 테이블과 조인 시 사용
- **법인명** (HRE): 사용자 표시용, 리포트 출력용

### 9.2 Corp 테이블 연동

Corp 테이블에 법인코드-법인명 매핑이 있다면:

```m
// Corp 테이블 예시
| 법인코드 | 법인명 | 기능통화 |
|----------|--------|----------|
| 1000     | HRE    | KRW      |
| 2000     | 예천솔라 | KRW    |
```

검증 로직 추가 가능:
```m
// TB에서 법인코드-법인명 일치 여부 검증
법인명검증 = Table.NestedJoin(TB, {"법인코드"}, Corp, {"법인코드"}, "Corp", JoinKind.LeftOuter),
확장 = Table.ExpandTableColumn(법인명검증, "Corp", {"법인명"}, {"Corp_법인명"}),
불일치 = Table.SelectRows(확장, each [법인명] <> [Corp_법인명])
// 불일치 있으면 경고
```

---

**문서 버전**: 1.0
**작성일**: 2026-01-24
