# HRE 합계잔액시산표 Power Query 분석

## 1. 시산표 구조 분석

### 1.1 원본 파일 구조 (합계잔액시산표)

```
열 구조 (Column A ~ E):
┌─────────────┬─────────────┬──────────────────────┬─────────────┬─────────────┐
│   차 변     │   차 변     │     계 정 과 목      │   대 변     │   대 변     │
│   잔 액     │   합 계     │                      │   합 계     │   잔 액     │
├─────────────┼─────────────┼──────────────────────┼─────────────┼─────────────┤
│ 1,854,605,784│90,645,957,993│[1030000] 보통예금    │88,791,352,209│            │
│            │ 15,200,000,000│[1060000] 기타단기금융│15,200,000,000│            │
│            │            │[1150001] 대손충당금_내│205,840,588  │205,840,588  │
└─────────────┴─────────────┴──────────────────────┴─────────────┴─────────────┘

행 번호별 내용:
- 1~7행: 헤더 (합계잔액시산표, 기간, 회사명, 회계단위 등) → 제외
- 8행: 컬럼 헤더 (잔액, 합계, 계정과목, 합계, 잔액) → 제외
- 9~134행: 데이터 → 필터링 후 포함
- 135행: 합계 → 제외
```

### 1.2 계정과목 패턴 분류

| 패턴 | 예시 | 처리 |
|------|------|------|
| `<< ... >>` | `<< 자          산 >>` | **제외** - 대분류 소계 |
| `[ ... ]` (공백 포함) | `[ 유  동    자  산 ]` | **제외** - 중분류 소계 |
| `< ... >` | `< 당  좌    자  산 >` | **제외** - 소분류 소계 |
| `[숫자...]` | `[1030000] 보 통 예 금` | **포함** - 계정과목 |
| `합      계` | `합      계` | **제외** - 합계행 |

### 1.3 필터링 조건 (Regex 관점)

```
포함 조건: 계정과목이 `^\[\d` 패턴으로 시작
- [1030000] → 매치 (숫자로 시작)
- [ 유동자산 ] → 불일치 (공백으로 시작)
- << 자산 >> → 불일치 (<로 시작)

M 언어 구현:
Text.StartsWith([계정과목], "[") and
  let char = Text.At(Text.AfterDelimiter([계정과목], "["), 0)
  in char >= "0" and char <= "9"
```

## 2. 기존 로직 분석

### 2.1 차변/대변 부호 처리

**기존 프로젝트 로직** (PTB 쿼리):
```m
// Master 테이블에서 부호 값 조인
#"확장된 부호" = Table.ExpandTableColumn(..., "Master", {..., "부호"}, {..., "부호"}),

// 부호 적용: 당기 * 부호
#"추가된 사용자 지정 항목" = Table.AddColumn(..., "당기 잔액", each [당기] * [#"부호"])
```

- `부호` = 1 (차변 계정: 자산, 비용)
- `부호` = -1 (대변 계정: 부채, 자본, 수익)

**HRE 프로젝트 로직**:
- 합계잔액시산표는 이미 차변잔액/대변잔액으로 분리되어 있음
- 계산식: `금액 = 차변잔액 - 대변잔액`
  - 자산 (차변 계정): 차변잔액 > 0, 대변잔액 = 0 → 양수
  - 부채 (대변 계정): 차변잔액 = 0, 대변잔액 > 0 → 음수

### 2.2 음수 표시 처리

원본 데이터에 이미 음수 값이 존재:
- `[1350001] 부가세대급금_일반전표`: 차변잔액 = -264,332,207
- `[2550001] 부가세예수금_일반전표`: 대변잔액 = -736,025,545

→ 별도 음수 변환 로직 불필요 (원본 그대로 사용)

## 3. SPO 폴더 구조 (신규)

### 3.1 기존 구조 vs 신규 구조

```
기존 구조 (BEP 프로젝트):
폴더-법인명-연도-파일
예: ★시연용폴더★/[HRE-KR01]/회사결산2512/시산표.xlsx

신규 구조 (HRE 프로젝트):
폴더-연도-법인명-파일
예: ★시연용폴더★/2512/HRE/12월 말 합계잔액시산표_에이치알이_수정.xlsx
```

### 3.2 신규 폴더 구조

```
★시연용폴더★/
└── {결산연월}/           (예: 2512 = 2025년 12월)
    ├── {법인명}/         (예: HRE, 예천솔라)
    │   └── *시산표*.xlsx  (파일명에 "시산표" 포함)
    ├── {법인명}/
    │   └── *시산표*.xlsx
    └── ...
```

### 3.3 법인코드 추출 변경

```
기존: Text.BetweenDelimiters([Folder Path], "[", "]")
     → /[HRE-KR01]/ 에서 "HRE-KR01" 추출

신규: 폴더 경로에서 법인명 폴더 추출
     → /2512/HRE/ 에서 "HRE" 추출
     경로 분석: 결산연월 폴더 다음 폴더가 법인명
```

## 4. 단일 파일 테스트용 M 함수

### 4.1 로컬 파일 기준 M 코드

```m
let
    // ========== 설정 ==========
    파일경로 = "C:\Users\jkim886\Desktop\DA\Project\HRE\test\시산표_샘플.xlsx",
    시트이름 = "Sheet1",  // 또는 실제 시트 이름

    // ========== 파일 로드 ==========
    원본 = Excel.Workbook(File.Contents(파일경로), null, true),
    시트 = 원본{[Item=시트이름, Kind="Sheet"]}[Data],

    // ========== 컬럼 이름 지정 ==========
    // 열: A=차변잔액, B=차변합계, C=계정과목, D=대변합계, E=대변잔액
    컬럼지정 = Table.RenameColumns(시트, {
        {"Column1", "차변잔액"},
        {"Column2", "차변합계"},
        {"Column3", "계정과목"},
        {"Column4", "대변합계"},
        {"Column5", "대변잔액"}
    }),

    // ========== 계정과목 필터링 ==========
    // "[숫자" 패턴으로 시작하는 행만 선택
    필터링 = Table.SelectRows(컬럼지정, each
        [계정과목] <> null and
        Text.StartsWith(Text.Trim([계정과목]), "[") and
        let
            afterBracket = Text.AfterDelimiter(Text.Trim([계정과목]), "["),
            firstChar = if Text.Length(afterBracket) > 0 then Text.At(afterBracket, 0) else ""
        in
            firstChar >= "0" and firstChar <= "9"
    ),

    // ========== 계정코드와 계정과목명 분리 ==========
    // [1030000] 보 통 예 금 → 계정코드: 1030000, 계정과목명: 보통예금
    계정코드추출 = Table.AddColumn(필터링, "계정코드", each
        Text.BetweenDelimiters(Text.Trim([계정과목]), "[", "]"),
        type text
    ),

    계정과목명추출 = Table.AddColumn(계정코드추출, "계정과목명", each
        let
            afterCode = Text.AfterDelimiter(Text.Trim([계정과목]), "]"),
            cleaned = Text.Replace(afterCode, " ", ""),  // 공백 제거
            trimmed = Text.Trim(cleaned)
        in
            trimmed,
        type text
    ),

    // ========== 숫자 타입 변환 ==========
    타입변환 = Table.TransformColumnTypes(계정과목명추출, {
        {"차변잔액", type number},
        {"대변잔액", type number}
    }),

    // ========== 금액 계산 (차변잔액 - 대변잔액) ==========
    금액계산 = Table.AddColumn(타입변환, "금액", each
        let
            차변 = if [차변잔액] = null then 0 else [차변잔액],
            대변 = if [대변잔액] = null then 0 else [대변잔액]
        in
            차변 - 대변,
        type number
    ),

    // ========== 필요 컬럼만 선택 ==========
    최종결과 = Table.SelectColumns(금액계산, {
        "계정코드", "계정과목명", "차변잔액", "대변잔액", "금액"
    })
in
    최종결과
```

### 4.2 SharePoint 단일 파일 테스트용 M 코드

```m
let
    // ========== 설정 ==========
    메인주소 = "https://pwckor.sharepoint.com/sites/KR-ASR-HRE_Consolidation",
    대상폴더 = "/★시연용폴더★/2512/HRE",

    // ========== SharePoint 파일 목록 ==========
    원본 = SharePoint.Files(메인주소, [ApiVersion = 15]),
    폴더필터 = Table.SelectRows(원본, each Text.Contains([Folder Path], 대상폴더)),

    // ========== 시산표 파일 필터 ==========
    시산표파일 = Table.SelectRows(폴더필터, each Text.Contains([Name], "시산표")),
    파일수 = Table.RowCount(시산표파일),

    // ========== 파일 수 검증 (1개만 허용) ==========
    // 여러 개면 오류 발생 (VBA에서 Alert 처리)
    검증결과 = if 파일수 = 0 then
            error "시산표 파일을 찾을 수 없습니다: " & 대상폴더
        else if 파일수 > 1 then
            error "시산표 파일이 여러 개 있습니다 (" & Text.From(파일수) & "개): " & 대상폴더
        else
            시산표파일{0},

    // ========== Excel 파일 로드 ==========
    엑셀내용 = Excel.Workbook(검증결과[Content], null, true),

    // 첫 번째 시트 사용 (또는 특정 시트명)
    첫번째시트 = 엑셀내용{0}[Data],

    // ========== 이후 로직은 4.1과 동일 ==========
    컬럼지정 = Table.RenameColumns(첫번째시트, {
        {"Column1", "차변잔액"},
        {"Column2", "차변합계"},
        {"Column3", "계정과목"},
        {"Column4", "대변합계"},
        {"Column5", "대변잔액"}
    }),

    필터링 = Table.SelectRows(컬럼지정, each
        [계정과목] <> null and
        Text.StartsWith(Text.Trim([계정과목]), "[") and
        let
            afterBracket = Text.AfterDelimiter(Text.Trim([계정과목]), "["),
            firstChar = if Text.Length(afterBracket) > 0 then Text.At(afterBracket, 0) else ""
        in
            firstChar >= "0" and firstChar <= "9"
    ),

    계정코드추출 = Table.AddColumn(필터링, "계정코드", each
        Text.BetweenDelimiters(Text.Trim([계정과목]), "[", "]"),
        type text
    ),

    계정과목명추출 = Table.AddColumn(계정코드추출, "계정과목명", each
        Text.Replace(Text.Trim(Text.AfterDelimiter(Text.Trim([계정과목]), "]")), " ", ""),
        type text
    ),

    타입변환 = Table.TransformColumnTypes(계정과목명추출, {
        {"차변잔액", type number},
        {"대변잔액", type number}
    }),

    금액계산 = Table.AddColumn(타입변환, "금액", each
        (if [차변잔액] = null then 0 else [차변잔액]) -
        (if [대변잔액] = null then 0 else [대변잔액]),
        type number
    ),

    최종결과 = Table.SelectColumns(금액계산, {
        "계정코드", "계정과목명", "차변잔액", "대변잔액", "금액"
    })
in
    최종결과
```

### 4.3 법인코드 포함 버전

```m
let
    // ========== 설정 ==========
    메인주소 = "https://pwckor.sharepoint.com/sites/KR-ASR-HRE_Consolidation",
    결산연월 = "2512",
    법인명 = "HRE",

    대상폴더 = "/★시연용폴더★/" & 결산연월 & "/" & 법인명,

    원본 = SharePoint.Files(메인주소, [ApiVersion = 15]),
    폴더필터 = Table.SelectRows(원본, each Text.Contains([Folder Path], 대상폴더)),
    시산표파일 = Table.SelectRows(폴더필터, each Text.Contains([Name], "시산표")),

    // 파일 수 검증
    파일수 = Table.RowCount(시산표파일),
    검증결과 = if 파일수 <> 1 then
        error "시산표 파일 오류: " & Text.From(파일수) & "개 발견 (" & 대상폴더 & ")"
        else 시산표파일{0},

    // Excel 로드
    엑셀내용 = Excel.Workbook(검증결과[Content], null, true),
    첫번째시트 = 엑셀내용{0}[Data],

    // 컬럼 지정
    컬럼지정 = Table.RenameColumns(첫번째시트, {
        {"Column1", "차변잔액"}, {"Column2", "차변합계"},
        {"Column3", "계정과목"}, {"Column4", "대변합계"}, {"Column5", "대변잔액"}
    }),

    // 계정과목 필터링 ([숫자...] 패턴)
    필터링 = Table.SelectRows(컬럼지정, each
        [계정과목] <> null and
        Text.StartsWith(Text.Trim([계정과목]), "[") and
        let
            afterBracket = Text.AfterDelimiter(Text.Trim([계정과목]), "["),
            firstChar = if Text.Length(afterBracket) > 0 then Text.At(afterBracket, 0) else ""
        in firstChar >= "0" and firstChar <= "9"
    ),

    // 법인코드 추가
    법인코드추가 = Table.AddColumn(필터링, "법인코드", each 법인명, type text),

    // 계정코드/계정과목명 분리
    계정코드추출 = Table.AddColumn(법인코드추가, "계정코드", each
        Text.BetweenDelimiters(Text.Trim([계정과목]), "[", "]"), type text),
    계정과목명추출 = Table.AddColumn(계정코드추출, "계정과목명", each
        Text.Replace(Text.Trim(Text.AfterDelimiter(Text.Trim([계정과목]), "]")), " ", ""), type text),

    // 숫자 변환 및 금액 계산
    타입변환 = Table.TransformColumnTypes(계정과목명추출, {
        {"차변잔액", type number}, {"대변잔액", type number}
    }),
    금액계산 = Table.AddColumn(타입변환, "금액", each
        (if [차변잔액] = null then 0 else [차변잔액]) -
        (if [대변잔액] = null then 0 else [대변잔액]), type number),

    // 최종 컬럼 선택
    최종결과 = Table.SelectColumns(금액계산, {
        "법인코드", "계정코드", "계정과목명", "차변잔액", "대변잔액", "금액"
    })
in
    최종결과
```

## 5. 예상 결과

### 5.1 입력 데이터 (원본 시산표)

| 차변잔액 | 차변합계 | 계정과목 | 대변합계 | 대변잔액 |
|----------|----------|----------|----------|----------|
| 1,854,605,784 | 90,645,957,993 | [1030000] 보 통 예 금 | 88,791,352,209 | |
| | | [1150001] 대손충당금_내부거래 | 205,840,588 | 205,840,588 |
| -264,332,207 | -264,332,207 | [1350001] 부가세대급금_일반전표 | | |
| | | [2510000] 외상매입금 | 547,531,518 | 547,531,518 |

### 5.2 출력 데이터 (변환 후)

| 법인코드 | 계정코드 | 계정과목명 | 차변잔액 | 대변잔액 | 금액 |
|----------|----------|------------|----------|----------|------|
| HRE | 1030000 | 보통예금 | 1,854,605,784 | 0 | 1,854,605,784 |
| HRE | 1150001 | 대손충당금_내부거래 | 0 | 205,840,588 | -205,840,588 |
| HRE | 1350001 | 부가세대급금_일반전표 | -264,332,207 | 0 | -264,332,207 |
| HRE | 2510000 | 외상매입금 | 0 | 547,531,518 | -547,531,518 |

## 6. VBA Alert 처리 (참고)

Power Query에서 `error`가 발생하면 VBA에서 처리:

```vba
' QueryRefresh 시 오류 처리
Private Sub RefreshPTB()
    On Error GoTo ErrorHandler

    ThisWorkbook.Connections("쿼리 - PTB").Refresh
    Exit Sub

ErrorHandler:
    If InStr(Err.Description, "시산표 파일이 여러 개") > 0 Then
        MsgBox "해당 폴더에 시산표 파일이 여러 개 있습니다." & vbCrLf & _
               "하나의 파일만 남기고 삭제해주세요.", vbExclamation, "파일 오류"
    ElseIf InStr(Err.Description, "시산표 파일을 찾을 수 없습니다") > 0 Then
        MsgBox "해당 폴더에 시산표 파일이 없습니다." & vbCrLf & _
               "파일을 업로드해주세요.", vbExclamation, "파일 오류"
    Else
        MsgBox "Power Query 오류: " & Err.Description, vbCritical, "오류"
    End If
End Sub
```

---

**문서 버전**: 1.0
**작성일**: 2026-01-23
**작성자**: HRE 프로젝트
