# HRE 연결마스터 구현 완성 가이드 (초보자용)

> **현재 상태**: VBA 모듈 + UserForms 임포트 완료
> **목표**: Power Query 구현 + VBA 코드 업데이트 + Excel 테이블 설정
> **작성일**: 2026-01-24

---

## 전체 작업 순서

```
1단계: Excel 시트 & 테이블 설정 (HideSheet, BSPL)
   ↓
2단계: VBA 코드 업데이트 (mod_03, mod_05)
   ↓
3단계: Power Query 쿼리 생성 (6개, 순서대로)
   ↓
4단계: 테스트 & 검증
```

---

# 1단계: Excel 시트 & 테이블 설정

## 1.1 왜 이게 먼저인가요?

Power Query가 **HideSheet의 테이블**을 참조해서 SharePoint 경로를 읽어옵니다.
테이블이 없으면 Power Query가 "테이블을 찾을 수 없습니다" 에러를 냅니다.

---

## 1.2 HideSheet에 Path 테이블 만들기

**목적**: SharePoint 메인 URL 저장 (Power Query `메인주소` 쿼리가 참조)

### 단계별 실행:

1. **HideSheet 시트로 이동** (숨김 상태면: 시트 탭 우클릭 → 숨기기 취소)

2. **E1:E2 범위에 입력**:
   ```
   E1: Path
   E2: https://pwckor.sharepoint.com/sites/KR-ASR-HRE_Consolidation
   ```

   > **중요**: E1 셀에 정확히 `Path`라고 입력해야 합니다! (대소문자 구분)

   실제 화면:
   ```
   ┌─────────────────────────────────────────────────────────────┐
   │  E1 │ Path                                                  │
   ├─────┼───────────────────────────────────────────────────────┤
   │  E2 │ https://pwckor.sharepoint.com/sites/KR-ASR-HRE_...    │
   └─────┴───────────────────────────────────────────────────────┘
   ```

3. **테이블로 변환**:
   - E1:E2 범위 선택
   - `Ctrl + T` 누르기
   - "머리글 포함" 체크 확인 → **확인**

4. **테이블 이름 변경**:
   - 테이블 선택 상태에서 상단 **테이블 디자인** 탭 클릭
   - 좌측 "테이블 이름"을 `Path`로 변경

---

## 1.3 HideSheet에 TempPath 테이블 만들기

**목적**: 결산 폴더 경로 저장 (Power Query `회사결산_주소` 쿼리가 참조)

### 단계별 실행:

1. **H1:H2 범위에 입력**:
   ```
   H1: 경로
   H2: /Shared Documents/★시연용폴더★/2512
   ```

   > **참고**: 경로는 나중에 frmDirectory 폼에서 선택하거나 직접 수정 가능

2. **테이블로 변환**: `Ctrl + T`

3. **테이블 이름**: `TempPath`

---

## 1.4 BSPL 시트의 PTB 테이블 수정 (8컬럼 → 9컬럼)

**왜 수정하나요?**
- 기존 BEP: 법인코드 바로 다음에 법인별CoA
- HRE: 법인코드 다음에 **법인명** 컬럼 추가

### 기존 (BEP 8컬럼):
```
B4  | C4        | D4             | E4      | F4       | G4   | H4   | I4
법인코드 | 법인별CoA | 법인별계정과목명 | PwC_CoA | PwC_계정명 | 대분류 | 중분류 | 금액
```

### 변경 (HRE 9컬럼):
```
B4  | C4   | D4        | E4             | F4      | G4       | H4   | I4   | J4
법인코드 | 법인명 | 법인별CoA | 법인별계정과목명 | PwC_CoA | PwC_계정명 | 대분류 | 중분류 | 금액
        ↑ 신규!
```

### 수정 방법:

1. **BSPL 시트로 이동**

2. **기존 PTB 테이블이 있다면**:
   - C열(법인별CoA)을 선택
   - 우클릭 → **삽입** → 열 삽입
   - 새로운 C4 셀에 `법인명` 입력

3. **테이블이 없다면** (새로 만들기):
   - B4:J4에 9개 헤더 입력:
     ```
     법인코드 | 법인명 | 법인별CoA | 법인별계정과목명 | PwC_CoA | PwC_계정명 | 대분류 | 중분류 | 금액
     ```
   - B4 선택 → `Ctrl + T` → 테이블 이름: `PTB`

---

# 2단계: VBA 코드 업데이트

## 2.1 왜 VBA 코드를 수정하나요?

법인명 컬럼이 추가되면서 **모든 컬럼 위치가 1칸씩 밀렸습니다**.

```
기존: 법인별CoA = 3번째 컬럼
변경: 법인별CoA = 4번째 컬럼 (법인명이 2번째로 들어왔으니까)
```

VBA에서 `coaRow.Range(1, 3)`처럼 숫자로 컬럼을 참조하는 부분을 모두 +1 해야 합니다.

---

## 2.2 mod_03_PTB_CoA_Input.bas 수정

**VBA 편집기 열기**: `Alt + F11`

**모듈 찾기**: 프로젝트 탐색기 → 모듈 → `mod_03_PTB_CoA_Input` 더블클릭

### 수정 1: Line 46 근처

**찾기** (`Ctrl + F`): `.Resize(, 5)`

**기존 코드**:
```vba
   With tblPTB.DataBodyRange
       Set visibleRange = .Resize(, 5).SpecialCells(xlCellTypeVisible)
   End With
```

**변경 코드**:
```vba
   With tblPTB.DataBodyRange
       ' HRE v1.00: 6컬럼으로 확대 (법인명 컬럼 추가)
       ' 법인코드, 법인명, 법인별CoA, 법인별계정과목명, PwC_CoA, PwC_계정명
       Set visibleRange = .Resize(, 6).SpecialCells(xlCellTypeVisible)
   End With
```

**설명**: PTB에서 복사할 컬럼 수가 5개에서 6개로 늘었습니다.

---

### 수정 2: Line 114 근처

**찾기**: `coaRow.Range(1, 3).Value`

**기존 코드**:
```vba
            ptbAccount = coaRow.Range(1, 3).Value  ' 법인별 CoA (from PTB)
```

**변경 코드**:
```vba
            ' HRE v1.00: 컬럼 인덱스 +1 (법인명 컬럼 추가로 인함)
            ptbAccount = coaRow.Range(1, 4).Value  ' 법인별 CoA (from PTB, 4번째 컬럼)
```

**설명**: 법인별CoA가 3번째에서 4번째 컬럼으로 이동했습니다.

---

### 수정 3: Line 144-145 근처

**찾기**: `coaRow.Range(1, 4).Value = suggestedAccount`

**기존 코드**:
```vba
            ' Populate suggestion columns
            coaRow.Range(1, 4).Value = suggestedAccount  ' PwC_CoA column
            coaRow.Range(1, 5).Value = suggestedDescription  ' PwC_계정명 column
```

**변경 코드**:
```vba
            ' Populate suggestion columns
            ' HRE v1.00: 컬럼 인덱스 +1 (법인명 컬럼 추가로 인함)
            coaRow.Range(1, 5).Value = suggestedAccount  ' PwC_CoA column (5번째)
            coaRow.Range(1, 6).Value = suggestedDescription  ' PwC_계정명 column (6번째)
```

---

## 2.3 mod_05_PTB_Highlight.bas 수정

**모듈 찾기**: `mod_05_PTB_Highlight` 더블클릭

### 수정: Line 80 근처

**찾기**: `rng.Cells(i, 4)`

**기존 코드**:
```vba
    For i = 1 To rng.Rows.count
        If IsEmpty(rng.Cells(i, 4)) Then
            rng.Cells(i, 1).Resize(1, lastCol).Interior.Color = vbYellow
        Else
            rng.Cells(i, 1).Resize(1, lastCol).Interior.Color = vbWhite
        End If
    Next i
```

**변경 코드**:
```vba
    ' HRE v1.00: PwC_CoA는 5번째 컬럼 (법인명 컬럼 추가로 인함)
    ' PTB 컬럼 순서: 법인코드(1), 법인명(2), 법인별CoA(3), 법인별계정과목명(4), PwC_CoA(5), ...
    For i = 1 To rng.Rows.count
        If IsEmpty(rng.Cells(i, 5)) Then
            rng.Cells(i, 1).Resize(1, lastCol).Interior.Color = vbYellow
        Else
            rng.Cells(i, 1).Resize(1, lastCol).Interior.Color = vbWhite
        End If
    Next i
```

**설명**: PwC_CoA가 비어있는 행을 노란색으로 표시하는 로직인데, PwC_CoA가 4번째에서 5번째로 이동했습니다.

---

## 2.4 VBA 컴파일 확인

1. VBA 편집기에서 **디버그** → **VBAProject 컴파일**
2. 에러 없으면 성공
3. 에러 있으면 해당 줄 확인 후 수정

---

# 3단계: Power Query 쿼리 생성 (6개)

## 3.1 Power Query란?

Excel에서 **외부 데이터를 가져와서 변환**하는 도구입니다.
SharePoint의 여러 법인 폴더에서 시산표 파일들을 자동으로 모아서 하나의 PTB 테이블로 만들어줍니다.

## 3.2 쿼리 생성 순서 (중요!)

```
① fnProcessSingleFile (함수) ← 다른 쿼리들이 의존
② 메인주소 ← Path 테이블 참조
③ 회사결산_주소 ← TempPath 테이블 참조
④ Raw_Data ← ②③ 의존
⑤ TB ← ④① 의존
⑥ PTB ← ⑤ 의존 (최종 출력)
```

**반드시 위 순서대로 만드세요!** 아래 쿼리가 위 쿼리를 참조하므로, 순서가 틀리면 에러납니다.

---

## 3.3 쿼리 생성 방법 (공통)

1. **데이터** 탭 → **데이터 가져오기** → **기타 원본에서** → **빈 쿼리**
2. Power Query 편집기 열림
3. **홈** 탭 → **고급 편집기** 클릭
4. 기존 코드 삭제 → 새 코드 붙여넣기
5. **완료** 클릭

---

## 쿼리 1: fnProcessSingleFile

**역할**: 개별 시산표 파일에서 법인코드, 법인명, CoA, 금액 추출하는 함수

**M 코드** (전체 복사):

```m
let
    정규화 = (val) => let
        str = if val = null then "" else Text.From(val),
        cleaned = Text.Trim(Text.Clean(str))
    in cleaned,

    법인코드_추출 = (메타셀) => let
        t0 = 정규화(메타셀),
        t1 = if t0 = "" then null else Text.AfterDelimiter(t0, ":"),
        t2 = if t1 = null then null else Text.BeforeDelimiter(t1, ".")
    in t2,

    법인명_추출 = (메타셀) => let
        t0 = 정규화(메타셀),
        t1 = if t0 = "" then null else Text.AfterDelimiter(t0, ":"),
        name = if t1 = null then null else Text.AfterDelimiter(t1, ".")
    in name,

    법인별CoA_파싱 = (rawCode) => let
        str = 정규화(rawCode),
        hasBracket = Text.Contains(str, "["),
        code = if hasBracket then Text.BeforeDelimiter(Text.AfterDelimiter(str, "["), "]")
               else if Text.Length(str) >= 7 then Text.Start(str, 7) else str
    in code,

    계정과목명_파싱 = (rawName) => let
        str = 정규화(rawName),
        hasBracket = Text.Contains(str, "]"),
        name = if hasBracket then Text.AfterDelimiter(str, "]") else str
    in name,

    당기_계산 = (차변, 대변) => let
        debit = if 차변 = null then 0 else Number.From(차변),
        credit = if 대변 = null then 0 else Number.From(대변)
    in debit - credit,

    fnProcessSingleFile = (fileContent as binary, 법인명_폴더 as text) =>
    let
        Source = Excel.Workbook(fileContent, null, true),
        FirstSheet = Source{0}[Data],
        메타셀 = FirstSheet{0}[Column1],
        법인코드 = 법인코드_추출(메타셀),
        법인명_메타 = 법인명_추출(메타셀),
        법인명_최종 = if 법인명_메타 <> null and 법인명_메타 <> "" then 법인명_메타 else 법인명_폴더,
        DataRows = Table.Skip(FirstSheet, 4),
        PromotedHeaders = Table.PromoteHeaders(DataRows, [PromoteAllScalars=true]),
        colNames = Table.ColumnNames(PromotedHeaders),
        차변컬럼 = List.First(List.Select(colNames, each Text.Contains(_, "차변")), "차변잔액"),
        대변컬럼 = List.First(List.Select(colNames, each Text.Contains(_, "대변")), "대변잔액"),
        계정코드컬럼 = List.First(List.Select(colNames, each Text.Contains(_, "계정코드") or Text.Contains(_, "코드")), "계정코드"),
        계정명컬럼 = List.First(List.Select(colNames, each Text.Contains(_, "계정과목") or Text.Contains(_, "계정명")), "계정과목명"),
        AddCols = Table.AddColumn(
            Table.AddColumn(
                Table.AddColumn(
                    Table.AddColumn(
                        Table.AddColumn(PromotedHeaders, "법인코드", each 법인코드, type text),
                        "법인명", each 법인명_최종, type text),
                    "법인별CoA", each try 법인별CoA_파싱(Record.Field(_, 계정코드컬럼)) otherwise null, type text),
                "법인별계정과목명", each try 계정과목명_파싱(Record.Field(_, 계정명컬럼)) otherwise null, type text),
            "당기", each try 당기_계산(Record.Field(_, 차변컬럼), Record.Field(_, 대변컬럼)) otherwise 0, type number),
        FinalColumns = Table.SelectColumns(AddCols, {"법인코드", "법인명", "법인별CoA", "법인별계정과목명", "당기"}),
        FilteredRows = Table.SelectRows(FinalColumns, each [법인별CoA] <> null and [법인별CoA] <> "")
    in FilteredRows
in fnProcessSingleFile
```

**저장 방법**:
1. 코드 붙여넣기 → **완료**
2. 좌측 쿼리 목록에서 쿼리 이름 우클릭 → **이름 바꾸기** → `fnProcessSingleFile`
3. **홈** → **닫기 및 로드** → **닫기 및 로드 위치...**
4. **연결만 만들기** 선택 → **확인**

---

## 쿼리 2: 메인주소

**역할**: HideSheet의 Path 테이블에서 SharePoint URL 읽기

**M 코드** (컬럼명 무관하게 동작하는 버전):
```m
let
    Source = Excel.CurrentWorkbook(){[Name="Path"]}[Content],
    FirstRow = Source{0},
    // 컬럼명에 상관없이 첫 번째 값 가져오기
    PathValue = Record.FieldValues(FirstRow){0},
    CleanedPath = Text.Trim(Text.From(PathValue))
in
    CleanedPath
```

> **참고**: 이 코드는 Path 테이블의 컬럼명이 "Path"가 아니어도 동작합니다.
> 만약 에러가 나면 HideSheet E1 셀이 정확히 `Path`인지 확인하세요.

**저장**: 쿼리 이름 `메인주소`, **연결만 만들기**

---

## 쿼리 3: 회사결산_주소

**역할**: HideSheet의 TempPath 테이블에서 결산 폴더 경로 읽기

**M 코드** (컬럼명 무관하게 동작하는 버전):
```m
let
    Source = Excel.CurrentWorkbook(){[Name="TempPath"]}[Content],
    FirstRow = Source{0},
    // 컬럼명에 상관없이 첫 번째 값 가져오기
    PathValue = Record.FieldValues(FirstRow){0},
    CleanedPath = Text.Trim(Text.From(PathValue))
in
    CleanedPath
```

**저장**: 쿼리 이름 `회사결산_주소`, **연결만 만들기**

---

## 쿼리 4: Raw_Data

**역할**: SharePoint에서 시산표 파일 목록 가져오기

**M 코드**:
```m
let
    메인주소 = 메인주소,
    회사결산_주소 = 회사결산_주소,
    Source = SharePoint.Files(메인주소, [ApiVersion = 15]),
    FilteredFolder = Table.SelectRows(Source, each Text.Contains([Folder Path], 회사결산_주소)),
    FilteredFiles = Table.SelectRows(FilteredFolder,
        each Text.Contains([Name], "시산표") and
             (Text.EndsWith([Name], ".xlsx") or Text.EndsWith([Name], ".xls"))),
    AddCorpFolder = Table.AddColumn(FilteredFiles, "법인명_폴더", each
        let
            path = [Folder Path],
            trimmed = if Text.EndsWith(path, "/") then Text.Start(path, Text.Length(path) - 1) else path,
            parts = Text.Split(trimmed, "/"),
            lastPart = List.Last(parts)
        in lastPart, type text),
    SelectedColumns = Table.SelectColumns(AddCorpFolder, {"법인명_폴더", "Name", "Content", "Folder Path"})
in
    SelectedColumns
```

**저장**: 쿼리 이름 `Raw_Data`, **연결만 만들기**

**첫 실행 시 인증 필요**:
- "인증 필요" 창 → **조직 계정** → **로그인**
- Office 365 계정으로 로그인

---

## 쿼리 5: TB

**역할**: 각 시산표 파일에 fnProcessSingleFile 함수 적용

**M 코드**:
```m
let
    Source = Raw_Data,
    ProcessFile = fnProcessSingleFile,
    AddProcessedData = Table.AddColumn(Source, "ProcessedData", each
        try ProcessFile([Content], [법인명_폴더])
        otherwise #table({"법인코드", "법인명", "법인별CoA", "법인별계정과목명", "당기"}, {}), type table),
    ExpandedData = Table.ExpandTableColumn(AddProcessedData, "ProcessedData",
        {"법인코드", "법인명", "법인별CoA", "법인별계정과목명", "당기"}),
    SelectedColumns = Table.SelectColumns(ExpandedData,
        {"법인코드", "법인명", "법인별CoA", "법인별계정과목명", "당기"}),
    FilteredRows = Table.SelectRows(SelectedColumns,
        each [법인코드] <> null and [법인별CoA] <> null),
    TypedColumns = Table.TransformColumnTypes(FilteredRows, {
        {"법인코드", type text}, {"법인명", type text}, {"법인별CoA", type text},
        {"법인별계정과목명", type text}, {"당기", type number}})
in
    TypedColumns
```

**저장**: 쿼리 이름 `TB`, **연결만 만들기**

---

## 쿼리 6: PTB (최종!)

**역할**: VBA가 채울 컬럼 추가 + BSPL 시트에 테이블 로드

**M 코드**:
```m
let
    Source = TB,
    AddPwC_CoA = Table.AddColumn(Source, "PwC_CoA", each null, type text),
    AddPwC_계정명 = Table.AddColumn(AddPwC_CoA, "PwC_계정명", each null, type text),
    Add대분류 = Table.AddColumn(AddPwC_계정명, "대분류", each null, type text),
    Add중분류 = Table.AddColumn(Add대분류, "중분류", each null, type text),
    Add금액 = Table.AddColumn(Add중분류, "금액", each [당기], type number),
    FinalColumns = Table.SelectColumns(Add금액, {
        "법인코드", "법인명", "법인별CoA", "법인별계정과목명",
        "PwC_CoA", "PwC_계정명", "대분류", "중분류", "금액"}),
    SortedTable = Table.Sort(FinalColumns, {{"법인코드", Order.Ascending}, {"법인별CoA", Order.Ascending}})
in
    SortedTable
```

**저장 (다르게!)**:
1. 쿼리 이름: `PTB`
2. **홈** → **닫기 및 로드** → **닫기 및 로드 위치...**
3. **테이블** 선택
4. **기존 워크시트** 선택 → `BSPL!$B$4` 입력
5. **확인**

---

# 4단계: 테스트 & 검증

## 4.1 Power Query 새로고침

1. **데이터** 탭 → **모두 새로 고침**
2. BSPL 시트 확인 → PTB 테이블에 데이터 채워짐

**확인 사항**:
- [ ] 9개 컬럼 표시 (법인코드 ~ 금액)
- [ ] 법인코드, 법인명 정상 추출
- [ ] PwC_CoA ~ 중분류는 빈 상태 (정상)

## 4.2 VBA CoA 매핑 테스트

1. 리본 메뉴 → **Update** (또는 해당 버튼)
2. 리본 메뉴 → **CoA First Drafting**
3. AddCoA 시트에 매핑 결과 표시

---

# 문제 해결 가이드

## 자주 발생하는 에러

### 에러 1: "레코드의 'Path' 필드를 찾을 수 없습니다"

**원인**: Path 테이블의 컬럼명이 "Path"가 아님

**해결 방법 A - 테이블 수정**:
1. HideSheet로 이동
2. E1 셀이 정확히 `Path`인지 확인
3. 다른 값이면 `Path`로 수정

**해결 방법 B - M 코드 수정** (권장):
위 가이드의 "메인주소" 쿼리는 이미 컬럼명에 상관없이 동작하도록 수정되어 있습니다.
`Record.FieldValues(FirstRow){0}` 코드가 컬럼명 무관하게 첫 번째 값을 가져옵니다.

---

### 에러 2: "TempPath 테이블을 찾을 수 없습니다"

**원인**: HideSheet에 TempPath 테이블이 없음

**해결**:
1. HideSheet H1:H2에 테이블 생성
2. 테이블 이름을 정확히 `TempPath`로 설정

---

### 에러 3: Raw_Data 결과가 비어있음

**원인**: TempPath 경로가 SharePoint 실제 경로와 불일치

**해결**:
1. SharePoint에서 실제 폴더 경로 확인
2. HideSheet TempPath 테이블의 경로 수정
3. 경로 구분자는 `/` 사용 (역슬래시 `\` 아님)

---

### 에러 4: 법인코드가 null

**원인**: 시산표 파일 A1 셀 형식이 맞지 않음

**해결**:
- 시산표 A1 셀 형식: `회사: 1000.에이치알이주식회사`
- 콜론(`:`)과 점(`.`)이 반드시 있어야 함

---

### 에러 5: VBA 컴파일 에러

**원인**: VBA 코드 수정 누락

**해결**: 2단계 VBA 수정 내용 다시 확인
- mod_03: 3곳 수정
- mod_05: 1곳 수정

---

# 체크리스트

## 1단계: Excel 설정
- [ ] HideSheet에 Path 테이블 (E1:E2)
- [ ] HideSheet에 TempPath 테이블 (H1:H2)
- [ ] BSPL의 PTB 테이블 9컬럼

## 2단계: VBA 수정
- [ ] mod_03: `.Resize(, 5)` → `.Resize(, 6)`
- [ ] mod_03: `Range(1, 3)` → `Range(1, 4)`
- [ ] mod_03: `Range(1, 4/5)` → `Range(1, 5/6)`
- [ ] mod_05: `Cells(i, 4)` → `Cells(i, 5)`
- [ ] VBA 컴파일 성공

## 3단계: Power Query
- [ ] fnProcessSingleFile (연결 전용)
- [ ] 메인주소 (연결 전용)
- [ ] 회사결산_주소 (연결 전용)
- [ ] Raw_Data (연결 전용)
- [ ] TB (연결 전용)
- [ ] PTB (BSPL!$B$4 로드)

## 4단계: 테스트
- [ ] 새로고침 시 데이터 로드
- [ ] CoA 매핑 정상 동작

---

# 참조 문서

| 문서 | 위치 | 내용 |
|------|------|------|
| HRE_SPO_M함수_전체.md | `PowerQuery/` | 단일 파일 처리 M 함수 상세 |
| HRE_PowerQuery_전체.md | `PowerQuery/` | 전체 쿼리 아키텍처 및 M 코드 |
| 완벽한_구현_가이드.md | 루트 | 전체 시스템 구현 매뉴얼 |
| CLAUDE.md | 루트 | 프로젝트 컨텍스트 |

---

**작성일**: 2026-01-24
**버전**: 1.00
