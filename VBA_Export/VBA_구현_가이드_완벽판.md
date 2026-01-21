# VBA 구현 가이드 (완벽판) - HRE 연결마스터

**작성일**: 2026-01-21
**대상**: VBA 초보자도 따라할 수 있도록 단계별 상세 설명
**목표**: 컴파일 에러 없이 연결마스터 실행 가능한 상태로 만들기

---

## 📋 목차

1. [현재 상황 요약](#1-현재-상황-요약)
2. [컴파일 에러 해결 (필수)](#2-컴파일-에러-해결-필수)
3. [한글 인코딩 검증](#3-한글-인코딩-검증)
4. [SharePoint 폴더구조](#4-sharepoint-폴더구조)
5. [Power Query 취합 로직](#5-power-query-취합-로직)
6. [최종 검증 체크리스트](#6-최종-검증-체크리스트)
7. [문제 해결 (Troubleshooting)](#7-문제-해결-troubleshooting)

---

## 1. 현재 상황 요약

### ✅ 완료된 작업
- [x] VBA 파일 21개 → CP949 인코딩 변환 (Windows)
- [x] 15개 시트 코드 복사-붙여넣기 (ThisWorkbook + 14개 worksheet)
- [x] 20개 일반 모듈 파일 가져오기 (File → Import File)
- [x] VBA 코드 정밀 분석 완료 (187개 에러 발견)

### ⚠️ 해결 필요 (현재 컴파일 실패 원인)
- [ ] **Worksheet CodeName 미설정** (187개 참조 실패)
- [ ] mod_11_Sync.bas Option Explicit 추가 (수정 완료, import 필요)

### 📊 에러 통계
| 에러 유형 | 발생 건수 | 심각도 |
|----------|----------|--------|
| Worksheet CodeName 참조 | 187개 | 🔴 Critical |
| Option Explicit 누락 | 1개 | 🟡 Medium |
| API 선언 (32/64비트) | 0개 | ✅ 해결됨 |
| 한글 인코딩 | 0개 | ✅ 해결됨 |

---

## 2. 컴파일 에러 해결 (필수)

### 2-1. 문제 설명

**에러 메시지**: "변수가 정의되지 않았습니다"

**발생 위치**:
```vba
' 현재_통합_문서_code.bas Line 36
HideSheet.Range("N2").Value = AppVersion
' ❌ HideSheet이 정의되지 않아서 실패
```

**원인**: Excel 시트에 VBA CodeName이 설정되지 않음

**영향**:
- 10개 모듈에서 187개 참조 실패
- VBAProject 컴파일 불가능
- 모든 매크로 실행 불가

---

### 2-2. 해결 방법: CodeName 설정 (5분 소요)

#### **Step 1: VBA Editor 열기**
1. Excel 파일 열기: `연결마스터_HRE_v1.00.xlsm`
2. 비밀번호 입력: `PwCDA7529`
3. `Alt + F11` 키 → VBA Editor 실행

#### **Step 2: Properties 창 열기**
1. VBA Editor 상단 메뉴 → **보기(View)** → **속성 창(Properties Window)**
2. 또는 `F4` 키 누르기
3. 좌측 하단에 Properties 창 표시 확인

#### **Step 3: CodeName 설정 (9개 시트)**

좌측 **Project Explorer** 창에서 시트를 선택하고, Properties 창에서 **(Name)** 속성을 아래 표대로 설정:

| 순서 | 시트 탭 이름 (Caption) | Properties 창 설정 |
|-----|----------------------|-------------------|
| 1 | BSPL | **(Name)**: BSPL |
| 2 | AddCoA | **(Name)**: AddCoA |
| 3 | CorpCoA 또는 법인별 CoA | **(Name)**: CorpCoA |
| 4 | CoAMaster 또는 계정 마스터 | **(Name)**: CoAMaster |
| 5 | Check | **(Name)**: Check |
| 6 | Verify 또는 검증 | **(Name)**: Verify |
| 7 | HideSheet | **(Name)**: HideSheet |
| 8 | CorpMaster 또는 법인 마스터 | **(Name)**: CorpMaster |
| 9 | CorpBSPL 또는 법인별BSPL | **(Name)**: CorpBSPL |

**설정 방법 상세:**
```
[Project Explorer에서]
1. "Microsoft Excel Objects" 폴더 확장
2. "Sheet1 (BSPL)" 같은 항목 클릭

[Properties 창에서]
3. 위쪽에 "(Name)" 속성 찾기 (괄호 포함!)
4. 오른쪽 값을 "BSPL"로 변경
5. Enter 키
6. 다음 시트로 이동하여 반복
```

**주의사항:**
- ⚠️ **(Name)** 속성과 **Name** 속성은 다릅니다!
- **(Name)**: VBA CodeName (코드에서 사용)
- **Name**: 시트 탭 이름 (Excel에서 보이는 이름)
- 반드시 **(Name)**을 설정하세요!

#### **Step 4: mod_11_Sync.bas 다시 Import**

1. VBA Editor → 좌측 Project Explorer
2. 기존 **mod_11_Sync** 모듈 찾기
3. 마우스 우클릭 → **모듈 내보내기(Export Module)** → 임시 저장
4. 마우스 우클릭 → **제거(Remove mod_11_Sync)** → "예" 선택
5. **파일(File)** → **파일 가져오기(Import File)**
6. 수정된 `mod_11_Sync.bas` 선택 (Option Explicit 추가됨)
7. 열기

#### **Step 5: 컴파일 검증**

1. VBA Editor 상단 메뉴 → **디버그(Debug)** → **VBAProject 컴파일(Compile VBAProject)**
2. 에러 없이 완료되면 성공! ✅
3. `Ctrl + S` → 저장

---

### 2-3. 컴파일 성공 확인 방법

**성공 시:**
```
[메뉴에 표시]
디버그 → VBAProject 컴파일  (클릭 가능, 메시지 없음)
```

**실패 시:**
```
[에러 메시지 창 표시]
컴파일 오류:
변수가 정의되지 않았습니다

[OK] 버튼 클릭 → 에러 발생 줄로 이동
```

**실패 원인:**
1. CodeName 철자 오류 (대소문자 정확히!)
2. 일부 시트 CodeName 설정 누락
3. mod_11_Sync.bas 업데이트 안됨

---

## 3. 한글 인코딩 검증

### 3-1. 검증이 필요한 이유

- VBA Editor는 **CP949 인코딩만 지원** (Windows 한글 표준)
- macOS에서 내보낸 파일은 **UTF-8 인코딩**
- Windows에서 UTF-8 → CP949 변환 필수
- 변환하지 않으면 한글이 깨져서 표시됨

### 3-2. 검증 방법

**VBA Editor에서 3개 모듈 열어서 확인:**

#### **① mod_10_Public.bas (Line 17)**
```vba
Public Const AppType = "연결마스터"
```
- ✅ "연결마스터" 정상 표시
- ❌ "������" 깨짐 → CP949 변환 필요

#### **② mod_06_VerifySum.bas (Line 10)**
```vba
GoEnd "선행 단계를 완료하세요!"
```
- ✅ "선행 단계를 완료하세요!" 정상 표시
- ❌ 물음표나 이상한 문자 → 변환 필요

#### **③ mod_17_ExchangeRate.bas (Line 133)**
```vba
MsgBox "평균환율 정보가 업데이트되었습니다.", vbInformation
```
- ✅ "평균환율 정보가 업데이트되었습니다." 정상 표시

### 3-3. 한글 깨짐 해결 (Windows 필수)

**도구**: `convert_to_cp949_windows.py`

**실행 방법:**
```powershell
# PowerShell 또는 Git Bash
cd C:\Users\[사용자명]\Desktop\Project\HRE\작업\VBA_Export

# Python 스크립트 실행
python convert_to_cp949_windows.py
```

**변환 대상 파일 (21개):**
- 현재_통합_문서_code.bas
- CoAMaster_code.bas
- mod_01 ~ mod_11 (11개)
- mod_16, mod_17 (2개)
- mod_Log.bas, mod_OpenPage.bas 등 (7개)

**변환 후 재 Import:**
1. VBA Editor에서 기존 모듈 모두 제거
2. 변환된 .bas 파일 다시 Import
3. 시트 코드는 복사-붙여넣기로 재입력

---

## 4. SharePoint 폴더구조

### 4-1. SharePoint 사이트 구조

**기본 URL**: `https://pwckor.sharepoint.com/sites/KR-ASR-HRE_Consolidation`

```
KR-ASR-HRE_Consolidation (Site)
│
├── Documents (문서 라이브러리)
│   │
│   ├── Financial Data (재무 데이터)
│   │   ├── {YYYY}년 {MM}월 (예: 2025년 12월)
│   │   │   ├── PTB (시산표)
│   │   │   │   ├── [법인코드]_합잔시산표_{YYYYMMDD}.xlsx
│   │   │   │   ├── HRE-001_합잔시산표_20251231.xlsx
│   │   │   │   └── HRE-002_합잔시산표_20251231.xlsx
│   │   │   │
│   │   │   ├── BSPL (재무제표)
│   │   │   │   ├── [법인코드]_BSPL_{YYYYMMDD}.xlsx
│   │   │   │   └── HRE-001_BSPL_20251231.xlsx
│   │   │   │
│   │   │   └── Raw Files (원본 파일)
│   │   │       └── [법인명]_원본자료_{YYYYMMDD}.xlsx
│   │   │
│   │   └── Master Data (마스터 데이터)
│   │       ├── CoA_Master.xlsx (계정과목 마스터)
│   │       ├── Corp_Master.xlsx (법인 마스터)
│   │       └── Exchange_Rate.xlsx (환율 정보)
│   │
│   ├── Reports (보고서)
│   │   └── {YYYY}년 {MM}월
│   │       ├── 연결재무제표_{YYYYMMDD}.xlsx
│   │       └── 검증보고서_{YYYYMMDD}.xlsx
│   │
│   └── Templates (템플릿)
│       ├── 합잔시산표_양식.xlsx
│       ├── BSPL_양식.xlsx
│       └── 연결마스터_HRE_v{XX}.xlsm (최신 버전)
│
└── Lists (목록)
    ├── PTB_연결시산표 (Power Query 연결)
    └── Raw_CoA (법인별 CoA)
```

### 4-2. 파일명 규칙

#### **시산표 (PTB)**
```
형식: [법인코드]_합잔시산표_{YYYYMMDD}.xlsx
예시: HRE-001_합잔시산표_20251231.xlsx

필수 컬럼:
- 법인코드 (예: HRE-001)
- 계정코드 (5자리, 예: 10300)
- 계정과목명 (예: 현금및현금성자산)
- 차변 (숫자)
- 대변 (숫자)
```

#### **재무제표 (BSPL)**
```
형식: [법인코드]_BSPL_{YYYYMMDD}.xlsx
예시: HRE-001_BSPL_20251231.xlsx

필수 시트:
- BS (재무상태표)
- PL (손익계산서)
```

#### **법인별 CoA 매핑**
```
형식: [법인코드]_CoA_Mapping_{YYYYMMDD}.xlsx

필수 컬럼:
- 법인코드
- 계정코드 (5자리)
- 계정과목명
- PwC_CoA (6자리, 예: 110100)
- PwC_계정과목명
- Variant Type (BASE / INTERCO_KR / INTERCO_IC)
```

### 4-3. Power Query 연결 설정

**Excel에서 설정 방법:**

1. **데이터(Data)** 탭 → **데이터 가져오기(Get Data)** → **SharePoint에서(From SharePoint Folder)**
2. SharePoint 사이트 URL 입력:
   ```
   https://pwckor.sharepoint.com/sites/KR-ASR-HRE_Consolidation
   ```
3. Office 365 계정으로 로그인
4. 연결할 목록 선택: `PTB_연결시산표`, `Raw_CoA`
5. 테이블 변환 → 필요한 컬럼 선택 → 로드

**연결된 쿼리 확인:**
- BSPL 시트 → `PTB` 테이블 → 쿼리 속성에서 연결 정보 확인
- CorpCoA 시트 → `Raw_CoA` 테이블 → 쿼리 속성 확인

---

## 5. Power Query 취합 로직

### 5-1. PTB (시산표) 취합 프로세스

**개념**: 여러 법인의 Excel 파일을 하나의 테이블로 통합

#### **Step 1: SharePoint 폴더 스캔**
```m
// Power Query M Language
let
    // SharePoint 사이트 연결
    Source = SharePoint.Tables(
        "https://pwckor.sharepoint.com/sites/KR-ASR-HRE_Consolidation",
        [ApiVersion = 15]
    ),

    // PTB_연결시산표 목록 선택
    PTBList = Source{[Name="PTB_연결시산표"]}[Items],
```

#### **Step 2: 필수 컬럼 선택**
```m
    // 필요한 컬럼만 선택 (성능 최적화)
    SelectedColumns = Table.SelectColumns(PTBList, {
        "법인코드",      // 예: HRE-001
        "계정코드",      // 예: 10300
        "계정과목명",    // 예: 현금및현금성자산
        "차변",          // 숫자
        "대변"           // 숫자
    }),
```

#### **Step 3: 계산 컬럼 추가**
```m
    // 차변-대변 계산
    AddNetBalance = Table.AddColumn(
        SelectedColumns,
        "차변-대변",
        each [차변] - [대변],
        type number
    ),

    // PwC CoA 매핑 컬럼 추가 (VBA에서 채움)
    AddPwCCoA = Table.AddColumn(
        AddNetBalance,
        "PwC_CoA",
        each null,
        type text
    ),

    AddPwCName = Table.AddColumn(
        AddPwCCoA,
        "PwC_계정과목명",
        each null,
        type text
    ),
```

#### **Step 4: 데이터 타입 강제**
```m
    // 데이터 타입 명시적 설정
    ChangedTypes = Table.TransformColumnTypes(AddPwCName, {
        {"법인코드", type text},
        {"계정코드", type text},
        {"계정과목명", type text},
        {"차변", Currency.Type},
        {"대변", Currency.Type},
        {"차변-대변", Currency.Type}
    })
in
    ChangedTypes
```

### 5-2. VBA와 Power Query 연동

**mod_05_PTB_Highlight.bas: QueryRefresh()**
```vba
Sub QueryRefresh()
    Dim tblPTB As ListObject

    ' BSPL 시트의 PTB 테이블 참조
    Set tblPTB = BSPL.ListObjects("PTB")

    ' Power Query 새로고침 (SharePoint에서 최신 데이터 가져옴)
    tblPTB.QueryTable.Refresh BackgroundQuery:=False

    ' 비동기 쿼리 완료 대기
    Application.CalculateUntilAsyncQueriesDone

    MsgBox "SPO로부터 데이터 새로고침 완료", vbInformation
End Sub
```

### 5-3. First Drafting (자동 CoA 매핑)

**mod_03_PTB_CoA_Input.bas: Fill_Input_Table()**

#### **알고리즘 흐름:**
```
1. PTB 테이블에서 노란색 행만 필터링
   └─> 아직 PwC_CoA가 매핑되지 않은 계정

2. Raw_CoA 테이블에서 Dictionary 생성
   └─> Key: 계정코드 (5자리 base code)
   └─> Value: {Variant Type: PwC_CoA}

3. 각 PTB 행에 대해:
   a) 계정코드에서 5자리 추출 (예: 11401_내부거래 → 11401)
   b) Variant Type 감지 (_내부거래 → INTERCO_KR)
   c) Dictionary 조회:
      - 정확한 Variant 매칭 우선
      - 없으면 BASE Variant 사용
      - 없으면 빈 칸 (수동 입력 필요)

4. CoA_Input 테이블에 결과 채움
```

#### **Variant Detection 로직:**
```vba
Private Function GetVariantType(accountCode As String) As String
    If InStr(accountCode, "_내부거래") > 0 Then
        GetVariantType = "INTERCO_KR"     ' 내부거래 (한국)
    ElseIf InStr(accountCode, "_IC") > 0 Then
        GetVariantType = "INTERCO_IC"     ' 내부거래 (그룹간)
    ElseIf Left(accountCode, 2) = "MC" Then
        GetVariantType = "CONSOLIDATION"  ' 연결조정
    Else
        GetVariantType = "BASE"           ' 일반 계정
    End If
End Function
```

#### **5-Digit Matching:**
```vba
Private Function GetBaseCode(accountCode As String) As String
    Dim baseCode As String
    baseCode = accountCode

    ' Variant 접미사 제거
    If InStr(baseCode, "_") > 0 Then
        baseCode = Left(baseCode, InStr(baseCode, "_") - 1)
    End If

    ' 앞 5자리만 추출 (HRE 표준)
    If Len(baseCode) >= 5 Then
        GetBaseCode = Left(baseCode, 5)   ' 예: 11401
    Else
        GetBaseCode = baseCode
    End If
End Function
```

### 5-4. 성능 최적화 전략

**Array-Based Operations:**
```vba
' ❌ 느린 방식 (셀 하나씩 접근)
For Each row In tblPTB.ListRows
    Debug.Print row.Range(1, 1).Value
Next row

' ✅ 빠른 방식 (배열로 한번에 로드)
Dim dataArray() As Variant
dataArray = tblPTB.DataBodyRange.Value

For i = 1 To UBound(dataArray)
    Debug.Print dataArray(i, 1)
Next i
```

**Dictionary Lookup:**
```vba
' ❌ 느린 방식 (중첩 루프, O(n²))
For Each ptbRow In tblPTB.ListRows
    For Each rawRow In tblRawCoA.ListRows
        If ptbRow.Range(1, 2) = rawRow.Range(1, 2) Then
            ' 매칭 처리
        End If
    Next rawRow
Next ptbRow

' ✅ 빠른 방식 (Dictionary, O(n))
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")

For Each rawRow In tblRawCoA.ListRows
    dict(rawRow.Range(1, 2).Value) = rawRow.Range(1, 5).Value
Next rawRow

For Each ptbRow In tblPTB.ListRows
    If dict.Exists(ptbRow.Range(1, 2).Value) Then
        ptbRow.Range(1, 4).Value = dict(ptbRow.Range(1, 2).Value)
    End If
Next ptbRow
```

---

## 6. 최종 검증 체크리스트

### 6-1. VBA 환경 검증

- [ ] **CodeName 설정 완료 (9개 시트)**
  - [ ] BSPL
  - [ ] AddCoA
  - [ ] CorpCoA
  - [ ] CoAMaster
  - [ ] Check
  - [ ] Verify
  - [ ] HideSheet
  - [ ] CorpMaster
  - [ ] CorpBSPL

- [ ] **모듈 Import 완료 (20개)**
  - [ ] mod_01_FilterSearch.bas
  - [ ] mod_02_FilterSearch_Master.bas
  - [ ] mod_03_PTB_CoA_Input.bas
  - [ ] mod_04_IntializeProgress.bas
  - [ ] mod_05_PTB_Highlight.bas
  - [ ] mod_06_VerifySum.bas
  - [ ] mod_09_CheckMaster.bas
  - [ ] mod_10_Public.bas
  - [ ] mod_11_Sync.bas (Option Explicit 추가 버전)
  - [ ] mod_16_Export.bas
  - [ ] mod_17_ExchangeRate.bas
  - [ ] mod_Log.bas
  - [ ] mod_MouseWheel.bas
  - [ ] mod_OpenPage.bas
  - [ ] mod_QueryProtection.bas
  - [ ] mod_Refresh.bas
  - [ ] mod_Ribbon.bas
  - [ ] mod_z_Module_GetCursor.bas
  - [ ] Module1.bas
  - [ ] Setup_CoAMaster.bas

- [ ] **시트 코드 복사-붙여넣기 완료 (15개)**
  - [ ] 현재_통합_문서_code.bas → ThisWorkbook
  - [ ] CoAMaster_code.bas → CoAMaster 시트
  - [ ] 나머지 13개 시트 (ADBS, AddCoA_ADBS, AddCoA, BSPL, Check, CorpCoA, CorpMaster, DirectoryURL, Guide, HideSheet, Memo, Verify 등)

- [ ] **컴파일 검증**
  - [ ] 디버그 → VBAProject 컴파일 → 에러 없음

### 6-2. 한글 텍스트 검증

- [ ] mod_10_Public.bas Line 17: `"연결마스터"` 정상 표시
- [ ] mod_06_VerifySum.bas Line 10: `"선행 단계를 완료하세요!"` 정상 표시
- [ ] mod_17_ExchangeRate.bas Line 133: `"평균환율 정보가 업데이트되었습니다."` 정상 표시

### 6-3. 기능 테스트

- [ ] **SPO 연결 설정**
  - [ ] 연결마스터 탭 → SPO 설정 버튼 클릭
  - [ ] SharePoint URL 입력: `https://pwckor.sharepoint.com/sites/KR-ASR-HRE_Consolidation`
  - [ ] HideSheet 시트 E2 셀에 저장 확인

- [ ] **결산연월 설정**
  - [ ] 연결마스터 탭 → 결산연월 설정 버튼
  - [ ] 연도/월 입력 (예: 2025년 12월)
  - [ ] Check 시트 상태: "Complete"

- [ ] **CoA 확인 및 데이터 합산**
  - [ ] 연결마스터 탭 → CoA 확인 및 데이터 합산 버튼
  - [ ] Query 새로고침 실행
  - [ ] PTB 테이블에 데이터 로드 확인
  - [ ] 노란색 강조 (미매핑 계정) 확인

- [ ] **환율 조회**
  - [ ] 연결마스터 탭 → 평균환율 조회 버튼
  - [ ] 날짜 선택 (시작일 ~ 종료일)
  - [ ] "환율정보(평균)" 시트 생성 확인
  - [ ] KEB Hana Bank 데이터 확인

### 6-4. 문서화 검증

- [ ] **CLAUDE.md** - AI 어시스턴트용 가이드 존재
- [ ] **README.md** - 사용자 가이드 존재
- [ ] **Windows_변환_가이드.md** - CP949 변환 가이드 존재
- [ ] **VBA_구현_가이드_완벽판.md** - 본 문서 (이 파일)
- [ ] **_archive/** 폴더 - 중복 문서 정리

---

## 7. 문제 해결 (Troubleshooting)

### 7-1. 컴파일 에러: "변수가 정의되지 않았습니다"

**증상:**
```
컴파일 오류:
변수가 정의되지 않았습니다

HideSheet.Range("N2").Value = AppVersion
```

**원인**: CodeName 설정 누락 또는 철자 오류

**해결 방법:**
1. F4 → Properties 창 열기
2. 에러 발생한 CodeName 확인 (예: HideSheet)
3. Project Explorer에서 해당 시트 선택
4. Properties 창 → **(Name)** 속성 정확히 입력
5. 대소문자 정확히 확인!

**체크리스트:**
- [ ] **(Name)** 속성인지 확인 (Name이 아님!)
- [ ] 철자 정확한지 확인 (예: HideSheet vs HIdEsHeEt)
- [ ] 9개 시트 모두 설정했는지 확인

---

### 7-2. 한글 깨짐: "������"

**증상:**
```vba
Public Const AppType = "������"  ' ❌ 깨짐
```

**원인**: UTF-8 → CP949 변환 안됨

**해결 방법:**
1. VBA Editor 닫기
2. Windows PowerShell 열기
3. `convert_to_cp949_windows.py` 실행
4. 변환된 파일 다시 Import
5. 시트 코드는 복사-붙여넣기

**절대 하지 말 것:**
- ❌ macOS Terminal에서 `iconv` 명령어 사용 (한글 누락 발생)
- ❌ VS Code에서 인코딩 변환 (불완전)
- ❌ 수동으로 한글 수정 (다음 Import 시 다시 깨짐)

---

### 7-3. Power Query 새로고침 실패

**증상:**
```
데이터 원본에 연결할 수 없습니다.
```

**원인**: SharePoint 인증 만료 또는 URL 오류

**해결 방법:**

**Step 1: 연결 재인증**
1. 데이터 탭 → 쿼리 및 연결
2. 쿼리 우클릭 → 속성
3. 정의 편집
4. 홈 탭 → 고급 편집기
5. SharePoint URL 확인:
   ```
   https://pwckor.sharepoint.com/sites/KR-ASR-HRE_Consolidation
   ```
6. 적용 → Office 365 계정 재로그인

**Step 2: 로컬 파일 대안**
```m
let
    Source = Excel.Workbook(
        File.Contents("C:\Users\...\합잔시산표_통합.xlsx"),
        null,
        true
    ),
    Sheet = Source{[Item="시산표",Kind="Sheet"]}[Data],
    // ... 나머지 변환 로직 동일
in
    ChangedTypes
```

---

### 7-4. Ribbon 메뉴 ("연결마스터" 탭) 안 보임

**증상:**
- Excel 상단에 "연결마스터" 탭이 표시되지 않음

**원인**: CustomUI XML 미적용 또는 손상

**해결 방법:**

**Option 1: Custom UI Editor 사용 (권장)**
1. [Custom UI Editor](https://github.com/OfficeDev/office-custom-ui-editor) 다운로드
2. `연결마스터_HRE_v1.00.xlsm` 열기
3. Insert → Office 2010 Custom UI Part
4. Ribbon XML 붙여넣기 (CLAUDE.md 참조)
5. Validate XML (초록색 체크)
6. Save → Excel 재시작

**Option 2: VBA 함수로 확인**
```vba
' Immediate 창에서 실행
? Application.CommandBars("연결마스터").Visible
' True면 존재, False면 없음
```

---

### 7-5. "파일이 손상되어 열 수 없습니다"

**증상:**
```
Excel에서 'file.xlsm' 파일을 열 수 없습니다.
파일 형식 또는 파일 확장명이 잘못되었습니다.
```

**원인**:
- Ribbon XML 문법 오류
- VBA 모듈 손상
- ZIP 압축 해제 실패

**해결 방법:**

**Step 1: 백업 복원**
```bash
# Git에서 이전 버전 복원
git checkout HEAD~1 연결마스터_HRE_v1.00.xlsm
```

**Step 2: Excel 복구 모드**
1. Excel 열기 (빈 화면)
2. 파일 → 열기 → `연결마스터_HRE_v1.00.xlsm` 선택
3. 열기 버튼 옆 화살표 → **열기 및 복구(Open and Repair)**
4. 복구 실행

**Step 3: XML 검증**
```bash
# PowerShell에서
Rename-Item "연결마스터_HRE_v1.00.xlsm" "file.zip"
Expand-Archive "file.zip" -DestinationPath "temp"
notepad "temp\customUI\customUI14.xml"  # XML 문법 확인
```

---

### 7-6. 매크로 실행 시 "권한 없음" 에러

**증상:**
```
실행 시간 오류 '70':
사용 권한이 거부되었습니다.
```

**원인**:
- 시트 보호 해제 안됨
- 비밀번호 오류

**해결 방법:**

**전역 변수 확인:**
```vba
' mod_10_Public.bas
Public Const PASSWORD As String = "BEP1234"           ' 시트 비밀번호
Public Const PASSWORD_Workbook As String = "PwCDA7529" ' 통합문서 비밀번호
```

**수동 해제:**
1. 시트 탭 우클릭 → 시트 보호 해제
2. 비밀번호: `BEP1234` 입력
3. 매크로 재실행

---

## 8. Git 커밋 체크리스트

### 8-1. 커밋 전 최종 확인

- [ ] 모든 VBA 파일 CP949 인코딩 완료
- [ ] mod_11_Sync.bas Option Explicit 추가
- [ ] 중복 문서 _archive/ 폴더로 이동
- [ ] 본 가이드 문서 작성 완료
- [ ] SharePoint 폴더구조 명시 완료
- [ ] Power Query 로직 설명 완료

### 8-2. 커밋 명령어

```bash
# 작업 디렉토리로 이동
cd /Users/jaewookim/Desktop/Project/HRE/작업/VBA_Export

# 상태 확인
git status

# 변경 파일 스테이징
git add mod_11_Sync.bas
git add VBA_구현_가이드_완벽판.md

# 커밋
git commit -m "Fix: mod_11_Sync Option Explicit 추가 + VBA 구현 가이드 완성

- mod_11_Sync.bas: Line 2에 Option Explicit 추가
- VBA_구현_가이드_완벽판.md: VBA 초보자용 완벽 구현 가이드 작성
  * CodeName 설정 단계별 가이드 (9개 시트)
  * SharePoint 폴더구조 및 파일명 규칙 명시
  * Power Query 취합 로직 상세 설명
  * First Drafting 알고리즘 설명
  * Troubleshooting 7가지 시나리오

Co-Authored-By: Claude Sonnet 4.5 <noreply@anthropic.com>"

# 원격 저장소에 푸시
git push origin main
```

---

## 9. 다음 단계 (Next Steps)

### 9-1. 즉시 실행 (5분)
1. [ ] Excel 파일 열기
2. [ ] Alt+F11 → VBA Editor
3. [ ] F4 → Properties 창
4. [ ] 9개 시트 CodeName 설정
5. [ ] mod_11_Sync.bas 다시 Import
6. [ ] Ctrl+S → 저장
7. [ ] 디버그 → VBAProject 컴파일

### 9-2. 기능 테스트 (10분)
1. [ ] SPO 설정
2. [ ] 결산연월 설정
3. [ ] Query 새로고침
4. [ ] 환율 조회
5. [ ] CoA First Drafting

### 9-3. 문서화 완성 (30분)
1. [ ] README.md 업데이트
2. [ ] CLAUDE.md 검토
3. [ ] 스크린샷 추가 (Properties 창)
4. [ ] Git 커밋 & 푸시

---

## 10. 참고 자료

### 10-1. VBA CodeName vs TabName

| 속성 | 위치 | 용도 | 예시 |
|-----|------|------|------|
| CodeName | Properties 창 (Name) | VBA 코드에서 참조 | `HideSheet.Range("A1")` |
| TabName | Properties 창 Name | Excel 탭 이름 | 시트 탭에 "숨김시트" 표시 |

**코드 예시:**
```vba
' ✅ CodeName 사용 (추천)
HideSheet.Range("A1").Value = "Hello"  ' 직접 참조 (빠름, 타입 안전)

' ⚠️ TabName 사용 (대안)
ThisWorkbook.Worksheets("숨김시트").Range("A1").Value = "Hello"  ' 문자열 검색 (느림)
```

### 10-2. CP949 vs UTF-8

| 특성 | CP949 | UTF-8 |
|-----|-------|-------|
| 지원 범위 | 한글 2,350자 | 전 세계 모든 문자 |
| 한글 크기 | 2바이트 | 3바이트 |
| VBA Editor | ✅ 지원 | ❌ 미지원 |
| macOS Excel | ❌ 미지원 | ✅ 기본값 |
| Windows Excel | ✅ 기본값 | ❌ 미지원 |

### 10-3. 유용한 단축키

| 단축키 | 기능 |
|--------|------|
| Alt+F11 | VBA Editor 열기/닫기 |
| F4 | Properties 창 열기 |
| F5 | 매크로 실행 |
| F7 | Code 창 열기 |
| Ctrl+G | Immediate 창 열기 |
| Ctrl+R | Project Explorer 열기 |
| Ctrl+Break | 실행 중단 |

---

## 작성 이력

| 버전 | 날짜 | 작성자 | 변경 내역 |
|------|------|--------|----------|
| 1.0 | 2026-01-21 | Claude Code | 초안 작성 (CodeName 설정 가이드) |
| 1.1 | 2026-01-21 | Claude Code | SharePoint 폴더구조 추가 |
| 1.2 | 2026-01-21 | Claude Code | Power Query 로직 상세 설명 |
| 1.3 | 2026-01-21 | Claude Code | Troubleshooting 7가지 시나리오 추가 |

---

**© 2026 Samil PwC. All rights reserved.**
**문의**: jaewookim@pwc.com (가상 주소)
