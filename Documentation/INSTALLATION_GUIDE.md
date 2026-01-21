# HRE 통합결산 시스템 - 설치 가이드

**Version**: 2.0
**Date**: 2026-01-21
**Author**: PwC Digital Assurance Team

---

## 목차
1. [Ribbon XML 설치](#1-ribbon-xml-설치)
2. [Power Query 설정](#2-power-query-설정)
3. [VBA 모듈 준비](#3-vba-모듈-준비)
4. [검증 및 테스트](#4-검증-및-테스트)
5. [문제 해결](#5-문제-해결)

---

## 1. Ribbon XML 설치

### 1.1 Custom UI Editor 다운로드
1. Custom UI Editor for Microsoft Office 다운로드
   - URL: https://bettersolutions.com/vba/ribbon/custom-ui-editor.htm
   - 또는 OpenXMLDeveloper 버전 사용 가능
2. 설치 후 실행

### 1.2 Ribbon XML 추가

#### 방법 A: Custom UI Editor 사용 (권장)
```
1. Custom UI Editor 실행
2. 메뉴: File > Open
3. HRE 통합결산.xlsm 파일 열기
4. 메뉴: Insert > Office 2010 Custom UI Part
5. 좌측 트리에서 "customUI14.xml" 선택
6. 우측 편집 영역에 `/Users/jaewookim/Desktop/Project/HRE/작업/customUI14.xml` 내용 복사
7. 메뉴: File > Save
8. Custom UI Editor 종료
9. Excel에서 파일 열어서 "HRE 통합결산" 탭 확인
```

#### 방법 B: 수동 ZIP 편집 (고급 사용자)
```
1. .xlsm 파일을 .zip으로 확장자 변경
2. ZIP 압축 해제
3. customUI 폴더 생성 (루트 레벨)
4. customUI14.xml 파일을 customUI 폴더에 복사
5. _rels 폴더 내 .rels 파일 편집:
   <Relationship Id="rId1"
                 Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
                 Target="customUI/customUI14.xml"/>
6. [Content_Types].xml 편집:
   <Override PartName="/customUI/customUI14.xml"
             ContentType="application/xml"/>
7. 모든 파일을 다시 ZIP으로 압축
8. .zip을 .xlsm으로 확장자 변경
```

### 1.3 Ribbon 동작 확인
```vba
' Excel 열기 > HRE 통합결산 탭 확인
' 7개 그룹 23개 버튼이 표시되어야 함:
' - 설정 (5)
' - 환율 (2) ← NEW
' - 계정과목 (2)
' - 검증 (3)
' - 도구 (6)
' - 내보내기 (1)
' - 정보 (4)
```

---

## 2. Power Query 설정

### 2.1 PTB Query 설정 (SharePoint 연결)

#### Step 1: 새 쿼리 생성
```
1. Excel 메뉴: 데이터 > 쿼리 및 연결 > 쿼리 편집
2. 새 쿼리: 새 원본 > 기타 원본 > 빈 쿼리
3. 고급 편집기 열기
```

#### Step 2: M 코드 붙여넣기
```
1. /Users/jaewookim/Desktop/Project/HRE/작업/PowerQuery_PTB.m 파일 열기
2. 전체 내용 복사
3. 고급 편집기에 붙여넣기
```

#### Step 3: 연결 정보 수정
```m
// 수정 필요한 부분:
SharePointSiteUrl = "https://pwckorea.sharepoint.com/sites/HRE_Consolidation"
ListName = "PTB_Data"  // 실제 SharePoint 리스트 이름으로 변경
```

#### Step 4: 권한 설정
```
1. 데이터 원본 설정 창이 나타나면
2. "조직 계정" 선택
3. Office 365 계정으로 로그인
4. 연결 > 확인
```

#### Step 5: 쿼리 이름 변경 및 로드
```
1. 쿼리 이름: "PTB"로 변경
2. 닫기 및 로드 > 닫기 및 로드 대상...
3. "테이블" 선택
4. 워크시트: "PTB" (새 워크시트)
5. 확인
```

### 2.2 Raw_CoA Query 설정 (SharePoint 연결)

#### Step 1-4: PTB와 동일하게 진행
```
파일: /Users/jaewookim/Desktop/Project/HRE/작업/PowerQuery_RawCoA.m
쿼리 이름: "Raw_CoA"
워크시트: "Raw_CoA" (새 워크시트)
```

#### SharePoint 리스트 설정 수정:
```m
SharePointSiteUrl = "https://pwckorea.sharepoint.com/sites/HRE_Consolidation"
ListName = "Raw_CoA_History"  // 실제 리스트 이름으로 변경
```

### 2.3 로컬 파일 Query 설정 (테스트용)

#### Step 1: 파일 경로 확인
```
실제 파일 위치:
/Users/jaewookim/Desktop/Project/HRE/참고/12월 말 합계잔액시산표_에이치알이_수정완료(01.13).xls
```

#### Step 2: M 코드 수정
```m
// /Users/jaewookim/Desktop/Project/HRE/작업/PowerQuery_PTB_LocalFile.m
FilePath = "/Users/jaewookim/Desktop/Project/HRE/참고/12월 말 합계잔액시산표_에이치알이_수정완료(01.13).xls"
SheetName = "Sheet1"  // 실제 시트 이름 확인 후 변경
HeaderRow = 1
```

#### Step 3: 로드
```
쿼리 이름: "PTB_Local"
워크시트: "PTB_Test" (새 워크시트)
```

### 2.4 Query 연결만 생성 (데이터 로드 안 함)

#### 배경 새로 고침 설정:
```
1. 쿼리 우클릭 > 속성
2. "백그라운드에서 새로 고침 사용" 체크
3. "통합 문서를 열 때 새로 고침" 체크 해제 (수동 새로 고침)
```

---

## 3. VBA 모듈 준비

### 3.1 필수 Callback 프로시저 목록

`mod_Ribbon.bas` 파일에 다음 프로시저가 있어야 합니다:

```vba
' ========== 설정 그룹 ==========
Public Sub SetSPO_OnAction(control As IRibbonControl)
    ' SharePoint 연결 설정 폼 표시
    frmSPO.Show
End Sub

Public Sub SetDirectory_OnAction(control As IRibbonControl)
    ' 디렉토리 설정 폼 표시
    frmDirectory.Show
End Sub

Public Sub SetDate_OnAction(control As IRibbonControl)
    ' 결산연월 설정 폼 표시
    frmSetDate.Show
End Sub

Public Sub AppendCorp_OnAction(control As IRibbonControl)
    ' 법인 추가 폼 표시
    frmAppendCorp.Show
End Sub

Public Sub SetScope_OnAction(control As IRibbonControl)
    ' 결산범위 설정 폼 표시
    frmSetScope.Show
End Sub

' ========== 환율 그룹 (NEW) ==========
Public Sub GetER_Flow_OnAction(control As IRibbonControl)
    ' 흐름환율 조회 (외환은행 API 연동)
    Call GetExchangeRate_Flow
End Sub

Public Sub GetER_Spot_OnAction(control As IRibbonControl)
    ' 마감환율 조회 (외환은행 API 연동)
    Call GetExchangeRate_Spot
End Sub

' ========== 계정과목 그룹 ==========
Public Sub Update_OnAction(control As IRibbonControl)
    ' CoA Master 업데이트
    Call UpdateCoAMaster
End Sub

Public Sub Synchro_CoA_OnAction(control As IRibbonControl)
    ' CoA 동기화 (mod_11_Sync.bas)
    Call SynchronizeCoA
End Sub

' ========== 검증 그룹 ==========
Public Sub Verify_BSPL_OnAction(control As IRibbonControl)
    ' 재무제표 검증 (mod_06_VerifySum.bas)
    Call VerifyBSPL
End Sub

Public Sub Verify_AD_OnAction(control As IRibbonControl)
    ' 취득처분 검증
    Call VerifyAD
End Sub

Public Sub Verify_Master_OnAction(control As IRibbonControl)
    ' Master 검증
    Call VerifyMaster
End Sub

' ========== 도구 그룹 ==========
Public Sub FilterSheet_OnAction(control As IRibbonControl)
    ' 현재 시트 필터 적용
    Call ApplyAutoFilter
End Sub

Public Sub UnfilterSheet_OnAction(control As IRibbonControl)
    ' 필터 해제
    Call ClearAutoFilter
End Sub

Public Sub ProtectQuery_OnAction(control As IRibbonControl)
    ' Power Query 보호
    Call ProtectQueries
End Sub

Public Sub UnprotectQuery_OnAction(control As IRibbonControl)
    ' Power Query 보호 해제
    Call UnprotectQueries
End Sub

Public Sub ManagePeople_OnAction(control As IRibbonControl)
    ' 사용자 관리 폼
    frmManagePeople.Show
End Sub

Public Sub Refresh_Data_OnAction(control As IRibbonControl)
    ' 전체 데이터 새로고침
    Call RefreshAllData
End Sub

' ========== 내보내기 그룹 ==========
Public Sub Export_File_OnAction(control As IRibbonControl)
    ' 파일 내보내기 (mod_16_Export.bas)
    Call ExportFile
End Sub

' ========== 정보 그룹 ==========
Public Sub IRVersion_OnAction(control As IRibbonControl)
    ' 버전 정보 표시
    Call ShowVersionInfo
End Sub

Public Sub IRSPO_OnAction(control As IRibbonControl)
    ' SPO 연결 정보 표시
    Call ShowSPOInfo
End Sub

Public Sub IRManual_OnAction(control As IRibbonControl)
    ' 매뉴얼 열기
    Call OpenManual
End Sub

Public Sub IRBugReport_OnAction(control As IRibbonControl)
    ' 버그 리포트 폼
    frmBugReport.Show
End Sub
```

### 3.2 신규 환율 조회 모듈 생성

`mod_17_ExchangeRate.bas` 파일 생성:

```vba
Option Explicit

' ========================================
' 모듈: mod_17_ExchangeRate
' 설명: 외환은행 API를 통한 환율 자동 조회
' 작성일: 2026-01-21
' ========================================

Private Const API_URL_BASE As String = "https://www.koreaexim.go.kr/site/program/financial/exchangeJSON"
Private Const API_KEY As String = "[발급받은_API_키]"  ' TODO: 외환은행에서 발급받은 키 입력

Public Sub GetExchangeRate_Flow()
    ' 흐름환율(평균환율) 조회
    ' 손익계산서 항목에 적용
    On Error GoTo ErrHandler

    Call SpeedUp

    Dim targetDate As String
    Dim currencyList As Variant
    Dim i As Long

    ' 결산연월에서 기간 계산
    targetDate = Format(GetClosingMonth, "YYYYMM")

    ' 조회 대상 통화 (예: USD, EUR, JPY, CNY)
    currencyList = Array("USD", "EUR", "JPY(100)", "CNY")

    For i = LBound(currencyList) To UBound(currencyList)
        Call FetchAndStoreExchangeRate(currencyList(i), targetDate, "평균")
    Next i

    MsgBox "흐름환율 조회가 완료되었습니다.", vbInformation

    Call SpeedDown
    Exit Sub

ErrHandler:
    Call SpeedDown
    MsgBox "흐름환율 조회 중 오류가 발생했습니다: " & Err.Description, vbCritical
End Sub

Public Sub GetExchangeRate_Spot()
    ' 마감환율(기말환율) 조회
    ' 재무상태표 항목에 적용
    On Error GoTo ErrHandler

    Call SpeedUp

    Dim targetDate As String
    Dim currencyList As Variant
    Dim i As Long

    ' 결산연월 마지막 날짜
    targetDate = Format(GetClosingMonth, "YYYYMMDD")

    currencyList = Array("USD", "EUR", "JPY(100)", "CNY")

    For i = LBound(currencyList) To UBound(currencyList)
        Call FetchAndStoreExchangeRate(currencyList(i), targetDate, "기말")
    Next i

    MsgBox "마감환율 조회가 완료되었습니다.", vbInformation

    Call SpeedDown
    Exit Sub

ErrHandler:
    Call SpeedDown
    MsgBox "마감환율 조회 중 오류가 발생했습니다: " & Err.Description, vbCritical
End Sub

Private Sub FetchAndStoreExchangeRate(currency As String, searchDate As String, rateType As String)
    ' API 호출 및 환율 데이터 저장
    ' TODO: HTTP 요청 로직 구현
    ' XMLHTTP 또는 WinHttp.WinHttpRequest 사용
End Sub
```

---

## 4. 검증 및 테스트

### 4.1 Ribbon 버튼 테스트

```
1. Excel 열기
2. "HRE 통합결산" 탭 클릭
3. 각 그룹별 버튼 클릭하여 오류 확인:
   - 설정 그룹: 5개 버튼 모두 폼/기능 실행 확인
   - 환율 그룹: 2개 버튼 (오류 발생 시 모듈 미구현 확인)
   - 계정과목 그룹: 2개 버튼 실행 확인
   - 검증 그룹: 3개 버튼 실행 확인
   - 도구 그룹: 6개 버튼 실행 확인
   - 내보내기: 1개 버튼 실행 확인
   - 정보: 4개 버튼 실행 확인
```

### 4.2 Power Query 테스트

```
1. 데이터 > 쿼리 및 연결
2. PTB 쿼리 우클릭 > 새로 고침
3. 데이터 로드 확인:
   - 행 개수 확인
   - PwC_CoA 컬럼 null 확인 (초기 상태)
   - 매핑상태 = "미완료" 확인

4. Raw_CoA 쿼리 새로 고침
5. 데이터 로드 확인:
   - 변동성유형 값 확인 ("고정", "변동")
   - 신뢰도점수 계산 확인
   - 최신매핑여부 확인
```

### 4.3 통합 워크플로 테스트

```
1. SPO 연결 설정 (Ribbon > 설정 > SPO 연결)
2. 결산연월 설정 (Ribbon > 설정 > 결산연월 설정)
3. 데이터 새로고침 (Ribbon > 도구 > 데이터 새로고침)
4. CoA 동기화 (Ribbon > 계정과목 > CoA 동기화)
5. 재무제표 검증 (Ribbon > 검증 > 재무제표 검증)
6. 파일 내보내기 (Ribbon > 내보내기 > 파일 내보내기)
```

---

## 5. 문제 해결

### 5.1 Ribbon 관련 오류

#### 오류: "리본을 로드할 수 없습니다"
```
원인: XML 문법 오류
해결:
1. customUI14.xml 파일을 XML 검증기로 확인
2. 닫는 태그 누락 확인
3. 특수문자 인코딩 확인 (& → &amp;)
```

#### 오류: "콜백 프로시저를 찾을 수 없습니다"
```
원인: mod_Ribbon.bas에 프로시저 없음
해결:
1. VBA 편집기 열기
2. mod_Ribbon.bas 모듈 확인
3. 위 3.1 섹션의 프로시저 모두 존재하는지 확인
4. Public Sub으로 선언되어 있는지 확인
```

### 5.2 Power Query 관련 오류

#### 오류: "데이터 원본에 연결할 수 없습니다"
```
원인: SharePoint URL 또는 인증 문제
해결:
1. SharePoint 사이트 URL 확인 (브라우저에서 접속 테스트)
2. 데이터 > 데이터 원본 설정 > 권한 편집
3. "조직 계정" 다시 로그인
4. 방화벽/프록시 설정 확인
```

#### 오류: "열을 찾을 수 없습니다"
```
원인: SharePoint 리스트 스키마 불일치
해결:
1. SharePoint 리스트 열 이름 확인
2. PowerQuery M 코드의 STEP 3 (SelectedColumns) 수정
3. 열 이름 정확히 일치하도록 수정 (대소문자, 공백 주의)
```

#### 오류: "데이터 형식 변환 실패"
```
원인: 데이터 타입 불일치
해결:
1. STEP 4 (TypedColumns)에서 데이터 타입 확인
2. 원본 데이터 샘플 확인
3. try ... otherwise null 구문으로 안전하게 변환:
   Table.TransformColumnTypes(..., [{"차변", type nullable number}])
```

### 5.3 VBA 실행 오류

#### 오류: "개체 변수 또는 With 블록 변수가 설정되지 않았습니다"
```
원인: ListObject 또는 Worksheet 참조 오류
해결:
1. PTB 쿼리가 로드되어 테이블이 생성되었는지 확인
2. Worksheets("PTB").ListObjects("PTB") 존재 확인
3. 즉시 실행 창에서 테스트:
   ?Worksheets("PTB").ListObjects(1).Name
```

#### 오류: "환율 조회 실패"
```
원인: API 키 미설정 또는 네트워크 오류
해결:
1. mod_17_ExchangeRate.bas의 API_KEY 상수 확인
2. 외환은행 오픈API 키 발급 (https://www.koreaexim.go.kr)
3. 인터넷 연결 확인
```

### 5.4 성능 최적화

#### 느린 쿼리 새로 고침
```
최적화 방법:
1. Power Query에서 결산연월 필터 추가 (과거 데이터 제외)
2. 백그라운드 새로 고침 활성화
3. 쿼리 연결만 유지 (데이터 로드 안 함)
```

#### VBA 실행 속도 개선
```
1. SpeedUp/SpeedDown 사용 확인
2. 배열 기반 처리로 변경 (Range 직접 접근 최소화)
3. Dictionary 객체 활용 (중첩 루프 제거)
```

---

## 6. 추가 리소스

### 6.1 참고 문서
- BEP VBA_Export 프로젝트 (참고용)
- Power Query M 함수 레퍼런스: https://docs.microsoft.com/power-query/
- Office Ribbon XML 스키마: https://docs.microsoft.com/office/dev/add-ins/

### 6.2 지원 연락처
- 개발팀: [이메일 주소]
- 버그 제보: Ribbon > 정보 > 버그 제보

### 6.3 버전 히스토리
- v2.0 (2026-01-21): 환율 조회 기능 추가, Power Query 기반 아키텍처 전환
- v1.98 (2025-09-13): BEP 기준 초기 버전

---

**© 2026 Samil PwC. All rights reserved.**
