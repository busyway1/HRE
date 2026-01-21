# VBA 코드 Import 가이드 (Windows)

## ⚠️ 중요: 파일 유형별 Import 방법

VBA 모듈은 **유형에 따라 import 방법이 다릅니다**.

---

## 1️⃣ 일반 모듈 (.bas) - 파일 가져오기 가능 ✅

### 대상 파일 (22개)
```
mod_01_FilterSearch.bas
mod_02_FilterSearch_Master.bas
mod_03_PTB_CoA_Input.bas
mod_04_IntializeProgress.bas
mod_05_PTB_Highlight.bas
mod_06_VerifySum.bas
mod_09_CheckMaster.bas
mod_10_Public.bas
mod_11_Sync.bas
mod_16_Export.bas
mod_17_ExchangeRate.bas
mod_Log.bas
mod_MouseWheel.bas
mod_OpenPage.bas
mod_QueryProtection.bas
mod_Refresh.bas
mod_Ribbon.bas
mod_z_Module_GetCursor.bas
Module1.bas
Setup_CoAMaster.bas

+ US-ASCII 파일들 (13개)
ADBS_code.bas
AddCoA_ADBS_code.bas
AddCoA_code.bas
BSPL_code.bas
Check_code.bas
CorpCoA_code.bas
CorpMaster_code.bas
DirectoryURL_code.bas
Guide_code.bas
HideSheet_code.bas
Memo_code.bas
Verify_code.bas
```

### Import 방법
1. VBA Editor (Alt+F11)
2. **파일 → 파일 가져오기** (File → Import File)
3. 위 파일들 전체 선택 (Ctrl+A 또는 Ctrl+클릭)
4. **열기**
5. 같은 이름 모듈 있으면 **"예"** (교체)

---

## 2️⃣ 워크북/시트 객체 코드 - 복사 붙여넣기만 가능 ❌

### 대상 파일 (2개)

#### ⚠️ 현재_통합_문서_code.bas (ThisWorkbook)
- **파일 가져오기 불가**
- **복사-붙여넣기만 가능**

#### ⚠️ CoAMaster_code.bas (CoAMaster Sheet)
- **파일 가져오기 불가**
- **복사-붙여넣기만 가능**

---

## 복사-붙여넣기 상세 방법

### 1단계: 소스 파일 열기

#### 방법 A: VS Code에서 열기 (추천)
```powershell
cd C:\Users\[사용자명]\Desktop\Project\HRE\작업\VBA_Export
code 현재_통합_문서_code.bas
```

#### 방법 B: 메모장에서 열기
- 파일 탐색기에서 우클릭 → **프로그램으로 열기** → **메모장**

---

### 2단계: 코드 복사

1. **VS Code 또는 메모장**에서 파일 열기
2. **전체 선택** (Ctrl+A)
3. **복사** (Ctrl+C)

⚠️ **주의**: 첫 줄부터 마지막 줄까지 **전체 코드** 복사
```vba
VERSION 1.0 CLASS          ← 이 줄부터
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ThisWorkbook"
...
End Sub                    ← 마지막 End Sub까지
```

---

### 3단계: VBA Editor에서 붙여넣기

#### ThisWorkbook (현재_통합_문서_code.bas)

1. **Alt+F11** → VBA Editor
2. 왼쪽 Project Explorer에서 **"ThisWorkbook"** 더블클릭
3. 오른쪽 코드 창에서 **기존 코드 전체 선택** (Ctrl+A)
4. **삭제** (Delete 또는 Backspace)
5. **복사한 코드 붙여넣기** (Ctrl+V)
6. **파일 → 저장** (Ctrl+S)

#### CoAMaster Sheet (CoAMaster_code.bas)

1. **Alt+F11** → VBA Editor
2. 왼쪽 Project Explorer에서 **"Microsoft Excel Objects"** 확장
3. **"CoAMaster" 시트** 더블클릭
4. 오른쪽 코드 창에서 **기존 코드 전체 선택** (Ctrl+A)
5. **삭제**
6. **복사한 코드 붙여넣기** (Ctrl+V)
7. **파일 → 저장** (Ctrl+S)

---

## ⚠️ 붙여넣기 시 주의사항

### 1. VERSION/Attribute 줄 처리

붙여넣을 때 다음 오류 발생 가능:
```
컴파일 오류: Attribute 문이 잘못되었습니다.
```

**해결방법**: VERSION/BEGIN/Attribute 줄 제외하고 붙여넣기

#### 올바른 복사 범위 (Option Explicit부터)
```vba
Option Explicit              ← 여기부터 복사
' ============================================================================
' Module: ThisWorkbook (현재_통합_문서)
...
End Sub                      ← 여기까지
```

#### 제외할 줄 (맨 위 10줄)
```vba
VERSION 1.0 CLASS           ← 제외
BEGIN                       ← 제외
  MultiUse = -1             ← 제외
END                         ← 제외
Attribute VB_Name = "ThisWorkbook"        ← 제외
Attribute VB_GlobalNameSpace = False      ← 제외
Attribute VB_Creatable = False            ← 제외
Attribute VB_PredeclaredId = True         ← 제외
Attribute VB_Exposed = True               ← 제외
```

---

### 2. 한글 깨짐 확인

붙여넣기 후 한글 확인:
```vba
' 정상
LogData_Access ThisWorkbook.Name, "종료"

' 깨짐 (다시 붙여넣기 필요)
LogData_Access ThisWorkbook.Name, "������"
```

---

## 전체 작업 순서 요약

### 1단계: 일반 모듈 Import (5분)
```
파일 → 파일 가져오기 → mod_*.bas 전체 선택 → 열기
```

### 2단계: ThisWorkbook 복사-붙여넣기 (2분)
```
VS Code 열기 → 현재_통합_문서_code.bas 전체 복사
→ VBA Editor → ThisWorkbook → 기존 코드 삭제 → 붙여넣기
```

### 3단계: CoAMaster Sheet 복사-붙여넣기 (2분)
```
VS Code 열기 → CoAMaster_code.bas 전체 복사
→ VBA Editor → CoAMaster 시트 → 기존 코드 삭제 → 붙여넣기
```

### 4단계: 검증 (2분)
1. **디버그 → VBAProject 컴파일**
2. 오류 없으면 성공 ✅
3. 한글 정상 표시 확인

---

## 문제 해결

### Q1. "파일을 찾을 수 없습니다"
**A**: `VBA_Export` 폴더에서 실행했는지 확인
```powershell
cd C:\Users\[사용자명]\Desktop\Project\HRE\작업\VBA_Export
```

### Q2. "Attribute 문이 잘못되었습니다"
**A**: VERSION/Attribute 줄 제외하고 `Option Explicit`부터 복사

### Q3. 한글이 "������" 처럼 깨짐
**A1**: CP949 변환이 제대로 안됨 → `python convert_to_cp949_windows.py` 재실행
**A2**: 파일을 UTF-8로 열었을 가능성 → VS Code 오른쪽 하단 인코딩 확인 (CP949로 변경)

### Q4. 기존 모듈을 삭제할까요?
**A**: **"예"** 선택 (기존 코드 교체)

### Q5. 모듈 이름이 중복됩니다
**A**: 기존 모듈 수동 삭제 후 다시 import
```
VBA Editor → 모듈 우클릭 → [모듈명] 제거 → 예
```

---

## 체크리스트

### Import 완료 확인
- [ ] mod_10_Public 열어서 17번 줄 `"연결마스터"` 한글 정상
- [ ] mod_06_VerifySum 열어서 `"선행 단계"` 한글 정상
- [ ] ThisWorkbook 열어서 `"종료"`, `"실행"` 한글 정상
- [ ] CoAMaster 시트 코드 열어서 한글 정상
- [ ] **디버그 → VBAProject 컴파일** 오류 없음 ✅

### 변환 전 파일 백업 (권장)
```powershell
# 원본 UTF-8 파일 백업
xcopy VBA_Export VBA_Export_UTF8_Backup /E /I

# 변환 실행
python convert_to_cp949_windows.py

# 문제 발생 시 복원
rmdir /S VBA_Export
xcopy VBA_Export_UTF8_Backup VBA_Export /E /I
```

---

**작성일**: 2026-01-21
**버전**: 1.1
**작성자**: Claude Code Assistant
