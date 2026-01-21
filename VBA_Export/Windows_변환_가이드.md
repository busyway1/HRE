# VBA UTF-8 → CP949 변환 가이드 (Windows)

## 배경

macOS에서 VBA 파일을 내보낼 때 UTF-8 인코딩으로 저장되지만, Windows VBA Editor는 CP949(Code Page 949) 인코딩만 지원합니다. 따라서 Windows에서 변환 작업이 필요합니다.

---

## 준비사항

1. **Python 3.x 설치** (이미 설치되어 있으면 생략)
   - https://www.python.org/downloads/
   - 설치 시 "Add Python to PATH" 체크 필수

2. **Git 설치** (이미 설치되어 있으면 생략)
   - https://git-scm.com/download/win
   - Git Bash 또는 PowerShell에서 사용 가능

---

## 실행 방법

### 1단계: Git Repository 업데이트

```powershell
# PowerShell 또는 Git Bash에서 실행
cd C:\Users\[사용자명]\Desktop\Project\HRE\작업\VBA_Export

# 최신 변경사항 가져오기
git pull origin main
```

---

### 2단계: Python 스크립트 실행

```powershell
# 현재 디렉토리에서 변환 스크립트 실행
python convert_to_cp949_windows.py
```

**실행 결과 예시:**
```
============================================================
VBA UTF-8 → CP949 변환 시작
============================================================

[완료] 현재_통합_문서_code.bas
[완료] CoAMaster_code.bas
[완료] mod_01_FilterSearch.bas
[완료] mod_02_FilterSearch_Master.bas
...
[완료] Setup_CoAMaster.bas

============================================================
변환 완료: 21개 성공, 0개 실패
============================================================

다음 단계:
1. Excel 파일 열기
2. Alt+F11 → VBA Editor
3. 파일 → 파일 가져오기
4. 변환된 .bas 파일들 선택
5. 기존 모듈 교체 확인
```

---

### 3단계: VBA Editor에서 코드 Import

⚠️ **중요**: 파일 유형에 따라 import 방법이 다릅니다!

#### 3-1. Excel 파일 열기
- `연결마스터_HRE_v1.00.xlsm` 파일 열기
- 비밀번호: `PwCDA7529`

#### 3-2. VBA Editor 열기
- `Alt + F11` 키 또는
- 개발 도구 탭 → Visual Basic

#### 3-3A. 일반 모듈 - 파일 가져오기 (22개)
1. **파일 → 파일 가져오기** (File → Import File)
2. `VBA_Export` 폴더로 이동
3. 다음 파일들 선택:
   - `mod_*.bas` (19개)
   - US-ASCII 파일들 (13개)
4. **열기** 클릭
5. 기존 모듈 교체 **"예"** 선택

#### 3-3B. ThisWorkbook - 복사 붙여넣기 ⚠️
**`현재_통합_문서_code.bas`는 파일 가져오기 불가!**

1. VS Code 또는 메모장에서 `현재_통합_문서_code.bas` 열기
2. **`Option Explicit`부터 마지막까지** 전체 복사 (Ctrl+A, Ctrl+C)
   - ⚠️ 주의: `VERSION`, `Attribute` 줄은 제외
3. VBA Editor → **ThisWorkbook** 더블클릭
4. 기존 코드 전체 삭제 (Ctrl+A, Delete)
5. 붙여넣기 (Ctrl+V)
6. 저장 (Ctrl+S)

#### 3-3C. CoAMaster Sheet - 복사 붙여넣기 ⚠️
**`CoAMaster_code.bas`는 파일 가져오기 불가!**

1. VS Code 또는 메모장에서 `CoAMaster_code.bas` 열기
2. **`Option Explicit`부터 마지막까지** 전체 복사
3. VBA Editor → **Microsoft Excel Objects** → **CoAMaster** 더블클릭
4. 기존 코드 전체 삭제
5. 붙여넣기
6. 저장

#### 상세 가이드
복사-붙여넣기 방법은 `VBA_Import_가이드.md` 참조

---

### 4단계: 한글 텍스트 검증

#### 4-1. 주요 모듈 확인
1. **mod_10_Public** 모듈 열기
   - 17번 줄: `Public Const AppType = "연결마스터"`
   - 한글이 정상적으로 보이면 성공

2. **mod_06_VerifySum** 모듈 열기
   - 10번 줄: `GoEnd "선행 단계를 완료하세요!"`
   - 한글이 정상적으로 보이면 성공

3. **mod_17_ExchangeRate** 모듈 열기
   - MsgBox 문자열에 "환율" 한글 확인

#### 4-2. 컴파일 테스트
1. **디버그 → VBAProject 컴파일** (Debug → Compile VBAProject)
2. 컴파일 오류 없으면 성공

---

## 문제 해결

### 오류 1: Python을 찾을 수 없습니다
```
'python'은(는) 내부 또는 외부 명령, 실행할 수 있는 프로그램...
```

**해결방법:**
```powershell
# Python 3 명령어 시도
python3 convert_to_cp949_windows.py

# 또는 PATH 확인
where python
```

---

### 오류 2: UnicodeEncodeError
```
[오류] mod_XX.bas: 'cp949' codec can't encode character...
```

**해결방법:**
- 해당 파일에 CP949에서 지원하지 않는 유니코드 문자 존재
- 수동으로 해당 문자를 ASCII로 교체 필요
- 예: ① → (1), ★ → *, ─ → -

---

### 오류 3: 한글이 깨져서 보임
```
Public Const AppType = "������"
```

**해결방법:**
1. VBA Editor 닫기
2. 파일이 실제로 CP949로 변환되었는지 확인:
   ```powershell
   # PowerShell에서 파일 인코딩 확인
   [System.IO.File]::ReadAllText("mod_10_Public.bas", [System.Text.Encoding]::GetEncoding(949))
   ```
3. 한글이 정상이면 → VBA Editor 재시작
4. 한글이 여전히 깨지면 → Git에서 원본 복원 후 재변환

---

## 변환 대상 파일 목록 (21개)

### Worksheet 코드 모듈 (2개)
- 현재_통합_문서_code.bas
- CoAMaster_code.bas

### 워크플로 모듈 (11개)
- mod_01_FilterSearch.bas
- mod_02_FilterSearch_Master.bas
- mod_03_PTB_CoA_Input.bas
- mod_04_IntializeProgress.bas
- mod_05_PTB_Highlight.bas
- mod_06_VerifySum.bas
- mod_09_CheckMaster.bas
- mod_10_Public.bas
- mod_11_Sync.bas
- mod_16_Export.bas
- mod_17_ExchangeRate.bas

### 유틸리티 모듈 (7개)
- mod_Log.bas
- mod_OpenPage.bas
- mod_QueryProtection.bas
- mod_Refresh.bas
- mod_Ribbon.bas
- mod_z_Module_GetCursor.bas
- Module1.bas

### 설정 모듈 (1개)
- Setup_CoAMaster.bas

---

## 변환하지 않는 파일 (13개)

다음 파일들은 한글이 없어서 변환 불필요:
- ADBS_code.bas
- AddCoA_ADBS_code.bas
- AddCoA_code.bas
- BSPL_code.bas
- Check_code.bas
- CorpCoA_code.bas
- CorpMaster_code.bas
- DirectoryURL_code.bas
- Guide_code.bas
- HideSheet_code.bas
- Memo_code.bas
- mod_MouseWheel.bas
- Verify_code.bas

---

## 참고사항

### CP949 인코딩이란?
- **Code Page 949**: Windows 한글 표준 인코딩
- EUC-KR의 확장 버전
- Windows VBA Editor 기본 인코딩
- 한글 2,350자 + 한자 4,888자 지원

### UTF-8 vs CP949
| 항목 | UTF-8 | CP949 |
|------|-------|-------|
| 문자 범위 | 전 세계 모든 문자 | 한글 + 일부 한자 |
| 바이트 크기 | 한글 3바이트 | 한글 2바이트 |
| VBA Editor | ❌ 지원 안함 | ✅ 지원 |
| macOS Excel | ✅ 기본값 | ❌ 지원 안함 |

### 왜 macOS에서 변환 안되나요?
- macOS Python의 `cp949` 코덱이 불완전
- `iconv` 명령어도 한글 누락
- Windows에서만 정확한 변환 가능

---

## 작업 완료 후

### Git 커밋 (선택사항)
```powershell
# 변환된 파일 확인
git status

# 변환된 파일 스테이징
git add *.bas

# 커밋
git commit -m "VBA 파일 UTF-8 → CP949 인코딩 변환 (Windows)"

# 푸시 (필요시)
git push origin main
```

---

## 문의사항

변환 중 문제 발생 시:
1. 오류 메시지 전체 복사
2. 문제가 발생한 파일명 확인
3. 관련 정보와 함께 문의

---

**작성일**: 2026-01-21
**버전**: 1.0
**작성자**: Claude Code Assistant
