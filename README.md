# HRE 연결마스터 프로젝트

**버전**: 1.00
**작성일**: 2026-01-21
**용도**: HRE 그룹 연결재무제표 작성 자동화 시스템

---

## 📂 프로젝트 구조

```
HRE/작업/
├── 완벽한_구현_가이드.md          ⭐ 메인 가이드 (VBA 초보자용 전체 매뉴얼)
├── README.md                      현재 파일 (프로젝트 개요)
│
├── VBA_Export/                    VBA 소스 코드 (34개 모듈)
│   ├── mod_01_FilterSearch.bas
│   ├── mod_10_Public.bas         (HRE 상수 업데이트)
│   ├── mod_17_ExchangeRate.bas   ⭐ 신규: 환율 조회
│   ├── mod_03_PTB_CoA_Input.bas  ⭐ 변종 인식 강화
│   ├── 현재_통합_문서_code.bas      (한글 인코딩 수정 완료)
│   └── ...
│
├── UserForms/                     UserForm 파일 (.frm, .frx)
│   ├── frmCalendar.frm
│   ├── frmCorp_Append.frm        ⭐ 지분율/기능통화 필드 추가 필요
│   └── ...
│
├── Documentation/                 기술 문서 (참고용)
│   ├── CLAUDE.md                 AI 개발자용 문서
│   ├── README.md                 영문 사용자 가이드
│   ├── IMPLEMENTATION_CHECKLIST.md
│   ├── INSTALLATION_GUIDE.md
│   └── DOCUMENTATION_SUMMARY.md
│
└── PowerQuery/                    Power Query M 코드
    ├── PowerQuery_PTB.m          SharePoint PTB 연결
    ├── PowerQuery_RawCoA.m       CoA 매핑 이력
    └── PowerQuery_PTB_LocalFile.m (로컬 파일 대안)
```

---

## 🚀 빠른 시작

### 1단계: 가이드 읽기
**`완벽한_구현_가이드.md`** 파일을 열어 전체 구현 절차를 확인하세요.

### 2단계: VBA 모듈 임포트
`VBA_Export/` 폴더의 모든 .bas 파일을 Excel VBA 편집기로 가져오세요.

### 3단계: 리본 메뉴 설치
Custom UI Editor를 사용하여 리본 XML을 추가하세요.

### 4단계: SharePoint 연결
Power Query를 사용하여 SharePoint 데이터를 연결하세요.

---

## ⚡ 핵심 기능

| 기능 | 설명 |
|------|------|
| **5-Digit CoA 매핑** | 계정코드 첫 5자리 기반 자동 매핑 |
| **변종 인식** | `_내부거래`, `_IC` 접미사 자동 감지 |
| **환율 조회** | KEB 하나은행 실시간 환율 API 연동 |
| **다중 통화** | KRW/USD/EUR/JPY/CNY 지원 |
| **자동 검증** | 차변=대변 공식 자동 확인 |

---

## 📋 주요 변경사항 (vs BEP)

| 항목 | BEP v1.98 | HRE v1.00 |
|------|-----------|-----------|
| CoA 매핑 | 전체 일치 | 첫 5자리 + 변종 |
| 환율 | 수동 입력 | API 자동 조회 |
| MC 연결 | 지원 | 제거 (불필요) |
| 내부거래 | 수동 표시 | 자동 플래그 |

---

## 🛠️ 기술 스택

- **플랫폼**: Microsoft Excel 2016+ (Microsoft 365 권장)
- **언어**: VBA (Visual Basic for Applications)
- **데이터 연결**: Power Query M Language
- **저장소**: SharePoint Online
- **API**: KEB 하나은행 환율 조회 (MSXML2.ServerXMLHTTP)

---

## 📞 문의

- **GitHub**: https://github.com/busyway1/HRE.git
- **SharePoint**: https://pwckor.sharepoint.com/sites/KR-ASR-HRE_Consolidation
- **이메일**: (담당자 이메일)

---

## 📄 라이선스

© 2026 Samil PwC. All rights reserved.

이 소프트웨어는 PwC의 내부 사용 목적으로만 제공됩니다.
무단 복제, 배포, 수정을 금지합니다.
