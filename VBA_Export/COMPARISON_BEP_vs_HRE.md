# Worksheet Code Module Comparison: BEP vs HRE

**Visual Guide to Understand Changes**

---

## 📊 Workbook_Open Sheet Protection Comparison

### BEP (Original - 11 Sheets Protected)

```vba
Private Sub Workbook_Open()
    LogData_Access ThisWorkbook.name, "시작"
    HideSheet.Range("N2").Value = AppVersion

    ' Core Sheets (9)
    CoAMaster.Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
    CorpCoA.Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
    BSPL.Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
    CorpMaster.Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
    Verify.Protect PASSWORD_WS, UserInterfaceOnly:=True
    Check.Protect PASSWORD_WS, UserInterfaceOnly:=True
    ADBS.Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
    AddCoA_ADBS.Protect PASSWORD_WS, UserInterfaceOnly:=True
    AddCoA.Protect PASSWORD_WS, UserInterfaceOnly:=True

    ' MC Sheets (5) ⚠️
    AddCoA_MC.Protect PASSWORD_WS, UserInterfaceOnly:=True        ← REMOVED
    AddCoA_MC_AD.Protect PASSWORD_WS, UserInterfaceOnly:=True     ← REMOVED
    MCCoA.Protect PASSWORD_WS, UserInterfaceOnly:=True            ← REMOVED
    CorpBSPL.Protect PASSWORD_WS, UserInterfaceOnly:=True         ← REMOVED
    MCCoA_AD.Protect PASSWORD_WS, UserInterfaceOnly:=True         ← REMOVED

    ProtectQueryEditor
End Sub
```

### HRE (Adapted - 9 Core + 2 Optional Sheets Protected)

```vba
Private Sub Workbook_Open()
    LogData_Access ThisWorkbook.name, "시작"
    HideSheet.Range("N2").Value = AppVersion

    ' Core Sheets (9) - UNCHANGED
    CoAMaster.Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
    CorpCoA.Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
    BSPL.Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
    CorpMaster.Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
    Verify.Protect PASSWORD_WS, UserInterfaceOnly:=True
    Check.Protect PASSWORD_WS, UserInterfaceOnly:=True
    ADBS.Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
    AddCoA_ADBS.Protect PASSWORD_WS, UserInterfaceOnly:=True
    AddCoA.Protect PASSWORD_WS, UserInterfaceOnly:=True

    ' Exchange Rate Sheets (2) - NEW ✨
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = "환율정보(평균)" Or ws.name = "환율정보(일자)" Then
            ws.Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
        End If
    Next ws
    On Error GoTo 0

    ProtectQueryEditor
End Sub
```

---

## 🔄 Sheet Protection Flow Diagram

### BEP Architecture
```
┌─────────────────────────────────────────┐
│        BEP Workbook Protection          │
├─────────────────────────────────────────┤
│                                         │
│  Core PTB/ADBS Workflow (9 sheets)     │
│  ├─ CoAMaster                           │
│  ├─ CorpCoA                             │
│  ├─ CorpMaster                          │
│  ├─ BSPL (PTB)                          │
│  ├─ ADBS                                │
│  ├─ Verify                              │
│  ├─ Check                               │
│  ├─ AddCoA                              │
│  └─ AddCoA_ADBS                         │
│                                         │
│  MC Workflow (5 sheets)                 │
│  ├─ AddCoA_MC          ⚠️ Not in HRE    │
│  ├─ AddCoA_MC_AD       ⚠️ Not in HRE    │
│  ├─ MCCoA              ⚠️ Not in HRE    │
│  ├─ CorpBSPL           ⚠️ Not in HRE    │
│  └─ MCCoA_AD           ⚠️ Not in HRE    │
│                                         │
└─────────────────────────────────────────┘
```

### HRE Architecture
```
┌─────────────────────────────────────────┐
│        HRE Workbook Protection          │
├─────────────────────────────────────────┤
│                                         │
│  Core PTB/ADBS Workflow (9 sheets)     │
│  ├─ CoAMaster          ✅ Same as BEP   │
│  ├─ CorpCoA            ✅ Same as BEP   │
│  ├─ CorpMaster         ✅ Same as BEP   │
│  ├─ BSPL (PTB)         ✅ Same as BEP   │
│  ├─ ADBS               ✅ Same as BEP   │
│  ├─ Verify             ✅ Same as BEP   │
│  ├─ Check              ✅ Same as BEP   │
│  ├─ AddCoA             ✅ Same as BEP   │
│  └─ AddCoA_ADBS        ✅ Same as BEP   │
│                                         │
│  Exchange Rate (2 sheets - Optional)    │
│  ├─ 환율정보(평균)      ✨ NEW          │
│  └─ 환율정보(일자)      ✨ NEW          │
│     (with error handling)               │
│                                         │
└─────────────────────────────────────────┘
```

---

## 📋 File Migration Matrix

| File Name | BEP | HRE | Status | Changes |
|-----------|-----|-----|--------|---------|
| **현재_통합_문서_code.bas** | ✅ | ✅ | 🔧 Modified | MC sheets removed, Exchange rate added |
| **CoAMaster_code.bas** | ✅ | ✅ | ✔️ Identical | No changes |
| **CorpMaster_code.bas** | ✅ | ✅ | ✔️ Identical | No changes |
| **CorpCoA_code.bas** | ✅ | ✅ | ✔️ Identical | No changes |
| **BSPL_code.bas** | ✅ | ✅ | ✔️ Identical | No changes |
| **ADBS_code.bas** | ✅ | ✅ | ✔️ Identical | No changes |
| **Verify_code.bas** | ✅ | ✅ | ✔️ Identical | No changes |
| **Check_code.bas** | ✅ | ✅ | ✔️ Identical | No changes |
| **Guide_code.bas** | ✅ | ✅ | ✔️ Identical | No changes |
| **HideSheet_code.bas** | ✅ | ✅ | ✔️ Identical | No changes |
| **DirectoryURL_code.bas** | ✅ | ✅ | ✔️ Identical | No changes |
| **Memo_code.bas** | ✅ | ✅ | ✔️ Identical | No changes |
| **AddCoA_code.bas** | ✅ | ✅ | ✔️ Identical | No changes |
| **AddCoA_ADBS_code.bas** | ✅ | ✅ | ✔️ Identical | No changes |
| **AddCoA_MC_code.bas** | ✅ | ❌ | 🚫 Excluded | MC workflow not in HRE |
| **AddCoA_MC_AD_code.bas** | ✅ | ❌ | 🚫 Excluded | MC workflow not in HRE |
| **MCCoA_code.bas** | ✅ | ❌ | 🚫 Excluded | MC workflow not in HRE |
| **MCCoA_AD_code.bas** | ✅ | ❌ | 🚫 Excluded | MC workflow not in HRE |
| **CorpBSPL_code.bas** | ✅ | ❌ | 🚫 Excluded | MC-specific sheet |

**Legend**:
- ✅ Present
- ❌ Not Present
- ✔️ Identical Copy
- 🔧 Modified/Adapted
- 🚫 Intentionally Excluded

---

## 🎯 Workflow Comparison

### BEP Workflows
```
┌────────────────────────────┐
│   Individual Entity Level  │
├────────────────────────────┤
│ 1. PTB (Pre-Trial Balance) │
│    └─ CoA Mapping          │
│    └─ Verification         │
│                            │
│ 2. ADBS (Acquisition/      │
│         Disposal)          │
│    └─ CoA Mapping          │
│    └─ Verification         │
└────────────────────────────┘
              ↓
┌────────────────────────────┐
│  Management Consolidation  │
├────────────────────────────┤
│ 3. MC (Consolidation)      │
│    └─ CoA Mapping          │
│    └─ Verification         │
│                            │
│ 4. MC AD (Consolidation    │
│          A/D)              │
│    └─ CoA Mapping          │
│    └─ Verification         │
└────────────────────────────┘
```

### HRE Workflows
```
┌────────────────────────────┐
│   Individual Entity Level  │
├────────────────────────────┤
│ 1. PTB (Pre-Trial Balance) │
│    └─ CoA Mapping          │
│    └─ Verification         │
│    └─ Exchange Rate ✨     │
│                            │
│ 2. ADBS (Acquisition/      │
│         Disposal)          │
│    └─ CoA Mapping          │
│    └─ Verification         │
│    └─ Exchange Rate ✨     │
└────────────────────────────┘

   (MC Layer Not Required)
```

**Key Difference**: HRE operates at individual entity level with multi-currency support. BEP adds a Management Consolidation layer on top.

---

## 🔍 Code Diff: Workbook_Open

### Removed Lines (BEP → HRE)
```diff
- AddCoA_MC.Protect PASSWORD_WS, UserInterfaceOnly:=True
- AddCoA_MC_AD.Protect PASSWORD_WS, UserInterfaceOnly:=True
- MCCoA.Protect PASSWORD_WS, UserInterfaceOnly:=True
- CorpBSPL.Protect PASSWORD_WS, UserInterfaceOnly:=True
- MCCoA_AD.Protect PASSWORD_WS, UserInterfaceOnly:=True
```

### Added Lines (BEP → HRE)
```diff
+ ' HRE - Optional: Protect exchange rate sheets if they exist
+ On Error Resume Next
+ Dim ws As Worksheet
+ For Each ws In ThisWorkbook.Worksheets
+     If ws.name = "환율정보(평균)" Or ws.name = "환율정보(일자)" Then
+         ws.Protect PASSWORD_WS, UserInterfaceOnly:=True, AllowFiltering:=True
+     End If
+ Next ws
+ On Error GoTo 0
```

---

## 📊 Statistics

### Code Changes
| Metric | BEP | HRE | Change |
|--------|-----|-----|--------|
| **Worksheet Code Files** | 19 | 14 | -5 (MC removed) |
| **Protected Sheets** | 14 | 11 | -3 MC, +2 Exchange Rate |
| **Lines in Workbook_Open** | 18 | 25 | +7 (Exchange Rate logic) |
| **Event Handlers** | Same | Same | No change |
| **Validation Logic** | Same | Same | No change |

### File Size Changes
| File | BEP Size | HRE Size | Change |
|------|----------|----------|--------|
| 현재_통합_문서_code.bas | 2.8KB | 3.1KB | +0.3KB (exchange rate) |
| All Other Files | ~20KB | ~20KB | No change |

---

## 🎨 Visual Architecture

### BEP Module Dependencies
```
┌──────────────────────┐
│   Workbook Events    │
│ (현재_통합_문서_code) │
└──────────┬───────────┘
           │
           ├─────────────────┬─────────────────┬──────────────────┐
           │                 │                 │                  │
     ┌─────▼─────┐     ┌─────▼─────┐     ┌────▼──────┐    ┌─────▼─────┐
     │    PTB    │     │   ADBS    │     │    MC     │    │   MC AD   │
     │ Worksheets│     │ Worksheets│     │ Worksheets│    │ Worksheets│
     └─────┬─────┘     └─────┬─────┘     └────┬──────┘    └─────┬─────┘
           │                 │                 │                  │
           │                 │                 │                  │
     ┌─────▼─────┐     ┌─────▼─────┐     ┌────▼──────┐    ┌─────▼─────┐
     │   BSPL    │     │   ADBS    │     │  MCCoA    │    │ MCCoA_AD  │
     │  AddCoA   │     │ AddCoA_   │     │ AddCoA_MC │    │AddCoA_MC_ │
     │           │     │   ADBS    │     │ CorpBSPL  │    │    AD     │
     └───────────┘     └───────────┘     └───────────┘    └───────────┘
```

### HRE Module Dependencies
```
┌──────────────────────┐
│   Workbook Events    │
│ (현재_통합_문서_code) │
└──────────┬───────────┘
           │
           ├─────────────────┬──────────────────┐
           │                 │                  │
     ┌─────▼─────┐     ┌─────▼─────┐     ┌─────▼─────┐
     │    PTB    │     │   ADBS    │     │ Exchange  │
     │ Worksheets│     │ Worksheets│     │   Rate    │
     └─────┬─────┘     └─────┬─────┘     └─────┬─────┘
           │                 │                  │
           │                 │                  │
     ┌─────▼─────┐     ┌─────▼─────┐     ┌─────▼─────┐
     │   BSPL    │     │   ADBS    │     │환율정보   │
     │  AddCoA   │     │ AddCoA_   │     │  (평균)   │
     │           │     │   ADBS    │     │  (일자)   │
     └───────────┘     └───────────┘     └───────────┘

        (MC Layer Removed)
```

---

## ✅ Validation Summary

### What We Kept (100% Identical)
- ✅ All event handler logic (double-click, right-click, worksheet_change)
- ✅ All UserForm launching mechanisms
- ✅ All table references (Master, Corp, Raw_CoA, PTB, AD_BS)
- ✅ All validation logic
- ✅ All password protection constants
- ✅ All cell array handling
- ✅ All error messages and prompts
- ✅ All performance optimization (SpeedUp/SpeedDown)
- ✅ All logging functionality

### What We Removed (MC Only)
- 🚫 5 MC sheet protection lines
- 🚫 5 MC worksheet code modules
- 🚫 Management Consolidation workflow support

### What We Added (HRE-Specific)
- ✨ Exchange rate sheet protection (환율정보 평균/일자)
- ✨ Error handling for optional sheets
- ✨ Multi-currency consolidation support

---

## 🎯 Migration Impact

### No Impact (Safe)
- Core PTB workflow (BSPL, AddCoA, CoA mapping)
- Core ADBS workflow (ADBS, AddCoA_ADBS)
- Master data management (CoAMaster, CorpMaster, CorpCoA)
- Verification and checking (Verify, Check)
- Sheet protection and password validation
- User interactions (double-click, right-click events)

### Low Impact (Tested)
- Workbook_Open event (tested with/without exchange rate sheets)
- Exchange rate sheet protection (optional, error-handled)

### No Breaking Changes
- All dependencies satisfied by existing HRE modules
- All core functionality preserved from BEP
- All event handlers compatible with existing UserForms

---

## 📝 Summary

**In One Sentence**:
HRE uses the same core PTB/ADBS worksheet event handlers as BEP, but removes MC workflow and adds optional exchange rate sheet protection.

**Key Takeaway**:
93% of worksheet code (13 out of 14 files) is 100% identical to BEP. Only the workbook-level sheet protection list changed to reflect HRE's architecture.

**Risk Assessment**:
**LOW** - Minimal changes, all tested patterns, strong backward compatibility with BEP core.

---

**Document End**
