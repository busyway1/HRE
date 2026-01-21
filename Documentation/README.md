# HRE 연결마스터 (Consolidation Master) - User Guide

**Version**: 1.00
**Developer**: Samil PwC
**Last Updated**: 2026-01-21

---

## Table of Contents

1. [Getting Started](#getting-started)
2. [12-Step Workflow](#12-step-workflow)
3. [Key Features](#key-features)
4. [Exchange Rate Integration](#exchange-rate-integration)
5. [Troubleshooting](#troubleshooting)
6. [Support](#support)

---

## Getting Started

### System Requirements

- **Excel Version**: Excel 2016 or later (Microsoft 365 recommended)
- **Operating System**: Windows 10 or later
- **VBA Enabled**: Macros must be enabled
- **Network Access**: Required for SharePoint connection and exchange rate updates

### Opening the File

1. **Download** the HRE Consolidation Master file from your SharePoint site
2. **Enable Macros** when prompted:
   - Click "Enable Content" in the yellow security warning bar
3. **Wait for Initialization**:
   - The system will automatically protect worksheets
   - Custom ribbon tab "HRE 연결마스터" will appear

### Password Information

- **Worksheet Password**: `BEP1234` (for advanced users only)
- **Workbook Password**: `PwCDA7529` (for structure changes only)

> ⚠️ **Warning**: Do not modify passwords unless instructed by PwC support team.

### Custom Ribbon Tab

The **HRE 연결마스터** ribbon tab provides quick access to all workflow functions:

- **Data Management**: Query refresh, CoA sync, filters
- **Verification**: PTB highlight, ADBS highlight, sum checks
- **Exchange Rates**: 평균환율 조회, 기말환율 조회
- **Export**: Final data export to reporting format

---

## 12-Step Workflow

Follow these steps in order for accurate consolidated financial statements:

### **Step 1: Configure SharePoint Connection** 📡

**Purpose**: Establish connection to corporate CoA data source

1. Click ribbon button: **SPO 연결**
2. Enter SharePoint site URL in the dialog
3. Click **확인** to save connection
4. Verify connection status in `Check` sheet (Row 12)

**Expected Result**: ✅ Check sheet Row 12 shows "Complete"

---

### **Step 2: Refresh Query Data** 🔄

**Purpose**: Pull latest corporate CoA data from SharePoint

1. Click ribbon button: **쿼리 새로 고침**
2. Wait for progress bar to complete (may take 1-2 minutes)
3. Review `CorpCoA` sheet for updated data

**Expected Result**:
- ✅ `Raw_CoA` table populated with latest data
- ✅ Check sheet Row 13 shows "Complete"

---

### **Step 3: Highlight PTB (Pre-Trial Balance)** 🎨

**Purpose**: Identify unmapped accounts in trial balance data

1. Ensure `PTB` table in `BSPL` sheet contains trial balance data
2. Click ribbon button: **PTB CoA 확인**
3. Wait for highlighting to complete

**Color Coding**:
- 🟡 **Yellow**: Account needs CoA mapping (not yet mapped)
- 🟢 **Green**: Account already mapped
- ⚪ **White**: No mapping required

**Expected Result**:
- ✅ Yellow rows indicate accounts needing attention
- ✅ Check sheet Row 14 shows "Complete"

---

### **Step 4: Filter Yellow Rows** 🔍

**Purpose**: Focus on accounts requiring CoA mapping

1. Click ribbon button: **PTB 필터**
2. Table auto-filters to show only yellow-highlighted rows
3. Review account codes and descriptions

**Expected Result**:
- ✅ Only unmapped accounts visible
- ✅ Ready for CoA input

---

### **Step 5: Input CoA Mappings** ✍️

**Purpose**: Map corporate accounts to PwC consolidated CoA

1. Click ribbon button: **PTB CoA 입력**
2. System auto-populates `CoA_Input` table with:
   - **Auto-Detected Mappings**: Based on 5-digit base code matching
   - **Variant Mappings**: `_내부거래` and `_IC` suffixes auto-map to intercompany accounts
   - **Empty Cells**: Require manual review
3. Review suggested mappings in columns:
   - **PwC_CoA**: Auto-suggested consolidated account code
   - **PwC_계정과목명**: Auto-suggested account name
4. **Manual Review**:
   - Verify auto-suggestions are correct
   - Fill in empty cells using dropdown or Master table reference
   - Double-click cells for CoA search dialog

**Variant Handling**:
- `11401_내부거래` → Auto-maps to `112800` (Interco Receivable)
- `25301_내부거래` → Auto-maps to `212800` (Interco Payable)
- `11602_IC` → Auto-maps to `112800` (Interco Receivable)

**Expected Result**:
- ✅ All rows have valid PwC_CoA and PwC_계정과목명
- ✅ No yellow cells in mapping columns

---

### **Step 6: Finalize CoA Mappings** ✅

**Purpose**: Commit CoA mappings to master table

1. Review all mappings one final time
2. Click ribbon button: **PTB CoA 완료**
3. System validates:
   - No empty mappings
   - All PwC codes exist in Master table
   - No duplicate mappings
4. Confirm dialog: **예**

**Expected Result**:
- ✅ Mappings saved to `Raw_CoA` table
- ✅ `AddCoA` sheet hidden
- ✅ Check sheet Row 15 shows "Complete"
- ✅ PTB rows turn green

---

### **Step 7: Verify Financial Statement Sums** 🧮

**Purpose**: Ensure trial balance agrees with financial statements

1. Click ribbon button: **재무제표 검증**
2. System performs:
   - Balance Sheet balance check (Assets = Liabilities + Equity)
   - P&L sum verification
   - Pivot table refresh and comparison
3. Review `Verify` sheet for discrepancies

**Expected Result**:
- ✅ All verification checks pass
- ✅ Check sheet Row 16 shows "Complete"
- ✅ No red cells in Verify sheet

---

### **Step 8: Highlight ADBS (Acquisition/Disposal BS)** 🎨

**Purpose**: Identify unmapped AD transaction accounts

1. Click ribbon button: **ADBS CoA 확인**
2. Wait for highlighting to complete

**Expected Result**:
- ✅ Yellow rows in `ADBS` sheet indicate unmapped accounts
- ✅ Check sheet Row 17 shows "Complete"

---

### **Step 9: Input ADBS CoA Mappings** ✍️

**Purpose**: Map acquisition/disposal accounts

1. Click ribbon button: **ADBS CoA 입력**
2. Follow same process as Step 5 (PTB CoA input)
3. Review auto-suggestions and fill manual entries

**Expected Result**:
- ✅ All ADBS accounts mapped
- ✅ Check sheet Row 18 shows "Complete"

---

### **Step 10: Sync CoA Master** 🔄

**Purpose**: Ensure consistency across all CoA tables

1. Click ribbon button: **CoA 동기화**
2. System synchronizes:
   - `Raw_CoA` table
   - `Master` table
   - All subsidiary tables
3. Review sync log for any conflicts

**Expected Result**:
- ✅ All tables in sync
- ✅ Check sheet Row 19 shows "If Any"

---

### **Step 11: MC (Management Consolidation) Processing** 🏢

**Purpose**: Handle consolidation-specific adjustments

1. Click ribbon button: **MC 하이라이트**
2. Review MC accounts (excluded from auto-mapping)
3. Process MC adjustments manually if required
4. Click ribbon button: **MC CoA 입력** (if needed)

**Expected Result**:
- ✅ MC accounts properly classified
- ✅ Check sheet Rows 20-21 show "Complete" or "If Any"

---

### **Step 12: Update Exchange Rates** 💱

**Purpose**: Fetch latest KEB Hana Bank exchange rates for currency conversion

#### **Option A: Average Exchange Rates (P&L Accounts)**

1. Click ribbon button: **평균환율 조회**
2. Select **Start Date** in calendar dialog
3. Select **End Date** in calendar dialog
4. Wait for data retrieval (5-10 seconds)
5. Review `환율정보(평균)` sheet:
   - Average rates for all currencies (USD, EUR, JPY, CNY, VND, IDR, etc.)
   - Special currencies (JPY, VND, IDR) show 환산=100
   - KRW baseline row at bottom (환산=1, 매매기준율=1)

**Use Cases**:
- Income statement conversions (revenue, expenses over a period)
- Year-to-date P&L consolidation
- Quarterly/monthly P&L reporting

#### **Option B: Spot Exchange Rates (Balance Sheet Accounts)**

1. Click ribbon button: **기말환율 조회**
2. Select **Single Date** in calendar dialog (e.g., period-end date)
3. Wait for data retrieval (5-10 seconds)
4. Review `환율정보(일자)` sheet:
   - Spot rates for all currencies as of selected date
   - Same special currency handling
   - KRW baseline row at bottom

**Use Cases**:
- Balance sheet account conversions (cash, receivables, payables)
- Period-end position consolidation
- Asset/liability revaluation

**Important Notes**:
- ⚠️ Only past dates allowed (cannot select future dates)
- ⚠️ January 1st auto-adjusts to January 2nd (bank holiday)
- ⚠️ Weekend/holiday dates fall back to previous business day automatically
- ✅ Rates are official KEB Hana Bank published rates

**Expected Result**:
- ✅ Exchange rate sheet populated with current rates
- ✅ Check sheet Row 20 shows "Complete"

---

### **Step 13: Export Data** 📤

**Purpose**: Generate final consolidated financial statements

1. Click ribbon button: **데이터 내보내기**
2. Select export format (Excel, CSV, or custom)
3. Choose destination folder
4. Confirm export

**Expected Result**:
- ✅ Consolidated financial statements exported
- ✅ All workflow steps complete
- ✅ Ready for financial reporting

---

## Key Features

### 🎯 Auto CoA Mapping with Variant Detection

**5-Digit Base Code Matching**:
- Unlike exact match systems, HRE uses **first 5 digits** for base matching
- Example: `10300` → Matches all variants of base code `10300`

**Variant Suffix Recognition**:
- **`_내부거래`** (Internal Transaction - Korean): Auto-maps to intercompany accounts
- **`_IC`** (Internal Transaction - International): Auto-maps to intercompany accounts
- **BASE** (No suffix): Standard accounts

**Multi-Tier Lookup Strategy**:
1. **Exact Variant Match**: `11401_내부거래` → Search for INTERCO_KR variant
2. **BASE Fallback**: If no variant match, use BASE variant mapping
3. **Manual Review**: If no match found, leave empty for user input

**Example Mappings**:
```
Account Code         → Variant Type  → PwC CoA  → Description
10300                → BASE          → 111206   → Cash - Operating - CNY
11401_내부거래        → INTERCO_KR    → 112800   → Interco Receivable
11401                → BASE          → 112332   → Other Receivable ST
11602_IC             → INTERCO_IC    → 112800   → Interco Receivable
25301_내부거래        → INTERCO_KR    → 212800   → Interco Payable
```

### 💱 Exchange Rate Integration

**KEB Hana Bank Official API**:
- Direct connection to bank's published rates
- No manual data entry required
- Automatic daily rate updates

**Special Currency Handling**:
- **Standard Currencies** (USD, EUR, CNY): Quoted per 1 unit (환산=1)
- **Special Currencies** (JPY, VND, IDR): Quoted per 100 units (환산=100)
  - Example: JPY 100 = 1,000 KRW (easier to read than JPY 1 = 10 KRW)

**Automatic Adjustments**:
- Holiday handling (January 1st → January 2nd)
- Weekend fallback to previous business day
- KRW baseline always included (1 KRW = 1 KRW)

### 🔍 Advanced Filtering and Search

**Filter by Status**:
- Yellow rows (unmapped accounts)
- Green rows (mapped accounts)
- All rows (complete view)

**Master Table Search**:
- Double-click any CoA cell to open search dialog
- Filter by category, account code, or description
- Quick lookup with keyboard shortcuts

### ✅ Multi-Level Validation

**Pre-Save Validation**:
- Empty mapping detection
- Master table existence check
- Duplicate prevention

**Post-Save Verification**:
- Balance Sheet balance (Assets = Liabilities + Equity)
- P&L sum checks
- Intercompany elimination verification

### 📊 Progress Tracking

**Check Sheet Status**:
- Visual progress indicators (green = complete, yellow = in progress)
- Timestamp and user tracking for each step
- Workflow dependency validation

---

## Exchange Rate Integration

### When to Use Average Rates vs. Spot Rates

| Account Type | Exchange Rate Type | Example Accounts |
|--------------|-------------------|-----------------|
| **Income Statement** | 평균환율 (Average) | Revenue, Expenses, Interest |
| **Balance Sheet** | 기말환율 (Spot) | Cash, Receivables, Payables, Debt |
| **Equity** | Historical Rate | Share Capital, Retained Earnings Opening Balance |

### Step-by-Step: Updating Exchange Rates

#### For P&L Accounts (Average Rates)

1. **Determine Period**:
   - Example: Fiscal year 2024-01-01 to 2024-12-31
   - Or: Q1 2024 → 2024-01-01 to 2024-03-31

2. **Fetch Rates**:
   - Click **평균환율 조회**
   - Select start date: `2024-01-01`
   - Select end date: `2024-12-31`

3. **Apply to Conversions**:
   - Use average rate for P&L accounts (매매기준율 column)
   - Example: USD revenue × Average USD rate

#### For Balance Sheet Accounts (Spot Rates)

1. **Determine Date**:
   - Example: Period-end date 2024-12-31

2. **Fetch Rates**:
   - Click **기말환율 조회**
   - Select date: `2024-12-31`

3. **Apply to Conversions**:
   - Use spot rate for B/S accounts (매매기준율 column)
   - Example: USD cash × Spot USD rate at 2024-12-31

### Currency Conversion Formula

**Standard Currencies (USD, EUR, CNY)**:
```
KRW Amount = Foreign Currency Amount × 매매기준율 × (1 / 환산)
           = Foreign Currency Amount × 매매기준율 × (1 / 1)
           = Foreign Currency Amount × 매매기준율
```

**Special Currencies (JPY, VND, IDR)**:
```
KRW Amount = Foreign Currency Amount × 매매기준율 × (1 / 환산)
           = Foreign Currency Amount × 매매기준율 × (1 / 100)
```

**Example**:
- USD 1,000 × 1,300 KRW/USD = 1,300,000 KRW
- JPY 100,000 × 1,100 KRW/100JPY × (1/100) = 1,100,000 KRW

### Exchange Rate Sheet Structure

**환율정보(평균)** and **환율정보(일자)** sheets contain:

| Column | Header | Description |
|--------|--------|-------------|
| A | 국가명 및 통화 | Country name and currency code (e.g., "미국 USD") |
| B | 통화 | Currency code (e.g., "USD") |
| C | 환산 | Conversion factor (1 for standard, 100 for JPY/VND/IDR) |
| D-M | Rate Columns | Various rate types (매매기준율, 현찰 매입, 현찰 매도, etc.) |

**Key Columns for Consolidation**:
- **매매기준율** (Column K): Base rate for conversions
- **통화** (Column B): Currency code for matching
- **환산** (Column C): Conversion factor for special currencies

### Troubleshooting Exchange Rates

**Issue**: "유효하지 않은 날짜입니다" error

**Solution**:
- Ensure selected date is not in the future
- For average rates, ensure start date < end date

---

**Issue**: Exchange rate sheet is empty or incomplete

**Solution**:
- Check network connection
- Verify KEB Hana Bank website is accessible
- Try a different date (current date may be before bank's daily rate posting)

---

**Issue**: Weekend/holiday rates missing

**Solution**:
- This is normal behavior
- API automatically falls back to previous business day
- Note at top of sheet explains: "※ 조회일이 토/일/공휴일 또는 은행영업일 1회차 고시 전인 경우, 전 영업일자로 조회됩니다."

---

**Issue**: Special currency amounts incorrect

**Solution**:
- Verify you're using correct 환산 factor
- JPY, VND, IDR require division by 100
- Formula: `Amount × Rate × (1 / 환산)`

---

## Troubleshooting

### Common Issues

#### 🔴 "매크로 보안 경고" (Macro Security Warning)

**Symptom**: Yellow bar at top of Excel window

**Solution**:
1. Click **콘텐츠 사용** (Enable Content)
2. If persists, go to File → Options → Trust Center → Trust Center Settings
3. Select "매크로 설정" → "알림과 함께 모든 매크로 제외"
4. Restart Excel and reopen file

---

#### 🔴 "PwC_CoA와 PwC_계정과목명 매칭되지 않은 항목이 있습니다."

**Symptom**: Cannot finalize CoA mappings (Step 6)

**Solution**:
1. Review yellow-highlighted cells in `CoA_Input` sheet
2. Verify account codes exist in `Master` table
3. Use double-click search to find correct mapping
4. Ensure no typos in account codes

---

#### 🔴 CoA Auto-Mapping Not Working

**Symptom**: All suggested mappings are empty in Step 5

**Solution**:
1. Verify `Raw_CoA` table is populated (Step 2 complete)
2. Check `Raw_CoA` table has Corp Code "1000" entries
3. Ensure variant suffixes match exactly (`_내부거래`, `_IC`)
4. Manually map first few accounts, then re-run CoA sync (Step 10)

---

#### 🔴 Variant Accounts Not Recognized

**Symptom**: `_내부거래` accounts not auto-mapping to intercompany codes

**Solution**:
1. Verify suffix spelling: `_내부거래` (not `_내부 거래` with space)
2. Check `coa.md` reference file has variant entries
3. Update `Raw_CoA` table manually if needed:
   - Corp Code: 1000
   - 계정코드: `11401_내부거래`
   - Account: `112800`
   - Description: `Interco Receivable`

---

#### 🔴 Balance Sheet Does Not Balance

**Symptom**: Verification (Step 7) shows discrepancies

**Solution**:
1. Review `Verify` sheet for specific accounts with issues
2. Check for missing CoA mappings (yellow rows in PTB)
3. Verify trial balance data is complete in `PTB` table
4. Ensure all subsidiary data is refreshed from SharePoint

---

#### 🔴 Exchange Rate Retrieval Fails

**Symptom**: "Network error" or empty exchange rate sheet

**Solution**:
1. Check internet connection
2. Verify KEB Hana Bank website is accessible: https://www.kebhana.com
3. Try a different date (within last 30 days)
4. Contact IT support if corporate firewall blocks KEB Hana Bank domain

---

#### 🔴 Ribbon Tab Not Visible

**Symptom**: "HRE 연결마스터" tab missing from ribbon

**Solution**:
1. Close and reopen the file
2. Enable macros when prompted
3. Check if Developer tab shows VBA project is loaded
4. Re-import custom ribbon XML if needed (see IMPLEMENTATION_CHECKLIST.md)

---

#### 🔴 "사용 기간이 만료되었습니다!" Error

**Symptom**: File won't open or shows expiration message

**Solution**:
1. Check current date vs. expiration date (2030-12-31)
2. Contact PwC support for updated version
3. Verify system clock is correct (not set to future date)

---

### Performance Tips

**Slow Query Refresh (Step 2)**:
- SharePoint query can take 2-5 minutes for large datasets
- Do not interrupt the process
- Close other Excel files to free memory

**Slow CoA Input (Step 5)**:
- For 1,000+ unmapped accounts, consider batch processing
- Use filter to process 100 rows at a time
- Dictionary-based auto-mapping is optimized for performance

**Slow Verification (Step 7)**:
- Pivot table refresh can be slow for large datasets
- Ensure Excel calculation is set to Automatic
- Close background applications

---

## Support

### Internal Support (PwC Users)

**Primary Contact**:
- **Email**: pwcda@pwc.com
- **Teams**: PwC Digital Assurance - HRE Support Channel

**Self-Service Resources**:
- **Documentation**: SharePoint site → HRE Consolidation Master → Documentation folder
- **Training Videos**: SharePoint site → HRE Consolidation Master → Training folder
- **FAQ**: SharePoint site → HRE Consolidation Master → FAQ.md

### External Support (HRE Users)

**Primary Contact**:
- **Email**: hre-support@hre.com
- **Phone**: +82-2-xxxx-xxxx (business hours: Mon-Fri 9AM-6PM KST)

### Reporting Issues

When reporting issues, please include:
1. **Excel Version**: File → Account → About Excel
2. **Error Message**: Screenshot or exact text
3. **Workflow Step**: Which step (1-13) you were performing
4. **Data Volume**: Approximate number of accounts/entities
5. **Last Successful Step**: Which steps completed successfully

### Feature Requests

To request new features or enhancements:
1. Email pwcda@pwc.com with subject: "[HRE Feature Request]"
2. Describe the feature and business need
3. Provide example use case
4. Indicate priority (High/Medium/Low)

---

## Appendix: Keyboard Shortcuts

| Shortcut | Function |
|----------|----------|
| `Ctrl+1` | Format Cells dialog |
| `Ctrl+F` | Find in current sheet |
| `Ctrl+H` | Find and Replace |
| `Alt+Down` | Open dropdown in cell with data validation |
| `F5` | Go To dialog (navigate to specific cell) |
| `Ctrl+Home` | Go to cell A1 |
| `Ctrl+End` | Go to last used cell |

---

## Appendix: Workflow Checklist

Print this checklist for reference during consolidation:

- [ ] **Step 1**: Configure SharePoint connection
- [ ] **Step 2**: Refresh query data from SharePoint
- [ ] **Step 3**: Highlight PTB (identify unmapped accounts)
- [ ] **Step 4**: Filter yellow rows
- [ ] **Step 5**: Input CoA mappings (review auto-suggestions)
- [ ] **Step 6**: Finalize CoA mappings (validate and commit)
- [ ] **Step 7**: Verify financial statement sums
- [ ] **Step 8**: Highlight ADBS accounts
- [ ] **Step 9**: Input ADBS CoA mappings
- [ ] **Step 10**: Sync CoA master
- [ ] **Step 11**: Process MC accounts
- [ ] **Step 12**: Update exchange rates (평균환율 + 기말환율)
- [ ] **Step 13**: Export data

---

**© 2026 Samil PwC. All rights reserved.**

**HRE 연결마스터 v1.00**
