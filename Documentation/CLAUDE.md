# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a VBA-based Excel application for HRE (Hanwha Renewable Energy) consolidated financial statements, developed for PwC. The tool manages Chart of Accounts (CoA) mapping for multi-entity consolidation, handling data synchronization, variant-aware mapping (internal transactions), exchange rate integration, verification, and reporting for financial consolidation work.

**Version**: 1.00
**License**: © Samil PwC. All rights reserved.
**Primary Language**: VBA (Visual Basic for Applications)
**Base System**: Adapted from BEP v1.98

## Architecture

### Core Design Pattern

The codebase follows a **modular architecture** with clear separation of concerns, inherited from BEP with HRE-specific enhancements:

1. **Numbered Modules (mod_01 to mod_17)**: Sequential workflow steps (mod_17 is HRE-specific)
2. **Entity-Specific Modules**: Worksheet-level business logic (e.g., `CoAMaster_code.bas`, `BSPL_code.bas`)
3. **UserForms**: UI components for data entry and configuration (`.frm` files)
4. **Worksheet Event Handlers**: Class modules for worksheet-specific behavior
5. **Shared Utilities**: `mod_10_Public.bas` contains global constants, helpers, and UI utilities

### Key Components

#### Workflow Modules (Sequential Processing)
- `mod_01_FilterSearch.bas` - Table filtering and search functionality
- `mod_02_FilterSearch_Master.bas` - Master table filtering
- `mod_03_PTB_CoA_Input.bas` - **[HRE ENHANCED]** Pre-Trial Balance CoA input with variant detection
- `mod_04_IntializeProgress.bas` - Progress tracking initialization
- `mod_05_PTB_Highlight.bas` - Pre-Trial Balance highlighting
- `mod_06_VerifySum.bas` - Financial statement verification and sum checks
- `mod_07_ADBS_Highlight.bas` - Acquisition/Disposal BS highlighting
- `mod_08_ADBS_CoA_Input.bas` - AD BS CoA input
- `mod_09_CheckMaster.bas` - Master data validation
- `mod_10_Public.bas` - **[HRE MODIFIED]** Global utilities, constants (AppName="HRE", version 1.00)
- `mod_11_Sync.bas` - CoA synchronization
- `mod_12_MC_Highlight.bas` - Management Consolidation highlighting
- `mod_13_MC_AD_Highlight.bas` - MC Acquisition/Disposal highlighting
- `mod_14_MC_CoA_Input.bas` - MC CoA input
- `mod_15_MC_CoA_AD_Input.bas` - MC AD CoA input
- `mod_16_Export.bas` - Data export functionality
- `mod_17_ExchangeRate.bas` - **[HRE NEW]** KEB Hana Bank exchange rate integration

#### Core Shared Modules
- `mod_10_Public.bas` - Application-wide utilities, constants, color schemes, permission validation
- `mod_Log.bas` - Usage logging via Google Forms (audit trail)
- `mod_Ribbon.bas` - Custom ribbon interface event handlers
- `mod_MouseWheel.bas` - Mouse wheel functionality for forms

#### Entity Worksheet Modules
- `CoAMaster_code.bas` - CoA master data management with double-click/right-click event handlers
- `BSPL_code.bas`, `ADBS_code.bas` - Balance Sheet and P&L sheets
- `CorpCoA_code.bas`, `CorpMaster_code.bas` - Corporate-level data
- `현재_통합_문서_code.bas` - Workbook-level events (Open, BeforeClose, sheet protection)

#### UserForms (UI Components)
- `frm*.frm` - Interactive forms for data management (CoA alterations, corporate appends, filters, etc.)
- Forms use VBA's MSForms controls for data entry and validation

### Data Flow Architecture

```
SPO (SharePoint) → Query Refresh → Raw_CoA Table
                                       ↓
                    PTB (Pre-Trial Balance) → Variant-Aware CoA Mapping → CoA_Input Table
                                       ↓                                          ↓
                              Verification & Highlighting                Exchange Rate Integration
                                       ↓                                          ↓
                                 Export/Reporting ←─────────────────────────────┘
```

**Key Tables**:
- `Master` - PwC standardized CoA master
- `Raw_CoA` - Client's chart of accounts from SPO (with variant support)
- `PTB` - Pre-Trial Balance data
- `CoA_Input` - Mapping input table (variant-aware)
- `AD_BS`, `AD_PL_비경상적` - Acquisition/Disposal tables
- `환율정보(평균)`, `환율정보(일자)` - **[HRE]** Exchange rate sheets

### State Management

Workflow state tracked in `Check` worksheet:
- Each workflow step (Cells 12-21, column 4) has status: "Complete", "In Progress", or empty
- **Row 20**: Exchange rate status (Step 12 in workflow)
- Progress tracked with timestamps and user info
- `isYellow`, `isYellow_ADBS` - Global variables for highlighting state

## Key Patterns and Conventions

### Password Protection
- **Worksheet Password**: `BEP1234` (constant `PASSWORD`)
- **Workbook Password**: `PwCDA7529` (constant `PASSWORD_Workbook`)
- All sheets protected with `UserInterfaceOnly:=True` to allow VBA manipulation
- Use `Unprotect PASSWORD` before edits, then `Protect PASSWORD, UserInterfaceOnly:=True` after

### Performance Optimization Pattern
```vba
Call SpeedUp          ' Disable screen updating, events, calculation
' ... perform operations ...
Call SpeedDown        ' Re-enable screen updating, events, calculation
```

### Progress Tracking Pattern
```vba
Call OpenProgress("Operation description")
Call CalculateProgress(0.5, "Halfway done...")
Call CalculateProgress(1, "Complete")  ' Auto-closes progress form
```

### Error Handling
- Most procedures use `On Error Resume Next` (permissive error handling)
- Critical failures use `GoEnd "Error message"` which displays message and terminates

### Table-Based Data Management
- Excel `ListObject` tables used throughout
- Dictionary objects (`Scripting.Dictionary`) for efficient lookups
- Array-based operations for performance optimization

### Color Coding (PwC Brand Colors)
- `PwCRed()`, `PwCOrg()`, `PwCYlw()`, `PwCGreen()`, `PwCBlue()` - Brand colors
- `vbYellow` highlighting indicates missing/incomplete CoA mappings
- `RGB(0, 176, 80)` green indicates processed/matched rows
- `RGB(198, 239, 206)` light green indicates "Complete" status in Check sheet

## CoA Master Data Setup

### HRE-Specific CoA Structure

HRE uses a **5-digit base code system** with **variant suffixes** to handle different transaction types.

See the full CoA master data structure in the reference file `coa.md` located in VBA_Export folder.

#### Variant Types

**BASE Variant** (no suffix):
- Standard accounts without special handling
- Example: `10300` → `111206` (Cash - Operating - CNY)

**INTERCO_KR Variant** (`_내부거래` suffix):
- Korean internal transaction accounts
- Example: `11401_내부거래` → `112800` (Interco Receivable)
- Example: `25301_내부거래` → `212800` (Interco Payable)

**INTERCO_IC Variant** (`_IC` suffix):
- International internal transaction accounts
- Example: `11602_IC` → `112800` (Interco Receivable)

**CONSOLIDATION Variant** (`MC` prefix):
- Management consolidation accounts (excluded from auto-mapping)
- Handled separately in MC modules

### 5-Digit Matching Logic

The system uses base code matching with variant-aware fallback:

1. Extract base code (first 5 digits before underscore)
2. Detect variant type from suffix
3. Try exact variant match
4. Fallback to BASE variant if no specific match
5. Leave empty for manual review if no match

## Exchange Rate Module (mod_17)

### KEB Hana Bank API Integration

**Security Note**: Uses trusted KEB Hana Bank official API. HTML parsing is performed on trusted content from the bank's API using the htmlfile COM object which is isolated from browser security contexts.

#### Average Exchange Rates (평균환율)
```vba
Call GetER_Flow()  ' Triggered by ribbon button "평균환율 조회"
```
- **Purpose**: P&L account conversions (income/expense over a period)
- **User Input**: Start date and end date selection via `frmCalendar`
- **Data Source**: KEB Hana Bank API for average rates
- **Output**: `환율정보(평균)` sheet with average rates for all currencies
- **Check Sheet Update**: Row 20, Status "Complete"

#### Spot Exchange Rates (기말환율)
```vba
Call GetER_Spot()  ' Triggered by ribbon button "기말환율 조회"
```
- **Purpose**: Balance Sheet account conversions (point-in-time balances)
- **User Input**: Single date selection via `frmCalendar`
- **Data Source**: KEB Hana Bank API for spot rates
- **Output**: `환율정보(일자)` sheet with spot rates for all currencies
- **Check Sheet Update**: Row 20, Status "Complete"

#### Special Currency Handling

**환산 (Conversion Factor)**:
- **Standard**: USD, EUR, CNY, etc. → `환산 = 1`
- **Special (per 100 units)**: JPY, VND, IDR → `환산 = 100`

**KRW Baseline**: Automatically added with `환산=1`, `매매기준율=1`

## Variant Detection Pattern

### Implementation in mod_03_PTB_CoA_Input.bas

**Dictionary Structure**:
- Nested dictionary: `baseCode -> variantType -> [Account, Description]`

**Variant Detection Functions**:
```vba
' GetBaseCode - Extract 5-digit base code
Private Function GetBaseCode(accountCode As String) As String
    ' "11401_내부거래" → "11401"
End Function

' GetVariantType - Detect variant type
Private Function GetVariantType(accountCode As String) As String
    ' "_내부거래" → "INTERCO_KR"
    ' "_IC" → "INTERCO_IC"
    ' "MC*" → "CONSOLIDATION"
    ' Else → "BASE"
End Function
```

**Auto-Mapping Logic**: Multi-tier lookup with exact variant match → BASE fallback → empty

## Common Development Tasks

### Testing CoA Mapping Flow (with Variants)
1. Ensure SPO connection configured
2. Run query refresh
3. Filter yellow rows
4. Open CoA input (auto-populates with variant-aware mapping)
5. Verify variant mappings
6. Complete mappings

### Testing Exchange Rate Integration
1. Click "평균환율 조회" button
2. Select date range
3. Verify exchange rate sheet populated
4. Check special currencies (JPY, VND, IDR) have 환산=100
5. Verify KRW baseline present

### Extending Variant Types
1. Add new variant suffix to `coa.md`
2. Update `GetVariantType` function
3. Dictionary auto-builds from `Raw_CoA` table

## Important Constants and Global State

**Application Metadata** (HRE-specific):
- `AppName` = "HRE"
- `AppType` = "연결마스터"
- `AppVersion` = "1.00"

**Date Management**:
- `GetClosingYear()`, `GetClosingMonth()`
- `RelDate()` - Release date (2026-01-21)
- `ExpDate()` - Expiration date (2030-12-31)

**Permission Validation** (HRE-specific):
- `IsPermittedEmail()` - Checks @pwc.com, @bepsolar.com, or @hre.com domains

## Working with UserForms

### Form Positioning
- Use `Call FramePosition(formObject)` to center forms
- Forms auto-branded with `AppName & " " & AppType`

### HRE-Specific Forms

**frmCalendar Enhancements**:
- `GetDate(position, monthsAhead)` method
- Date validation: Only allows dates up to today
- January 1st holiday handling

**frmCorp_Append Enhancements**:
- Supports variant suffix input

## Worksheet References

**Key Worksheets**:
- `CoAMaster` - CoA Master
- `CorpCoA` - Corporate CoA (with variant support)
- `BSPL` - Balance Sheet & P&L
- `Check` - Workflow status tracking (Row 20: Exchange rate status)
- `환율정보(평균)` - **[HRE NEW]** Average exchange rates
- `환율정보(일자)` - **[HRE NEW]** Spot exchange rates

## External Dependencies

- SharePoint/SPO Integration
- Google Forms Logging
- Windows API (User32.dll, Gdi32.dll)
- Outlook COM
- **KEB Hana Bank API** - **[HRE NEW]** Exchange rate data
- **MSXML2.ServerXMLHTTP** - **[HRE NEW]** HTTP client
- **htmlfile COM Object** - **[HRE NEW]** HTML parsing

## Development Best Practices

### Memory Management
- Always `Set obj = Nothing`
- Free HTTP and HTML objects after exchange rate operations

### Array Operations
- Use `Dictionary` for O(1) lookups
- Nested dictionaries for variant-aware mappings

### Code Organization
- Variant-related functions prefixed with `Get`
- Follow module numbering (mod_18+ for future features)

### Exchange Rate Best Practices
- Date validation (not in future)
- Holiday handling
- API error handling with `On Error Resume Next`
- Sheet cleanup before refresh

---

**HRE-Specific Summary**:
- **Version 1.00** (base: BEP 1.98)
- **New Module**: mod_17_ExchangeRate (KEB Hana Bank integration)
- **Enhanced Module**: mod_03_PTB_CoA_Input (variant-aware mapping)
- **CoA Structure**: 5-digit base codes with variant suffixes
- **Exchange Rates**: Average (P&L) and Spot (B/S) with special currency handling
- **Permitted Domains**: @pwc.com, @bepsolar.com, @hre.com
