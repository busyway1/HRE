# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**HRE Consolidation Master** - An Excel VBA + Power Query system for automated consolidated financial statement preparation. Adapted from BEP v1.98 for HRE Group's specific consolidation requirements.

- **Platform**: Microsoft Excel 2016+ (Microsoft 365 recommended)
- **Language**: VBA (Visual Basic for Applications) + Power Query M Language
- **Data Source**: SharePoint Online Lists
- **Version**: 1.00

## Architecture

### Core Components

```
VBA_Export/
├── mod_10_Public.bas      # Global constants, utilities, error handling, PwC colors
├── mod_03_PTB_CoA_Input.bas   # CoA First Drafting with variant detection
├── mod_17_ExchangeRate.bas    # KEB Hana Bank exchange rate API
├── mod_Ribbon.bas         # Custom ribbon callback handlers
├── 현재_통합_문서_code.bas    # ThisWorkbook events (protection, logging)
└── [Other modules]        # Filter, sync, export, verification logic

UserForms/                 # VBA UserForms (.frm + .frx pairs)
├── frmCalendar.frm       # Date picker
├── frmSPO.frm            # SharePoint URL configuration
├── frmCorp_Append.frm    # Add corporation
└── [Other forms]

PowerQuery/               # Power Query M code templates
├── PowerQuery_PTB.m      # SharePoint PTB list connection
├── PowerQuery_RawCoA.m   # CoA mapping history
└── PowerQuery_PTB_LocalFile.m  # Local file fallback
```

### Data Flow

1. **Power Query** fetches PTB (Pre-Trial Balance) from SharePoint List
2. **VBA (mod_03)** performs First Drafting - auto-maps accounts using 5-digit base codes + variant detection
3. **User** reviews/corrects mappings via AddCoA sheet
4. **VBA (mod_05/06)** validates mappings and generates consolidated statements

### Key Concepts

**5-Digit CoA Matching**: HRE uses first 5 digits of account codes for matching (e.g., `11401_내부거래` → `11401`)

**Variant Types**: Account suffixes indicate special handling:
- `_내부거래` → INTERCO_KR (intercompany Korean)
- `_IC` → INTERCO_IC (intercompany Group)
- `MC` prefix → CONSOLIDATION (excluded from standard mapping)
- No suffix → BASE

**Sheet CodeNames**: VBA references worksheets directly by CodeName (set in VBA Editor Properties):
- `BSPL`, `AddCoA`, `CorpCoA`, `CoAMaster`, `Check`, `Verify`, `HideSheet`, `CorpMaster`, `CorpBSPL`

## Critical Setup Requirements

### VBA Environment

1. **Set Sheet CodeNames** in VBA Editor (Alt+F11 → F4 Properties):
   - Set `(Name)` property (NOT `Name`) to match expected CodeNames
   - Required: BSPL, AddCoA, CorpCoA, CoAMaster, Check, Verify, HideSheet, CorpMaster, CorpBSPL

2. **Windows Encoding**: VBA files must be CP949 encoded for Korean text. macOS-exported UTF-8 files need conversion before import.

3. **Option Explicit**: All modules require `Option Explicit`

### Passwords

```vba
' Sheet protection password
PASSWORD = "BEP1234"

' Workbook structure protection
PASSWORD_Workbook = "PwCDA7529"
```

## Key Functions Reference

### mod_03_PTB_CoA_Input.bas
- `Fill_Input_Table()` - First Drafting: auto-populate CoA_Input table with suggested mappings
- `Fill_CoA_Table()` - Finalize mappings, validate against Master, update Raw_CoA
- `GetBaseCode(accountCode)` - Extract first 5 digits
- `GetVariantType(accountCode)` - Detect variant suffix

### mod_17_ExchangeRate.bas
- `GetER_Flow()` - Fetch average exchange rates for P&L (period range)
- `GetER_Spot()` - Fetch spot rates for B/S (single date)
- Uses KEB Hana Bank API via MSXML2.ServerXMLHTTP

### mod_10_Public.bas
- `SpeedUp()` / `SpeedDown()` - Toggle calculation/events/screen updating
- `GoEnd(msg)` - Error handling with cleanup
- `Msg(str, style)` - Branded message box
- `GetClosingYear()` / `GetClosingMonth()` - Read from HideSheet

### mod_Ribbon.bas
- Callback handlers for custom ribbon: `SetSPO`, `SetDate`, `Update`, `GetER_Flow`, etc.

## Power Query Notes

Power Query templates in `/PowerQuery_*.m` files:
- Replace `[SHAREPOINT_SITE_URL]` and `[LIST_NAME]` placeholders
- Default SharePoint site: `https://pwckor.sharepoint.com/sites/KR-ASR-HRE_Consolidation`
- Queries output to ListObjects (Excel Tables) referenced by VBA

## Differences from BEP v1.98

| Feature | BEP | HRE |
|---------|-----|-----|
| CoA Matching | Exact match | 5-digit + variant |
| Exchange Rate | Manual input | API auto-fetch |
| MC (Management Consolidation) | Supported | Removed |
| Intercompany Detection | Manual | Auto-detect suffixes |

## Korean Encoding

VBA Editor on Windows requires CP949 encoding. If Korean appears as `������`:
1. Export files as UTF-8 from macOS
2. Convert to CP949 using Python: `convert_to_cp949_windows.py`
3. Re-import into VBA Editor

## File Import Order

When setting up the project in a new Excel workbook:
1. Import all `.bas` modules via File → Import File
2. Copy worksheet code (`.bas` files with sheet names) into corresponding sheet modules
3. Set all sheet CodeNames in Properties window
4. Compile VBA project (Debug → Compile VBAProject)
5. Configure ribbon XML using Custom UI Editor

## Debugging Tips

- Compile check: Debug → Compile VBAProject
- "Variable not defined" error usually means missing CodeName
- For API issues, check `MSXML2.ServerXMLHTTP` compatibility
- Power Query connection failures: verify SharePoint authentication
