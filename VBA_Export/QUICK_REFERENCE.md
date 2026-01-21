# HRE Module Migration - Quick Reference

## ✅ Ready to Use (No Changes Needed)
- mod_01_FilterSearch.bas
- mod_02_FilterSearch_Master.bas
- mod_04_IntializeProgress.bas
- mod_05_PTB_Highlight.bas
- mod_09_CheckMaster.bas
- mod_11_Sync.bas
- mod_MouseWheel.bas
- mod_QueryProtection.bas
- mod_Refresh.bas
- mod_z_Module_GetCursor.bas
- Module1.bas

## ⚠️ Needs URL Updates
**mod_Log.bas**
```vba
' Line 22: Update HRE usage logging form
formUrl = "https://docs.google.com/forms/d/e/YOUR_HRE_FORM_ID/formResponse"

' Line 43: Update HRE access logging form
formUrl = "https://docs.google.com/forms/d/e/YOUR_HRE_ACCESS_FORM_ID/formResponse"
```

**mod_OpenPage.bas**
```vba
' Line 17: Update HRE Google Form URL
URL = "https://docs.google.com/forms/YOUR_HRE_FORM"

' Line 23: Update HRE Notion manual URL
URL = "https://www.notion.so/HRE-Manual-YOUR_PAGE_ID"
```

## ⚠️ Needs Code Completion
**mod_06_VerifySum.bas**
- Only RefreshPivotVerify() and VerifyBS() migrated
- Add from BEP source:
  - VerifyIS()
  - ValidateCorpCodes()
  - ValidateSheetColors()
  - IsValidColor()

**mod_16_Export.bas**
- Verify sheet names array matches actual HRE workbook
- Current list (line 36):
```vba
sheetNames = Array("계정 마스터", "CoA 마스터", "법인별 CoA", "합계 BSPL", "검증", "취득, 처분 BSPL", "연결관리대장", "연결관리대장(처분)")
```

## Import Order Recommendation
1. mod_10_Public.bas (already exists - foundation)
2. Utility modules (mod_Log, mod_MouseWheel, etc.)
3. Workflow modules in sequence (mod_01 → mod_16)

## Key Constants (Already Set in mod_10_Public)
```vba
AppName = "HRE"
AppType = "연결마스터"
AppVersion = "1.00"
PASSWORD = "BEP1234"
PASSWORD_Workbook = "PwCDA7529"
```

## Check Sheet Row Mapping (Rows 12-23)
All workflow steps reference Check.Cells(row, 4):
- Row 20 = Verification step (same as BEP structure)
