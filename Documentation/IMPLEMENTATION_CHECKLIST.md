# HRE 연결마스터 - Implementation Checklist

**Version**: 1.00
**Document Date**: 2026-01-21
**Target System**: Excel 2016+ on Windows 10+

This checklist ensures complete and correct deployment of the HRE Consolidation Master application.

---

## Table of Contents

1. [Pre-Implementation Checklist](#pre-implementation-checklist)
2. [Module Import Instructions](#module-import-instructions)
3. [UserForm Import Instructions](#userform-import-instructions)
4. [Ribbon Installation](#ribbon-installation)
5. [Power Query Setup](#power-query-setup)
6. [Testing Checklist](#testing-checklist)
7. [Deployment Checklist](#deployment-checklist)

---

## Pre-Implementation Checklist

### System Requirements Verification

- [ ] **Excel Version**: Excel 2016 or later installed
  - Verify: File → Account → About Excel
  - Minimum build: 16.0.4266.1001

- [ ] **Windows Version**: Windows 10 or later
  - Verify: Settings → System → About
  - Required for proper DPI scaling and API calls

- [ ] **VBA Support**: Visual Basic for Applications enabled
  - Verify: Developer tab visible in ribbon
  - If not: File → Options → Customize Ribbon → Check "Developer"

- [ ] **Macro Security**: Set to "Disable all macros with notification"
  - Path: File → Options → Trust Center → Trust Center Settings → Macro Settings
  - Select: "알림과 함께 모든 매크로 제외"

- [ ] **Network Access**: Internet connection available
  - Required for SharePoint queries and exchange rate updates
  - Test: Access https://www.kebhana.com

- [ ] **Permissions**: User has admin rights (for initial setup)
  - Required for: Installing CustomUI Editor, modifying workbook structure

### Software Tools

- [ ] **CustomUI Editor** (for ribbon installation)
  - Download: https://github.com/fernandreu/office-custom-ui/releases
  - Version: 2.8.0 or later
  - Alternative: Manual ZIP method (see Ribbon Installation section)

- [ ] **VBA Editor**: Accessible via Alt+F11
  - Verify: Press Alt+F11 in Excel
  - Should open VBA project explorer

- [ ] **7-Zip or WinRAR** (for manual ribbon method)
  - Only needed if CustomUI Editor not available
  - Download: https://www.7-zip.org

### Backup Preparation

- [ ] **Create Baseline File**: Save blank Excel workbook as `HRE_연결마스터_v1.00.xlsm`
  - File format: Excel Macro-Enabled Workbook (.xlsm)
  - Location: Secure network drive or SharePoint

- [ ] **Backup Copy**: Create `HRE_연결마스터_v1.00_BACKUP.xlsm` before any imports
  - Keep in separate folder for rollback if needed

- [ ] **Source Files Ready**: Ensure VBA_Export folder contains all modules
  - Location: `/Users/jaewookim/Desktop/Project/HRE/작업/VBA_Export/`
  - Count: 42+ VBA files (.bas, .frm, .frx)

---

## Module Import Instructions

### Standard Modules (22 modules)

Import these modules in the order listed to maintain dependencies:

#### Phase 1: Utility Modules (Import First)

1. **mod_10_Public.bas**
   - [ ] Open VBA Editor (Alt+F11)
   - [ ] Right-click VBAProject → Import File
   - [ ] Select `mod_10_Public.bas`
   - [ ] Verify import: Check module appears in Modules folder
   - [ ] **Critical**: This must be imported first (contains global constants)

2. **mod_Log.bas**
   - [ ] Import file `mod_Log.bas`
   - [ ] Verify `LogData` and `LogData_Access` functions exist

3. **mod_MouseWheel.bas**
   - [ ] Import file `mod_MouseWheel.bas`
   - [ ] Verify mouse wheel API declarations

4. **mod_Ribbon.bas**
   - [ ] Import file `mod_Ribbon.bas`
   - [ ] Verify ribbon callback procedures exist

#### Phase 2: Workflow Modules (Sequential Import)

5. **mod_01_FilterSearch.bas**
   - [ ] Import file
   - [ ] Verify `FilterPTB` procedure

6. **mod_02_FilterSearch_Master.bas**
   - [ ] Import file
   - [ ] Verify `FilterMaster` procedure

7. **mod_03_PTB_CoA_Input.bas** ⭐ **[HRE ENHANCED]**
   - [ ] Import file
   - [ ] Verify `Fill_Input_Table` procedure exists
   - [ ] **Important**: Check for variant detection functions:
     - `GetBaseCode()`
     - `GetVariantType()`
   - [ ] Verify nested dictionary structure in `Fill_Input_Table`

8. **mod_04_IntializeProgress.bas**
   - [ ] Import file
   - [ ] Verify `InitializeProgress` procedure

9. **mod_05_PTB_Highlight.bas**
   - [ ] Import file
   - [ ] Verify `HighlightPTB` procedure

10. **mod_06_VerifySum.bas**
    - [ ] Import file
    - [ ] Verify `VerifySum` procedure

11. **mod_07_ADBS_Highlight.bas**
    - [ ] Import file
    - [ ] Verify `HighlightADBS` procedure

12. **mod_08_ADBS_CoA_Input.bas**
    - [ ] Import file
    - [ ] Verify `Fill_ADBS_Input_Table` procedure

13. **mod_09_CheckMaster.bas**
    - [ ] Import file
    - [ ] Verify `CheckMaster` procedure

14. **mod_11_Sync.bas**
    - [ ] Import file
    - [ ] Verify `SyncCoA` procedure

15. **mod_12_MC_Highlight.bas**
    - [ ] Import file
    - [ ] Verify `HighlightMC` procedure

16. **mod_13_MC_AD_Highlight.bas**
    - [ ] Import file
    - [ ] Verify `HighlightMC_AD` procedure

17. **mod_14_MC_CoA_Input.bas**
    - [ ] Import file
    - [ ] Verify `Fill_MC_Input_Table` procedure

18. **mod_15_MC_CoA_AD_Input.bas**
    - [ ] Import file
    - [ ] Verify `Fill_MC_AD_Input_Table` procedure

19. **mod_16_Export.bas**
    - [ ] Import file
    - [ ] Verify `ExportData` procedure

20. **mod_17_ExchangeRate.bas** ⭐ **[HRE NEW]**
    - [ ] Import file
    - [ ] **Important**: Verify KEB Hana Bank integration:
      - `GetER_Flow()` procedure (average rates)
      - `GetER_Spot()` procedure (spot rates)
      - `GetHtmlFlow()` function
      - `GetHtmlSpot()` function
      - `UpdateCheckStatus()` function
    - [ ] Verify security note comments are present (line 14-17)

#### Phase 3: Helper Modules

21. **mod_OpenPage.bas**
    - [ ] Import file
    - [ ] Verify `OpenPage` procedure

22. **mod_QueryProtection.bas**
    - [ ] Import file
    - [ ] Verify query protection logic

23. **mod_Refresh.bas**
    - [ ] Import file
    - [ ] Verify `QueryRefresh` procedure

24. **mod_z_Module_GetCursor.bas**
    - [ ] Import file
    - [ ] Verify cursor position API functions

### Import Verification Checklist

After importing all modules:

- [ ] **Module Count**: 24 standard modules in Modules folder
- [ ] **No Compile Errors**: VBA Editor → Debug → Compile VBAProject (no errors)
- [ ] **Global Constants**: Verify in Immediate Window (Ctrl+G):
  ```vba
  ? AppName         ' Should return "HRE"
  ? AppVersion      ' Should return "1.00"
  ? AppType         ' Should return "연결마스터"
  ```

---

## UserForm Import Instructions

### UserForm Files (20 forms with .frx binary files)

**Important**: Each .frm file has a corresponding .frx file (binary storage for controls). Both must be in the same directory for successful import.

#### Phase 1: Core Forms

1. **FrmProgress.frm** + **FrmProgress.frx**
   - [ ] Import `FrmProgress.frm`
   - [ ] Verify `.frx` file auto-imported
   - [ ] Open form in design view
   - [ ] Verify controls: `LblMsg`, `FraRate`, `ProgressBar`, `FraBar`
   - [ ] Close form

2. **frmCalendar.frm** + **frmCalendar.frx** ⭐ **[HRE CRITICAL]**
   - [ ] Import `frmCalendar.frm`
   - [ ] Verify `.frx` file auto-imported
   - [ ] **Important**: This form is used for exchange rate date selection
   - [ ] Verify `GetDate()` method exists
   - [ ] Verify calendar controls and date validation logic

3. **frmFilter.frm** + **frmFilter.frx**
   - [ ] Import `frmFilter.frm`
   - [ ] Verify filter controls (ComboBox, ListBox)

4. **frmFilter_Master.frm** + **frmFilter_Master.frx**
   - [ ] Import `frmFilter_Master.frm`
   - [ ] Verify master filter controls

#### Phase 2: CoA Management Forms

5. **frmCoA_Alter.frm** + **frmCoA_Alter.frx**
   - [ ] Import form
   - [ ] Verify CoA alteration controls

6. **frmCoA_Delete.frm** + **frmCoA_Delete.frx**
   - [ ] Import form
   - [ ] Verify deletion confirmation logic

7. **frmCoA_Update.frm** + **frmCoA_Update.frx**
   - [ ] Import form
   - [ ] Verify update controls

8. **frmMaster_Alter.frm** + **frmMaster_Alter.frx**
   - [ ] Import form
   - [ ] Verify master alteration controls

#### Phase 3: Corporate Management Forms

9. **frmCorp_Alter.frm** + **frmCorp_Alter.frx**
   - [ ] Import form
   - [ ] Verify corporate alteration controls

10. **frmCorp_Append.frm** + **frmCorp_Append.frx** ⭐ **[HRE ENHANCED]**
    - [ ] Import form
    - [ ] **Important**: Verify variant suffix support
    - [ ] Check validation logic for `_내부거래`, `_IC` suffixes
    - [ ] Verify append controls

#### Phase 4: Utility Forms

11. **frmSPO.frm** + **frmSPO.frx**
    - [ ] Import form
    - [ ] Verify SharePoint connection controls

12. **frmDate.frm** + **frmDate.frx**
    - [ ] Import form
    - [ ] Verify date picker controls

13. **frmDirectory.frm** + **frmDirectory.frx**
    - [ ] Import form
    - [ ] Verify directory browser controls

14. **frmScope.frm** + **frmScope.frx**
    - [ ] Import form
    - [ ] Verify scope selection controls

#### Phase 5: People Management Forms

15. **frmPeople.frm** + **frmPeople.frx**
    - [ ] Import form
    - [ ] Verify people list controls

16. **frmAddPerson.frm** + **frmAddPerson.frx**
    - [ ] Import form
    - [ ] Verify person addition controls

17. **frmEditPerson.frm** + **frmEditPerson.frx**
    - [ ] Import form
    - [ ] Verify person edit controls

### UserForm Import Verification

After importing all forms:

- [ ] **Form Count**: 17+ forms in Forms folder
- [ ] **No Missing .frx Files**: Each form has corresponding .frx
- [ ] **No Design Errors**: Open each form in design view (F7), verify no "missing control" errors
- [ ] **Test Form Display**: Run test macro:
  ```vba
  Sub TestFormDisplay()
      frmCalendar.Show
  End Sub
  ```
  - [ ] Form displays correctly
  - [ ] Controls are visible and properly positioned
  - [ ] Form can be closed without errors

---

## Ribbon Installation

Choose **either** Method A (recommended) or Method B (manual).

### Method A: Using CustomUI Editor (Recommended)

#### Prerequisites

- [ ] CustomUI Editor installed (version 2.8.0+)
- [ ] HRE workbook saved and closed

#### Steps

1. **Open Workbook in CustomUI Editor**
   - [ ] Right-click `HRE_연결마스터_v1.00.xlsm`
   - [ ] Select "Open with" → CustomUI Editor
   - [ ] If not in context menu, launch CustomUI Editor and use File → Open

2. **Add Custom UI XML**
   - [ ] In CustomUI Editor, click "Insert" → "Office 2010+ Custom UI Part"
   - [ ] Paste the following XML into the editor:

```xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="HRETab" label="HRE 연결마스터">

        <!-- Data Management Group -->
        <group id="DataGroup" label="데이터 관리">
          <button id="btnSPO" label="SPO 연결" size="large"
                  imageMso="DatabaseConnect"
                  onAction="ShowSPODialog"/>
          <button id="btnQueryRefresh" label="쿼리 새로 고침" size="large"
                  imageMso="RefreshArrows"
                  onAction="QueryRefresh"/>
          <button id="btnCoASync" label="CoA 동기화" size="large"
                  imageMso="DatabaseSynchronize"
                  onAction="SyncCoA"/>
        </group>

        <!-- Filtering Group -->
        <group id="FilterGroup" label="필터">
          <button id="btnFilterPTB" label="PTB 필터" size="large"
                  imageMso="Filter"
                  onAction="FilterPTB"/>
          <button id="btnFilterMaster" label="Master 필터" size="large"
                  imageMso="FilterAdvanced"
                  onAction="FilterMaster"/>
        </group>

        <!-- CoA Input Group -->
        <group id="CoAInputGroup" label="CoA 입력">
          <button id="btnPTBInput" label="PTB CoA 입력" size="large"
                  imageMso="TableInsert"
                  onAction="Fill_Input_Table"/>
          <button id="btnPTBComplete" label="PTB CoA 완료" size="large"
                  imageMso="AcceptTableStyleChanges"
                  onAction="Fill_CoA_Table"/>
          <button id="btnADBSInput" label="ADBS CoA 입력" size="large"
                  imageMso="TableInsert"
                  onAction="Fill_ADBS_Input_Table"/>
        </group>

        <!-- Verification Group -->
        <group id="VerifyGroup" label="검증">
          <button id="btnPTBHighlight" label="PTB CoA 확인" size="large"
                  imageMso="ReviewHighlightColor"
                  onAction="HighlightPTB"/>
          <button id="btnADBSHighlight" label="ADBS CoA 확인" size="large"
                  imageMso="ReviewHighlightColor"
                  onAction="HighlightADBS"/>
          <button id="btnVerifySum" label="재무제표 검증" size="large"
                  imageMso="DatabaseVerify"
                  onAction="VerifySum"/>
        </group>

        <!-- Exchange Rate Group (HRE NEW) -->
        <group id="ExchangeRateGroup" label="환율">
          <button id="btnERFlow" label="평균환율 조회" size="large"
                  imageMso="CurrencyConvert"
                  onAction="GetER_Flow"
                  screentip="평균환율 조회"
                  supertip="기간 평균환율을 KEB 하나은행에서 조회합니다. P&L 계정에 사용됩니다."/>
          <button id="btnERSpot" label="기말환율 조회" size="large"
                  imageMso="CurrencySymbol"
                  onAction="GetER_Spot"
                  screentip="기말환율 조회"
                  supertip="특정일 기말환율을 KEB 하나은행에서 조회합니다. B/S 계정에 사용됩니다."/>
        </group>

        <!-- MC Processing Group -->
        <group id="MCGroup" label="MC">
          <button id="btnMCHighlight" label="MC 하이라이트" size="large"
                  imageMso="ReviewHighlightColor"
                  onAction="HighlightMC"/>
          <button id="btnMCInput" label="MC CoA 입력" size="large"
                  imageMso="TableInsert"
                  onAction="Fill_MC_Input_Table"/>
        </group>

        <!-- Export Group -->
        <group id="ExportGroup" label="내보내기">
          <button id="btnExport" label="데이터 내보내기" size="large"
                  imageMso="ExportExcel"
                  onAction="ExportData"/>
        </group>

      </tab>
    </tabs>
  </ribbon>
</customUI>
```

3. **Validate XML**
   - [ ] Click "Validate" button in CustomUI Editor
   - [ ] Ensure no errors reported
   - [ ] If errors: Check XML syntax, ensure all tags are closed

4. **Save and Close**
   - [ ] Click "Save" in CustomUI Editor
   - [ ] Close CustomUI Editor
   - [ ] Open Excel workbook
   - [ ] Verify "HRE 연결마스터" tab appears in ribbon

5. **Callback Verification**
   - [ ] Verify all ribbon callbacks exist in `mod_Ribbon.bas`:
     - `ShowSPODialog`
     - `QueryRefresh`
     - `SyncCoA`
     - `FilterPTB`, `FilterMaster`
     - `Fill_Input_Table`, `Fill_CoA_Table`, `Fill_ADBS_Input_Table`
     - `HighlightPTB`, `HighlightADBS`, `VerifySum`
     - `GetER_Flow`, `GetER_Spot` ⭐ **[HRE NEW]**
     - `HighlightMC`, `Fill_MC_Input_Table`
     - `ExportData`

---

### Method B: Manual ZIP Method (Alternative)

Use this method if CustomUI Editor is not available.

#### Prerequisites

- [ ] 7-Zip or WinRAR installed
- [ ] HRE workbook saved and closed
- [ ] Text editor (Notepad++ recommended)

#### Steps

1. **Change File Extension**
   - [ ] Rename `HRE_연결마스터_v1.00.xlsm` to `HRE_연결마스터_v1.00.zip`
   - [ ] If Windows hides extensions: File Explorer → View → Options → View tab → Uncheck "Hide extensions"

2. **Extract ZIP Contents**
   - [ ] Right-click ZIP file → Extract All
   - [ ] Extract to folder: `HRE_연결마스터_TEMP`

3. **Create customUI Folder**
   - [ ] Navigate to extracted folder
   - [ ] Create new folder: `customUI` (case-sensitive)
   - [ ] Inside `customUI`, create file: `customUI14.xml`

4. **Paste XML Content**
   - [ ] Open `customUI14.xml` in text editor
   - [ ] Paste the XML from Method A above
   - [ ] Save and close file

5. **Update [Content_Types].xml**
   - [ ] In extracted root folder, open `[Content_Types].xml`
   - [ ] Add this line before `</Types>`:
     ```xml
     <Override PartName="/customUI/customUI14.xml" ContentType="application/xml"/>
     ```
   - [ ] Save and close file

6. **Update .rels**
   - [ ] Navigate to `_rels` folder
   - [ ] Open `.rels` file in text editor
   - [ ] Add this line before `</Relationships>`:
     ```xml
     <Relationship Id="customUIRelID" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>
     ```
   - [ ] Save and close file

7. **Repackage ZIP**
   - [ ] Select all files/folders in extracted directory (not the folder itself)
   - [ ] Right-click → 7-Zip → Add to archive
   - [ ] Archive format: ZIP
   - [ ] Compression level: Store (no compression)
   - [ ] Create archive: `HRE_연결마스터_v1.00.zip`

8. **Restore .xlsm Extension**
   - [ ] Rename `HRE_연결마스터_v1.00.zip` back to `HRE_연결마스터_v1.00.xlsm`

9. **Test**
   - [ ] Open Excel workbook
   - [ ] Check for "HRE 연결마스터" tab in ribbon
   - [ ] If error: "Errors were detected in file..." → Re-check XML syntax

---

## Power Query Setup

### SharePoint Connection Configuration

#### Prerequisites

- [ ] SharePoint site URL available
- [ ] User has Read access to SharePoint list
- [ ] Network connectivity to SharePoint

#### Method 1: Using SPO Dialog (Recommended)

1. **Launch SPO Form**
   - [ ] Open HRE workbook
   - [ ] Click ribbon: HRE 연결마스터 → SPO 연결
   - [ ] `frmSPO` dialog appears

2. **Configure Connection**
   - [ ] Enter SharePoint site URL:
     ```
     https://[company].sharepoint.com/sites/[sitename]
     ```
   - [ ] Click **확인**

3. **Authenticate**
   - [ ] If prompted, enter credentials
   - [ ] Select "Organizational account"
   - [ ] Sign in with @pwc.com or @hre.com email

4. **Verify Connection**
   - [ ] Check `Check` sheet Row 12: Should show "Complete"
   - [ ] Navigate to `CorpCoA` sheet
   - [ ] Verify `Raw_CoA` table structure exists

#### Method 2: Manual Power Query Setup

1. **Create New Query**
   - [ ] Data tab → Get Data → From Other Sources → Blank Query

2. **Enter M Code**
   - [ ] In Power Query Editor, click Advanced Editor
   - [ ] Paste M code:
     ```m
     let
         Source = SharePoint.Tables("https://[company].sharepoint.com/sites/[sitename]", [ApiVersion = 15]),
         #"Corporate_CoA_Table" = Source{[Id="Corporate_CoA"]}[Items],
         #"Renamed Columns" = Table.RenameColumns(#"Corporate_CoA_Table",{
             {"Corp_Code", "법인코드"},
             {"Account_Code", "계정코드"},
             {"Account_Name", "법인별 계정과목명"},
             {"PwC_CoA", "PwC_CoA"},
             {"PwC_Account_Name", "PwC_계정과목명"},
             {"Category", "분류"},
             {"Account", "Account"},
             {"Description", "Description"}
         })
     in
         #"Renamed Columns"
     ```

3. **Load to Table**
   - [ ] Click Close & Load To
   - [ ] Select "Table"
   - [ ] Worksheet: `CorpCoA`
   - [ ] Cell: `A4` (or first available cell)
   - [ ] Click OK

4. **Name Query**
   - [ ] In Queries & Connections pane, right-click query
   - [ ] Rename to: `Raw_CoA`

5. **Configure Refresh**
   - [ ] Data tab → Queries & Connections
   - [ ] Right-click `Raw_CoA` → Properties
   - [ ] Refresh control:
     - [ ] Uncheck "Enable background refresh"
     - [ ] Uncheck "Refresh data when opening the file"
     - [ ] Set "Refresh every" to 0 (manual refresh only)

---

## Testing Checklist

### Phase 1: Basic Functionality Tests

#### Test 1: Macro Initialization
- [ ] Open workbook
- [ ] Macros enabled automatically
- [ ] No error messages on open
- [ ] `Workbook_Open` event executes:
  - [ ] All sheets protected with `UserInterfaceOnly:=True`
  - [ ] Hidden sheets remain hidden (AddCoA, HideSheet)

#### Test 2: Ribbon Visibility
- [ ] "HRE 연결마스터" tab visible in ribbon
- [ ] All button groups visible:
  - [ ] 데이터 관리 (3 buttons)
  - [ ] 필터 (2 buttons)
  - [ ] CoA 입력 (3 buttons)
  - [ ] 검증 (3 buttons)
  - [ ] 환율 (2 buttons) ⭐ **[HRE NEW]**
  - [ ] MC (2 buttons)
  - [ ] 내보내기 (1 button)
- [ ] Button icons displayed correctly (no red X)

#### Test 3: Progress Bar
- [ ] Run test macro:
  ```vba
  Sub TestProgress()
      Call OpenProgress("Test message")
      Call CalculateProgress(0.5, "Halfway done")
      Application.Wait (Now + TimeValue("0:00:02"))
      Call CalculateProgress(1, "Complete")
  End Sub
  ```
- [ ] Progress form displays centered on Excel window
- [ ] Progress bar animates smoothly
- [ ] Form closes automatically at 100%

---

### Phase 2: CoA Mapping Tests (Variant Detection)

#### Test 4: Base Code Extraction
- [ ] Open VBA Editor (Alt+F11)
- [ ] Immediate Window (Ctrl+G)
- [ ] Run tests:
  ```vba
  ? GetBaseCode("10300")              ' Expected: "10300"
  ? GetBaseCode("11401_내부거래")     ' Expected: "11401"
  ? GetBaseCode("11602_IC")           ' Expected: "11602"
  ? GetBaseCode("25301_내부거래")     ' Expected: "25301"
  ```
- [ ] All results match expected values

#### Test 5: Variant Type Detection
- [ ] Immediate Window tests:
  ```vba
  ? GetVariantType("10300")           ' Expected: "BASE"
  ? GetVariantType("11401_내부거래")  ' Expected: "INTERCO_KR"
  ? GetVariantType("11602_IC")        ' Expected: "INTERCO_IC"
  ? GetVariantType("MC1234")          ' Expected: "CONSOLIDATION"
  ```
- [ ] All results match expected values

#### Test 6: Auto-Mapping with Variants
- [ ] Populate `PTB` table with test data:
  | 법인코드 | 법인별 CoA | 법인별 계정과목명 | 잔액 |
  |---------|-----------|----------------|-----|
  | 1000    | 10300     | 보통예금        | 1000000 |
  | 1000    | 11401_내부거래 | 단기대여금(내부) | 500000 |
  | 1000    | 11602_IC  | 미수수익(IC)    | 200000 |

- [ ] Click ribbon: PTB CoA 확인
- [ ] Yellow highlighting appears on all three rows
- [ ] Click ribbon: PTB 필터
- [ ] Only yellow rows visible
- [ ] Click ribbon: PTB CoA 입력
- [ ] `CoA_Input` table populated:
  - [ ] Row 1: `10300` → `111206` (Cash - Operating - CNY)
  - [ ] Row 2: `11401_내부거래` → `112800` (Interco Receivable)
  - [ ] Row 3: `11602_IC` → `112800` (Interco Receivable)
- [ ] Verify auto-suggestions are correct
- [ ] Click ribbon: PTB CoA 완료
- [ ] No validation errors
- [ ] PTB rows turn green
- [ ] `Raw_CoA` table updated with new mappings

---

### Phase 3: Exchange Rate Integration Tests

#### Test 7: Average Exchange Rate Retrieval
- [ ] Click ribbon: 평균환율 조회
- [ ] `frmCalendar` dialog appears
- [ ] Select start date: `2024-01-01`
- [ ] Click OK
- [ ] `frmCalendar` appears again for end date
- [ ] Select end date: `2024-12-31`
- [ ] Click OK
- [ ] Wait 5-10 seconds for retrieval
- [ ] Verify `환율정보(평균)` sheet created:
  - [ ] Header: "조회 기간 : 2024-01-01 ~ 2024-12-31"
  - [ ] Note: "※ 조회일이 토/일/공휴일..."
  - [ ] Table headers: 국가명 및 통화 | 통화 | 환산 | ...
  - [ ] Currency rows: USD, EUR, JPY, CNY, VND, IDR, etc.
  - [ ] Special currencies: JPY, VND, IDR have 환산=100
  - [ ] Last row: "대한민국 KRW" with 환산=1, 매매기준율=1
- [ ] Check sheet Row 20: Status "Complete"

#### Test 8: Spot Exchange Rate Retrieval
- [ ] Click ribbon: 기말환율 조회
- [ ] `frmCalendar` dialog appears
- [ ] Select date: `2024-12-31`
- [ ] Click OK
- [ ] Wait 5-10 seconds for retrieval
- [ ] Verify `환율정보(일자)` sheet created:
  - [ ] Header: "조회 기준일 : 2024-12-31"
  - [ ] Same structure as average rates
  - [ ] KRW baseline row present
- [ ] Check sheet Row 20: Status "Complete"

#### Test 9: Exchange Rate Error Handling
- [ ] **Test future date validation**:
  - [ ] Click ribbon: 기말환율 조회
  - [ ] Try to select tomorrow's date
  - [ ] Expected error: "유효하지 않은 날짜입니다..."

- [ ] **Test date range validation**:
  - [ ] Click ribbon: 평균환율 조회
  - [ ] Select start date: `2024-12-31`
  - [ ] Select end date: `2024-01-01` (before start)
  - [ ] Expected error: "유효하지 않은 종료 날짜입니다..."

- [ ] **Test January 1st handling**:
  - [ ] Click ribbon: 평균환율 조회
  - [ ] Select start date: `2024-01-01`
  - [ ] Select end date: `2024-01-31`
  - [ ] Verify header note: "...1월 1일은 공휴일이므로 1월 2일부터로 조회"

#### Test 10: Special Currency Conversion
- [ ] Open `환율정보(일자)` sheet
- [ ] Find JPY row:
  - [ ] 통화 column: `JPY`
  - [ ] 환산 column: `100`
  - [ ] 매매기준율 column: e.g., `1,100` (example rate)
- [ ] Manual calculation test:
  ```
  JPY 100,000 amount
  Rate: 1,100 KRW per 100 JPY
  환산: 100

  Formula: 100,000 × 1,100 × (1/100) = 1,100,000 KRW
  ```
- [ ] Verify formula logic is correct

---

### Phase 4: Workflow Integration Tests

#### Test 11: End-to-End Workflow (Steps 1-13)
- [ ] **Step 1**: SPO 연결 → Status "Complete"
- [ ] **Step 2**: 쿼리 새로 고침 → Raw_CoA populated
- [ ] **Step 3**: PTB CoA 확인 → Yellow highlighting
- [ ] **Step 4**: PTB 필터 → Only yellow rows
- [ ] **Step 5**: PTB CoA 입력 → Auto-suggestions populated
- [ ] **Step 6**: PTB CoA 완료 → Mappings saved, rows green
- [ ] **Step 7**: 재무제표 검증 → Verification passed
- [ ] **Step 8**: ADBS CoA 확인 → ADBS highlighted
- [ ] **Step 9**: ADBS CoA 입력 → ADBS mappings complete
- [ ] **Step 10**: CoA 동기화 → All tables synced
- [ ] **Step 11**: MC 처리 → MC accounts handled
- [ ] **Step 12**: 환율 조회 → Exchange rates updated
- [ ] **Step 13**: 데이터 내보내기 → Export successful

#### Test 12: Check Sheet Validation
- [ ] Open `Check` sheet
- [ ] Verify all status cells (Row 12-21, Column 4):
  - [ ] Row 12: SPO connection → "Complete"
  - [ ] Row 13: Query refresh → "Complete"
  - [ ] Row 14: PTB highlight → "Complete"
  - [ ] Row 15: PTB CoA input → "Complete"
  - [ ] Row 16: FS verification → "Complete"
  - [ ] Row 17: ADBS highlight → "Complete"
  - [ ] Row 18: ADBS CoA input → "Complete"
  - [ ] Row 19: CoA sync → "If Any"
  - [ ] Row 20: Exchange rate → "Complete" ⭐ **[HRE NEW]**
  - [ ] Row 21: MC processing → "If Any"
- [ ] Verify timestamps (Column 5) populated
- [ ] Verify user names (Column 6) populated

---

### Phase 5: Error Handling Tests

#### Test 13: Invalid Mapping Validation
- [ ] In `CoA_Input` table, leave PwC_CoA cell empty
- [ ] Click ribbon: PTB CoA 완료
- [ ] Expected result:
  - [ ] Yellow highlighting on empty cell
  - [ ] Error message: "PwC_CoA와 PwC_계정과목명 매칭되지 않은 항목이 있습니다."
  - [ ] Workflow stops (mappings not saved)

#### Test 14: Master Table Mismatch
- [ ] In `CoA_Input` table, enter invalid PwC code: `999999`
- [ ] Click ribbon: PTB CoA 완료
- [ ] Expected result:
  - [ ] Yellow highlighting on invalid code
  - [ ] Error message: "...매칭되지 않은 항목이 있습니다."

#### Test 15: Network Failure Handling
- [ ] Disconnect network
- [ ] Click ribbon: 평균환율 조회
- [ ] Select dates
- [ ] Expected result:
  - [ ] Graceful error handling (no crash)
  - [ ] User-friendly error message
  - [ ] Exchange rate sheet empty or shows last cached data

---

### Phase 6: Performance Tests

#### Test 16: Large Dataset Performance
- [ ] Populate `PTB` table with 1,000+ rows
- [ ] Click ribbon: PTB CoA 확인
- [ ] Measure time: Should complete in < 30 seconds
- [ ] Click ribbon: PTB CoA 입력
- [ ] Measure time: Auto-mapping should complete in < 60 seconds
- [ ] Verify memory usage: Task Manager → Excel.exe < 2 GB RAM

#### Test 17: Exchange Rate Performance
- [ ] Click ribbon: 평균환율 조회
- [ ] Select 1-year date range
- [ ] Measure time: Should complete in < 15 seconds
- [ ] Verify no hanging or freezing
- [ ] Verify Excel remains responsive during retrieval

---

## Deployment Checklist

### Pre-Deployment Validation

- [ ] **All tests passed**: Phases 1-6 complete with no failures
- [ ] **No VBA compile errors**: Debug → Compile VBAProject successful
- [ ] **No runtime errors**: Tested all ribbon buttons with real data
- [ ] **Documentation complete**: README.md and CLAUDE.md reviewed
- [ ] **Variant detection confirmed**: Test cases 4-6 passed
- [ ] **Exchange rate integration confirmed**: Test cases 7-10 passed

### File Preparation

- [ ] **Version check**: File name `HRE_연결마스터_v1.00.xlsm`
- [ ] **Passwords set**:
  - [ ] Worksheet password: `BEP1234`
  - [ ] Workbook password: `PwCDA7529`
- [ ] **Sheets protected**: All worksheets except exchange rate sheets
- [ ] **Hidden sheets**: `AddCoA`, `HideSheet` set to `xlSheetVeryHidden`
- [ ] **Workbook structure protected**: ThisWorkbook.Protect PASSWORD_Workbook

### Metadata Update

- [ ] **Application constants** verified in `mod_10_Public.bas`:
  ```vba
  AppName = "HRE"
  AppVersion = "1.00"
  AppType = "연결마스터"
  RelDate = DateSerial(2026, 1, 21)
  ExpDate = DateSerial(2030, 12, 31)
  ```
- [ ] **Permitted emails** include HRE domain:
  ```vba
  permittedDomains = Array("@pwc.com", "@bepsolar.com", "@hre.com")
  ```

### Documentation Packaging

- [ ] **README.md**: Copied to SharePoint documentation folder
- [ ] **CLAUDE.md**: Copied to codebase repository
- [ ] **IMPLEMENTATION_CHECKLIST.md**: Copied to deployment package
- [ ] **coa.md**: Reference file included in VBA_Export folder

### SharePoint Deployment

- [ ] **Upload to SharePoint**:
  - [ ] Site: `https://[company].sharepoint.com/sites/HRE_Consolidation`
  - [ ] Library: `Shared Documents/Tools/`
  - [ ] File name: `HRE_연결마스터_v1.00.xlsm`

- [ ] **Set permissions**:
  - [ ] PwC users: Read/Edit
  - [ ] HRE users: Read only
  - [ ] Admins: Full Control

- [ ] **Create version history**:
  - [ ] Check in file with comment: "Initial release v1.00 - Exchange rate integration"
  - [ ] Enable version history tracking

### User Communication

- [ ] **Notification email sent** to:
  - [ ] PwC Digital Assurance team
  - [ ] HRE finance team
  - [ ] Key stakeholders

- [ ] **Email includes**:
  - [ ] Link to SharePoint file
  - [ ] Link to README.md (user guide)
  - [ ] Training session schedule (if applicable)
  - [ ] Support contact information

### Training Materials

- [ ] **Video tutorials recorded**:
  - [ ] Basic workflow (Steps 1-13)
  - [ ] CoA mapping with variants
  - [ ] Exchange rate usage (average vs. spot)
  - [ ] Troubleshooting common issues

- [ ] **Quick reference card created**:
  - [ ] 1-page PDF with workflow steps
  - [ ] Variant suffix reference table
  - [ ] Exchange rate selection guide

### Post-Deployment Monitoring

- [ ] **Week 1 check-in**:
  - [ ] Review usage logs (Google Forms)
  - [ ] Collect user feedback
  - [ ] Address any critical issues

- [ ] **Week 2 check-in**:
  - [ ] Review error reports
  - [ ] Verify exchange rate data accuracy
  - [ ] Check Check sheet completion rates

- [ ] **Month 1 review**:
  - [ ] Compile lessons learned
  - [ ] Plan v1.01 enhancements
  - [ ] Update documentation based on feedback

---

## Rollback Plan

### Emergency Rollback (Critical Issues)

If critical errors occur post-deployment:

1. **Immediate Actions**:
   - [ ] Notify users via email/Teams: "System temporarily offline for maintenance"
   - [ ] Remove file from SharePoint (or set to Read-Only for admins only)
   - [ ] Restore backup version: `HRE_연결마스터_v1.00_BACKUP.xlsm`

2. **Root Cause Analysis**:
   - [ ] Review error logs
   - [ ] Test in isolated environment
   - [ ] Identify specific failing component:
     - [ ] Variant detection logic?
     - [ ] Exchange rate API?
     - [ ] Ribbon callbacks?
     - [ ] SharePoint connection?

3. **Hotfix Deployment** (if issue is minor):
   - [ ] Create branch: `v1.00_hotfix`
   - [ ] Apply minimal fix
   - [ ] Test thoroughly (Phases 1-6)
   - [ ] Deploy with version: `v1.00.1`

4. **Full Rollback** (if issue is major):
   - [ ] Revert to BEP v1.98 (if available)
   - [ ] Notify users: "HRE-specific features temporarily disabled"
   - [ ] Plan v1.01 with comprehensive fixes

---

## Version Control

### File Naming Convention

- **Development**: `HRE_연결마스터_v1.00_DEV_YYYYMMDD.xlsm`
- **Testing**: `HRE_연결마스터_v1.00_TEST_YYYYMMDD.xlsm`
- **Production**: `HRE_연결마스터_v1.00.xlsm`
- **Backup**: `HRE_연결마스터_v1.00_BACKUP_YYYYMMDD.xlsm`

### Change Log Template

When creating v1.01+, maintain change log:

```markdown
## v1.00 (2026-01-21)
- Initial release
- Base: BEP v1.98
- NEW: mod_17_ExchangeRate (KEB Hana Bank integration)
- ENHANCED: mod_03_PTB_CoA_Input (variant-aware mapping)
- ADDED: _내부거래, _IC variant suffix support
- ADDED: 5-digit base code matching
- UPDATED: AppName="HRE", AppVersion="1.00"
- UPDATED: Permitted domains include @hre.com

## v1.01 (TBD)
- [Future enhancements]
```

---

## Sign-Off

### Implementation Team

- [ ] **Developer**: __________________ Date: __________
  - All modules imported correctly
  - All forms functional
  - All tests passed

- [ ] **QA Tester**: __________________ Date: __________
  - Test cases 1-17 completed
  - No critical issues found
  - Performance acceptable

- [ ] **Project Manager**: __________________ Date: __________
  - Documentation complete
  - Deployment ready
  - Users notified

### Approval

- [ ] **PwC Lead**: __________________ Date: __________
  - Approved for production deployment

- [ ] **HRE Stakeholder**: __________________ Date: __________
  - Accepted for use in HRE consolidation

---

**End of Implementation Checklist**

**HRE 연결마스터 v1.00**
**© 2026 Samil PwC. All rights reserved.**
