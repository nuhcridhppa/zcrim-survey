# Apps Script setup: Officer Portal, new Family Report, Declared Income

Open the **Project Declared** sheet, then **Extensions → Apps Script**.

## 1. Add three files (copy from this folder)

| In the editor click | Name it | Paste the contents of |
|---|---|---|
| **＋ → Script** | `ReportV2` | [`ReportV2.gs`](ReportV2.gs) |
| **＋ → HTML** | `ReportV2` | [`ReportV2.html`](ReportV2.html) |
| **＋ → Script** | `Officer` | [`Officer.gs`](Officer.gs) |

If `ReportV2` already exists, replace its contents. `ReportV2.gs` changed, so the portal can use it too.

## 2. Six small edits in `Code.gs`

Use **Ctrl+F** to find each line, then change it as shown.

> **Easier:** ask Claude for the finished `Code.gs` with every edit below already applied, then replace the whole file (Ctrl+A, paste).
> That version also lets officers log in with a **username**: `dcrim` and `adc`. Their passwords are set in the Passwords sheet.

### Edit 1: add the ADC account (inside `var ZCRIMS = [`)
Find:
```js
  { name:"DCRIM",        mobile:"9812529666", email:"farhan461@gmail.com",          block:"ALL (Observer)", dcrim:true },
```
Add this line **right below it**, using the ADC's real mobile and email:
```js
  { name:"ADC",          mobile:"ADC_MOBILE",  email:"ADC_EMAIL",                   block:"ALL (Officer)", adc:true },
```

### Edit 2: let ADC log in from a PC (in `loginZcrim`)
Find:
```js
  if (isDesktop && !zcrim.master && !zcrim.dcrim) {
```
Replace with:
```js
  if (isDesktop && !zcrim.master && !zcrim.dcrim && !zcrim.adc) {
```

### Edit 3: allow the portal's functions (in `handleApiRequest`)
Find:
```js
                   "uploadPhoto","submitSurvey","skipFamily","getLeaderboard","getDashboardData","flagResurvey","getDcrimProgress"];
```
Replace with:
```js
                   "uploadPhoto","submitSurvey","skipFamily","getLeaderboard","getDashboardData","flagResurvey","getDcrimProgress",
                   "getOfficerOverview","getVillageFamilies","searchFamilies","getFamilyReport"];
```
Then find:
```js
      getDcrimProgress:    function(a){ return getDcrimProgress(a[0]); }
```
Replace with:
```js
      getDcrimProgress:    function(a){ return getDcrimProgress(a[0]); },
      getOfficerOverview:  function(a){ return getOfficerOverview(a[0]); },
      getVillageFamilies:  function(a){ return getVillageFamilies(a[0], a[1], a[2]); },
      searchFamilies:      function(a){ return searchFamilies(a[0], a[1]); },
      getFamilyReport:     function(a){ return getFamilyReport(a[0], a[1], a[2]); }
```

### Edit 4: save Declared Income (in `submitSurvey`, two places)
Find:
```js
      remarks:    cellSafe(sanitizeStr(m.remarks,   200))
    };
```
Replace with:
```js
      remarks:    cellSafe(sanitizeStr(m.remarks,   200)),
      declared:   String(m.declaredIncome || "").replace(/[^\d]/g, "").slice(0, 9)
    };
```
Then find:
```js
      memSh.appendRow([ts,f.id,f.name,f.block,f.ward,idx+1,m.name,m.age,m.income,m.occupation,m.remarks,payload.zcrimName,surveyTag]);
```
Replace with:
```js
      memSh.appendRow([ts,f.id,f.name,f.block,f.ward,idx+1,m.name,m.age,m.income,m.occupation,m.remarks,payload.zcrimName,surveyTag,m.declared]);
```

### Edit 5: column heading for Declared Income (in `ensureSheetsExist`)
Find:
```js
    "MemberData": ["Timestamp","FamilyID","FamilyName","Block","Ward","MemberNo","MemberName","Age","IncomeRange","Occupation","Remarks","ZCRIM","SurveyTag"],
```
Replace with:
```js
    "MemberData": ["Timestamp","FamilyID","FamilyName","Block","Ward","MemberNo","MemberName","Age","IncomeRange","Occupation","Remarks","ZCRIM","SurveyTag","DeclaredIncome"],
```

### Edit 6: keep officer accounts off the leaderboards (two places)
Find **both** of these (in `getDashboardData` and `getLeaderboard`):
```js
ZCRIMS.filter(function(z){ return !z.master && !z.test; })
```
Replace each with:
```js
ZCRIMS.filter(function(z){ return !z.master && !z.dcrim && !z.adc && !z.test; })
```

### Add the menu item for the new report (in `onOpen`)
Find:
```js
    .addItem("🖨️ Search & Print Family Report (PDF)","searchAndPrintReport")
```
Add below it:
```js
    .addItem("🖨️ Family Report (new)","showFamilyReportV2")
```

## 3. Save, set up, publish

1. Click **💾 Save**.
2. Back in the Sheet, reload the page, then run **⚙️ ZCRIM Admin → 🏗️ Setup / Fix All Sheets**. This adds the `DeclaredIncome` heading.
3. In Apps Script: **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy**. Keep the same deployment so the app's link doesn't change.
4. Give the ADC a password: **⚙️ ZCRIM Admin → 🔄 Reset Password for One ZCRIM**, then enter the ADC's mobile. The password is emailed to them.

## Who sees what

| Account | Where it opens after login | Can use a PC |
|---|---|---|
| Admin (`master:true`) | Officer Portal, plus a **📝 Survey** button | ✅ |
| DCRIM (`dcrim:true`) | Officer Portal | ✅ |
| ADC (`adc:true`) | Officer Portal | ✅ |
| Field ZCRIM | Village list → survey | ❌ phone only |

The Officer Portal has three tabs:
- **📊 Progress:** district totals, 14-day trend, block bars and the ZCRIM leaderboard.
- **🏘️ Villages:** block → village → every family, marked surveyed, pending, re-survey or absent.
- **🔍 PPP Search:** find any PPP ID or name and open the family's report, which can be printed or saved as PDF.

All officers see all blocks. The portal only reads data; officers cannot change it.
