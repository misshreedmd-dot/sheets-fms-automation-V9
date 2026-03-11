# 📊 sheets-fms-automation

> **A fully automated Flow Management System built on Google Sheets + Google Apps Script.**
> Track orders through multiple departments, auto-generate forms, calculate TAT (Turnaround Time), monitor delays in real-time, and diagnose issues — all without writing a single formula manually.

---

## 🚀 What Is This?

`sheets-fms-automation` is a production-ready **Google Apps Script** that transforms a blank Google Sheet into a complete **order flow tracking system**. It is designed for businesses that move work through multiple departments (e.g. Casting → CAD → Finishing → Dispatch) and need to track:

- When each step was **planned** to be done
- When it was **actually** done
- How much **time was delayed** at each step
- Who is responsible and what method they use

Everything is driven by **Google Forms** — no manual data entry into the sheet. Forms are auto-generated, triggers are auto-configured, and formulas are batch-written in seconds.

---

## ✨ Key Features at a Glance

| Category | Feature |
|---|---|
| 🧙 Setup Wizard | Step-by-step dialogs to build your entire FMS — no coding needed |
| ⚡ Batch Formula Write | Writes 1,000 rows of formulas in ~3 seconds — timeout impossible |
| ▶️ Resume on Timeout | Saves progress — resume exactly where it stopped if interrupted |
| 📋 Auto Form Generation | Creates Form 1 (new orders) + Form 2 (step updates) automatically |
| 📑 Form 2 Sections | Form 2 has a separate page per step with navigation — no irrelevant fields shown |
| ➕ Extra Columns per Step | Each step can have custom columns (CAD File, Revision, Approval, etc.) |
| 🕐 TAT Formula Engine | 4 TAT types: Hours / Days / Fixed Time / Whenever — respects working hours, weekends, holidays |
| 🔍 Diagnostic Checker | Full health check with PASS/FAIL/WARN report — find any issue in seconds |
| 🎨 Auto Styling | Color-coded delays: 🔴 late / 🟢 on time — applied automatically |
| 🏖️ Holiday Support | Holidays sheet — all formulas skip holiday dates automatically |

---

## 📦 When To Use This

Use `sheets-fms-automation` when:

- Your business processes **orders through multiple steps/departments**
- You want to track **planned vs actual completion time** per step
- You want **Google Forms** to feed data into a tracker automatically
- You need **TAT (Turnaround Time)** calculated in working hours or days
- You want to see **which orders are delayed** at a glance
- You need **custom data captured per step** (e.g. file uploads, revision numbers, approvals)
- You want a system that is easy to **diagnose and fix** when something goes wrong

**Example use cases:**
- 🏭 Manufacturing order tracking (Casting → Machining → QC → Dispatch)
- 💎 Jewellery production flow (Design → CAD → Casting → Polishing → Dispatch)
- 🖨️ Print job workflow (Order → Design → Proof → Print → Delivery)
- 🏗️ Project milestone tracking
- 📦 Supply chain order management

---

## 🗂️ Sheet Structure

```
Row 1   — Hidden config: NOW() | Opening Time | Closing Time | Working Days | Field Defs JSON | Step Defs JSON
Rows 2–3 — Hidden: Step numbers and IDs (Step1, s1, Step2, s2...)
Rows 4–7 — Hidden: What / Who / How / When per step
Row 8   — Hidden: TAT reference values
Row 9   — HEADERS (frozen, dark blue)
Row 10+ — DATA (1000 rows, all formulas pre-written)

Col A        = Unique ID (auto-incremented)
Col B        = Timestamp (form submission time)
Col C+       = User-defined fields (from wizard)
Col 100      = Field definitions (JSON, hidden)
Col 101      = Step definitions with exact column positions (JSON, hidden)
```

**Each step occupies a variable number of columns:**
```
Normal step:   Planned | Actual | Status | Time Delay           (4 cols)
Step + 1 extra: Planned | Actual | CAD File | Status | Time Delay  (5 cols)
Step + 2 extra: Planned | Actual | CAD File | Revision | Status | Time Delay  (6 cols)
```

---

## ⏱️ TAT Formula Types

| Type | Trigger | Example |
|---|---|---|
| **H — Hours** | N working hours after previous step's Actual | `3` hours after Casting done |
| **D — Days** | N working days after previous step's Actual | `1` day after approval |
| **T — Fixed Time** | Always by a fixed time of day | Every day by `18:00` |
| **N — Whenever** | No formula — blank, filled manually | Ad-hoc steps |

All TAT formulas respect:
- ✅ Opening and closing time
- ✅ Working days (Mon–Sun pattern configurable)
- ✅ Public holidays (Holidays sheet)
- ✅ Empty row guards (no errors on blank rows)
- ✅ IFERROR wrapping (no visible formula errors)

---

## 📋 Forms

### Form 1 — New Order Entry
- Fully dynamic — fields match exactly what you defined in the wizard
- Supported field types: **Text, Date, Dropdown**
- On submit: row added to FMS with Unique ID + Timestamp + all field values
- Duplicate check: skips if same identifier already exists

### Form 2 — Step Status Update
- **Page 1:** Identifier field + Step selection dropdown
- **One section per step** — selecting a step navigates to that step's page
- Each step's section shows only **that step's extra columns** (if any)
- Every section has a **Status = Done** field
- On submit: writes Done to Status col + Actual timestamp + all extra col values directly to FMS row

---

## ➕ Extra Columns per Step

Each step in the wizard can have **zero or more extra columns** between Actual and Status.

Supported extra column types:

| Type | Form widget | FMS storage |
|---|---|---|
| **Text** | Text input | String |
| **Date** | Date picker | Date formatted dd/MM/yyyy |
| **Number** | Text input | Float |
| **Dropdown** | Dropdown list | Selected value |

**Example:**
```
Step: CAD
Extra columns:
  → CAD File    (Text)
  → Revision    (Dropdown: R1, R2, R3)
  → Approved By (Text)

FMS columns: Planned | Actual | CAD File | Revision | Approved By | Status | Time Delay
```

Column positions are calculated at build time and stored in JSON (col 101). The Form 2 handler reads exact column numbers — no recalculation needed.

---

## 🔍 FMS Diagnostic Checker

Run **BMP Formulas → 🔍 Run FMS Diagnostic** at any time to get a full health report.

The diagnostic checks **6 sections** and writes results to an **"FMS Diagnostics"** sheet:

| Section | What is checked |
|---|---|
| **Sheets** | FMS / Holidays / Form Links / Form 1 response / Form 2 response sheets exist |
| **Config** | A1=NOW(), opening time C1, closing time D1, working days E1, iterative calc on |
| **Definitions** | Field defs (col 100) and step defs with col positions (col 101) found and parseable |
| **Steps** | Per step: ID in row 3, name in row 4, Planned formula, Time Delay formula, all headers, extra col headers |
| **Data** | Duplicate identifiers, missing IDs, missing timestamps, Done rows without Actual timestamp |
| **Triggers** | onChange_new + onFormSubmit_Router both exist, no duplicate triggers |

**Report output:**
```
✅ PASS  — green row  — everything correct
❌ FAIL  — red row    — fix required, detail shown
⚠️ WARN  — yellow row — minor issue or not set up yet
```

Summary alert shows total PASS / FAIL / WARN count. Every check is also written to **Apps Script Executions** log for deeper debugging.

---

## ▶️ Resume Step Writing

If step writing times out (rare with batch write, but possible with 15 steps):

1. Click **BMP Formulas → ▶️ Resume Step Writing**
2. It shows: *"4 of 10 steps written — resume from step 5?"*
3. Click Yes — continues exactly from where it stopped
4. Clears automatically when all steps are done

Progress is saved to `PropertiesService` after every step — no data is lost.

---

## ⚡ Batch Formula Write (v9)

Previous versions wrote formulas row by row — 1 API call per cell. For 1,000 rows × 5 steps × 3 formula columns = **15,000 API calls** — causing 6-minute timeouts.

v9 builds the full array of 1,000 formulas first, then writes the entire column in **one API call**:

```
Before v9:  15,000 API calls → potential timeout
After  v9:  15 API calls     → ~3 seconds, timeout impossible
```

---

## 🛠️ Installation & Setup

### Prerequisites
- A Google account
- Google Sheets (any blank sheet)
- No coding knowledge required

---

### Step 1 — Paste the Script

1. Open a **new blank Google Sheet**
2. Click **Extensions → Apps Script**
3. Delete all existing code
4. Paste the entire contents of `FMS_COMPLETE_SCRIPT_v9.js`
5. Click 💾 **Save** (Ctrl+S)
6. Click **Run → onOpen** to authorize
7. Click **Review Permissions → Allow**
8. Go back to your sheet and press **F5** to refresh

The **BMP Formulas** menu will appear in the menu bar.

---

### Step 2 — First Time Setup

```
BMP Formulas → Run This First Time After Install
```

Enter:
- **Opening Time** — e.g. `10:00`
- **Closing Time** — e.g. `18:00`
- **Working Days** — 7-character pattern (Mon to Sun), `0` = working, `1` = off
  - `0000011` = Saturday + Sunday off
  - `0000001` = Sunday off only
  - `0000111` = Friday + Saturday + Sunday off

✅ Creates the **Holidays sheet** automatically.

---

### Step 3 — Create FMS Sheet

```
BMP Formulas → 🆕 Create New FMS Sheet
```

Enter:
- Opening and closing time for this FMS
- **Number of fields** for new order entry (e.g. `3`)
- For each field: **Name** + **Type** (Text / Date / Dropdown)
  - If Dropdown: enter choices comma-separated

**Example:**
```
Field 1: PO Number    → Text
Field 2: Client Name  → Text
Field 3: Due Date     → Date
```

✅ FMS sheet created with your fields.

---

### Step 4 — Add Steps

```
BMP Formulas → ➕ Add Steps to FMS
```

Enter number of departments (1–15). For each step:

| Prompt | Example |
|---|---|
| Department name | `CAD` |
| Who is responsible | `Priya` |
| Method | `System` |
| TAT Type | `H` |
| TAT Value (if H/D/T) | `4` |
| Extra columns? | `Yes` |
| → Column name | `CAD File` |
| → Data type | `Text` |
| → Another extra col? | `Yes` |
| → Column name | `Revision` |
| → Data type | `Dropdown` |
| → Choices | `R1, R2, R3` |
| → Another extra col? | `No` |

> ⚠️ If it times out mid-way → click **▶️ Resume Step Writing**

✅ Steps written to FMS sheet with all formulas.

---

### Step 5 — Generate Forms

```
BMP Formulas → 📋 Generate Forms for Active FMS
```

- Make sure the **FMS sheet tab is selected** before clicking
- Wait ~15 seconds

✅ Creates:
- **Form 1** — New Order Entry
- **Form 2** — Step Status Update (with step sections and navigation)
- Links saved to **Form Links** sheet

---

### Step 6 — Setup Triggers

```
BMP Formulas → ⚙️ Setup All Triggers (Run Once)
```

✅ Creates exactly **2 triggers**:
1. `onChange_new` — auto-timestamps when Done is typed manually
2. `onFormSubmit_Router` — handles both Form 1 and Form 2

> ⚠️ Go to **Extensions → Apps Script → Triggers** (clock icon) and confirm only 2 exist. Delete any duplicates.

---

### Step 7 — Add Holidays (Optional)

Go to the **Holidays** sheet → add dates in column A from row 3 downward.

Format: `dd/MM/yyyy` — e.g. `26/01/2025`

---

### Step 8 — Run Diagnostic

```
BMP Formulas → 🔍 Run FMS Diagnostic
```

Fix any ❌ FAIL rows shown in the **FMS Diagnostics** sheet. Re-run until all are ✅ PASS.

---

## 📅 Daily Usage

```
New order arrives
  └─ Staff submits Form 1 (New Order Entry)
       └─ Row appears in FMS automatically with ID + Timestamp + all fields

Step completed
  └─ Staff submits Form 2 (Status Update)
       └─ Selects step → fills extra columns (if any) → Status = Done
            └─ FMS updates: Actual timestamp + Done + extra col values written

Monitor
  └─ Time Delay column: 🔴 red = late | 🟢 green = on time
  └─ Planned vs Actual visible for every step of every order
```

---

## 🔧 Troubleshooting

| Problem | Solution |
|---|---|
| BMP Formulas menu not showing | Refresh page (F5) |
| Steps timed out halfway | Click ▶️ Resume Step Writing |
| Form not updating FMS | Run 🔍 Diagnostic → check Triggers section |
| Actual time not stamping | Check Apps Script Executions for errors |
| Duplicate trigger firing | Extensions → Apps Script → Triggers → delete extras |
| Wrong column data written | Run 🔍 Diagnostic → check Step definitions |
| Any unknown issue | Run 🔍 FMS Diagnostic → read all ❌ FAIL rows |

---

## 📁 Files

```
sheets-fms-automation/
├── FMS_COMPLETE_SCRIPT_v9.js   ← Main script — paste this into Apps Script
└── README.md                   ← This file
```

---

## 🔄 Version History

| Version | Changes |
|---|---|
| v9 | Batch formula write, resume on timeout, extra cols per step, Form 2 sections, Diagnostic checker, Actual timestamp fix |
| v8 | IFERROR guards, row-by-row Planned formulas, single router trigger, Holidays sheet |
| v7 | Dynamic fields, ARRAYFORMULA Actual, WORKDAY.INTL guards |

---

## 📄 License

MIT License — free to use, modify, and distribute.

---

## 🙏 Contributing

Pull requests welcome. If you find a bug or want a new feature, open an issue describing your use case.

---

*Built with ❤️ using Google Apps Script — no external dependencies, no paid tools, runs entirely inside Google Workspace.*
