"""
Power BI Dashboard Build Guide Generator
=========================================
Run this script to produce PowerBI_Build_Guide.html — a step-by-step
reference you can keep open in your browser while building in Power BI Desktop.

    python generate_pbi_guide.py
"""

GUIDE_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>HRIS Power BI Build Guide</title>
<style>
  /* ── Base ── */
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    background: #f4f6f9;
    color: #1a1a2e;
    line-height: 1.7;
  }

  /* ── Layout ── */
  .sidebar {
    position: fixed; top: 0; left: 0;
    width: 240px; height: 100vh;
    background: #1a1a2e; color: #cdd3de;
    padding: 28px 16px; overflow-y: auto;
    font-size: 0.82rem;
  }
  .sidebar h2 { color: #fff; font-size: 0.95rem; margin-bottom: 20px; letter-spacing: .5px; }
  .sidebar a {
    display: block; padding: 6px 10px; border-radius: 6px;
    color: #aab4c8; text-decoration: none; margin-bottom: 3px;
    transition: background .15s;
  }
  .sidebar a:hover, .sidebar a.active { background: #2d3561; color: #fff; }
  .sidebar .section-label {
    color: #5c6bc0; font-size: 0.72rem; font-weight: 700;
    text-transform: uppercase; letter-spacing: 1px;
    margin: 18px 0 6px 10px;
  }
  .main { margin-left: 240px; padding: 40px 48px; max-width: 980px; }

  /* ── Section headings ── */
  h1 { font-size: 2rem; font-weight: 800; color: #1a1a2e; margin-bottom: 6px; }
  .subtitle { color: #666; font-size: 1rem; margin-bottom: 40px; }
  h2 {
    font-size: 1.35rem; font-weight: 700; color: #1a1a2e;
    margin: 48px 0 16px; padding-bottom: 8px;
    border-bottom: 3px solid #5c6bc0;
  }
  h3 { font-size: 1.05rem; font-weight: 700; color: #2d3561; margin: 24px 0 10px; }
  h4 { font-size: 0.92rem; font-weight: 700; color: #555; margin: 18px 0 8px; }
  p  { margin-bottom: 12px; }

  /* ── Steps ── */
  .steps { counter-reset: step; list-style: none; padding: 0; }
  .steps li {
    counter-increment: step;
    position: relative; padding: 14px 16px 14px 56px;
    margin-bottom: 8px; background: #fff;
    border-radius: 8px; border: 1px solid #e4e8f0;
  }
  .steps li::before {
    content: counter(step);
    position: absolute; left: 14px; top: 14px;
    width: 28px; height: 28px; border-radius: 50%;
    background: #5c6bc0; color: #fff;
    font-weight: 700; font-size: 0.85rem;
    display: flex; align-items: center; justify-content: center;
  }

  /* ── DAX code blocks ── */
  .dax-block {
    background: #1e1e2e; border-radius: 10px;
    padding: 20px 24px; margin: 12px 0 20px;
    position: relative; overflow-x: auto;
  }
  .dax-block .label {
    font-size: 0.7rem; font-weight: 700; letter-spacing: 1px;
    color: #888; text-transform: uppercase; margin-bottom: 10px;
  }
  .dax-block pre {
    font-family: "Cascadia Code", "Fira Code", Consolas, monospace;
    font-size: 0.85rem; color: #cdd3de; white-space: pre; line-height: 1.6;
  }
  /* Simple DAX token colours */
  .kw  { color: #c792ea; }   /* keywords: CALCULATE, VAR, RETURN, IF */
  .fn  { color: #82aaff; }   /* functions: COUNTROWS, DIVIDE, AVERAGE */
  .str { color: #c3e88d; }   /* strings: "Active" */
  .num { color: #f78c6c; }   /* numbers */
  .ref { color: #89ddff; }   /* table[column] references */
  .cmt { color: #546e7a; font-style: italic; }  /* -- comments */

  .copy-btn {
    position: absolute; top: 14px; right: 14px;
    background: #2d3561; color: #aab4c8;
    border: none; border-radius: 5px; padding: 4px 12px;
    font-size: 0.75rem; cursor: pointer; transition: background .2s;
  }
  .copy-btn:hover { background: #5c6bc0; color: #fff; }

  /* ── Visual cards ── */
  .visual-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
    gap: 14px; margin: 16px 0;
  }
  .visual-card {
    background: #fff; border-radius: 10px;
    border: 1px solid #e4e8f0; padding: 16px;
  }
  .visual-card .icon { font-size: 1.5rem; margin-bottom: 8px; }
  .visual-card strong { display: block; font-size: 0.9rem; margin-bottom: 4px; }
  .visual-card p { font-size: 0.82rem; color: #666; margin: 0; }

  /* ── Relationship diagram ── */
  .rel-table {
    width: 100%; border-collapse: collapse; margin: 16px 0; font-size: 0.85rem;
  }
  .rel-table th {
    background: #2d3561; color: #fff;
    padding: 10px 14px; text-align: left;
  }
  .rel-table td { padding: 9px 14px; border-bottom: 1px solid #e4e8f0; background: #fff; }
  .rel-table tr:hover td { background: #f0f2ff; }
  .pill {
    display: inline-block; padding: 2px 10px; border-radius: 20px;
    font-size: 0.75rem; font-weight: 600;
  }
  .pill-many { background: #e8eaf6; color: #3949ab; }
  .pill-one  { background: #e8f5e9; color: #2e7d32; }
  .pill-active   { background: #e8f5e9; color: #2e7d32; }
  .pill-inactive { background: #fff3e0; color: #e65100; }

  /* ── Tips & warnings ── */
  .tip, .warn, .info {
    border-radius: 8px; padding: 14px 18px; margin: 14px 0; font-size: 0.88rem;
  }
  .tip  { background: #e8f5e9; border-left: 4px solid #2e7d32; color: #1b5e20; }
  .warn { background: #fff3e0; border-left: 4px solid #e65100; color: #bf360c; }
  .info { background: #e8eaf6; border-left: 4px solid #3949ab; color: #1a237e; }
  .tip::before  { content: "✅ TIP: ";  font-weight: 700; }
  .warn::before { content: "⚠️ IMPORTANT: "; font-weight: 700; }
  .info::before { content: "ℹ️ NOTE: "; font-weight: 700; }

  /* ── Page badges ── */
  .page-badge {
    display: inline-block; background: #5c6bc0; color: #fff;
    font-size: 0.75rem; font-weight: 700; letter-spacing: .5px;
    padding: 3px 12px; border-radius: 20px; margin-bottom: 10px;
  }

  /* ── Measure index table ── */
  .measure-table { width: 100%; border-collapse: collapse; font-size: 0.83rem; margin: 12px 0; }
  .measure-table th { background: #f0f2ff; color: #2d3561; padding: 8px 12px; text-align: left; border-bottom: 2px solid #c5cae9; }
  .measure-table td { padding: 7px 12px; border-bottom: 1px solid #e8eaf6; }
  .measure-table tr:hover td { background: #f5f6ff; }

  /* ── Visual build cards ── */
  .build-visual {
    background: #fff; border: 1px solid #e4e8f0; border-radius: 10px;
    padding: 20px 22px; margin: 18px 0;
  }
  .build-visual h4 { margin: 0 0 12px; font-size: 1rem; color: #1a1a2e; }
  .viztype-badge {
    display: inline-block; font-size: 0.72rem; font-weight: 700;
    letter-spacing: .5px; padding: 3px 10px; border-radius: 20px;
    background: #e8eaf6; color: #3949ab; margin-bottom: 10px;
  }
  .field-table { width: 100%; border-collapse: collapse; margin: 12px 0; font-size: 0.83rem; }
  .field-table th { background: #f5f6ff; color: #3949ab; padding: 7px 12px; text-align: left; border-bottom: 2px solid #c5cae9; }
  .field-table td { padding: 7px 12px; border-bottom: 1px solid #eee; vertical-align: top; }
  .field-table td:first-child { font-weight: 600; color: #555; width: 140px; white-space: nowrap; }
  .well-name { font-family: monospace; background: #f0f2ff; padding: 1px 6px; border-radius: 4px; font-size: 0.8rem; }
  .format-list { margin: 10px 0 0; padding-left: 18px; font-size: 0.85rem; }
  .format-list li { margin-bottom: 5px; }

  /* ── Canvas basics diagram ── */
  .pane-diagram {
    display: grid; grid-template-columns: 1fr 1fr 1fr;
    gap: 12px; margin: 16px 0;
  }
  .pane-box {
    background: #fff; border: 2px solid #c5cae9; border-radius: 8px;
    padding: 14px; text-align: center;
  }
  .pane-box .pane-icon { font-size: 1.8rem; margin-bottom: 6px; }
  .pane-box strong { display: block; font-size: 0.9rem; color: #2d3561; margin-bottom: 4px; }
  .pane-box p { font-size: 0.78rem; color: #666; margin: 0; }

  /* ── Print ── */
  @media print {
    .sidebar { display: none; }
    .main { margin-left: 0; }
    .copy-btn { display: none; }
  }
</style>
</head>
<body>

<!-- ═══════════════════════════════════════════════════════ SIDEBAR -->
<nav class="sidebar">
  <h2>HRIS Power BI Guide</h2>

  <div class="section-label">Setup</div>
  <a href="#step1">1. Load Data</a>
  <a href="#step2">2. Build Data Model</a>
  <a href="#step3">3. Create Date Table</a>
  <a href="#step4">4. Create Measures</a>

  <div class="section-label">Dashboard Pages</div>
  <a href="#page1">Page 1 — Headcount</a>
  <a href="#page2">Page 2 — Compensation</a>
  <a href="#page3">Page 3 — Attrition</a>
  <a href="#page4">Page 4 — Performance</a>
  <a href="#page5">Page 5 — Data Quality</a>

  <div class="section-label">Reference</div>
  <a href="#all-measures">All DAX Measures</a>
  <a href="#tips">Final Tips</a>
</nav>

<!-- ═══════════════════════════════════════════════════════ MAIN -->
<main class="main">

<h1>HRIS Power BI Dashboard</h1>
<p class="subtitle">Step-by-step build guide &mdash; keep this open in your browser while working in Power BI Desktop</p>

<div class="info">This guide assumes you have already run <strong>generate_hris.py</strong> and <strong>clean_hris.py</strong> and that <strong>HRIS_Dataset_Cleaned.xlsx</strong> is in your HRIS folder. Use the cleaned file — not the raw one.</div>

<!-- ═══════════════════════════════════ STEP 1: LOAD DATA -->
<h2 id="step1">Step 1 &mdash; Load Data into Power BI</h2>

<ol class="steps">
  <li>Open <strong>Power BI Desktop</strong> (download free at powerbi.microsoft.com if needed)</li>
  <li>Click <strong>Home &rarr; Get Data &rarr; Excel Workbook</strong></li>
  <li>Navigate to your HRIS folder and select <strong>HRIS_Dataset_Cleaned.xlsx</strong></li>
  <li>In the Navigator window, check all 5 tables:<br>
      &nbsp;&nbsp;☑ Employees &nbsp; ☑ Compensation &nbsp; ☑ Leave &nbsp; ☑ Performance &nbsp; ☑ Termination</li>
  <li>Click <strong>Load</strong> (not "Transform Data" — the cleaning is already done)</li>
  <li>Wait for the load to finish. You should see all 5 tables in the <strong>Fields</strong> pane on the right.</li>
</ol>

<div class="warn">If any column shows as "ABC" (text) when it should be a date, go to Power Query: Home &rarr; Transform Data, click the column, then change the type to "Date" in the Data Type dropdown.</div>

<!-- ═══════════════════════════════════ STEP 2: DATA MODEL -->
<h2 id="step2">Step 2 &mdash; Build the Data Model (Relationships)</h2>

<p>Click the <strong>Model view</strong> icon on the left sidebar (it looks like three connected boxes). This is where you tell Power BI how your tables are connected.</p>

<p>Create the following 4 relationships by dragging <strong>EmployeeID</strong> from each child table onto <strong>EmployeeID</strong> in Employees:</p>

<table class="rel-table">
  <tr>
    <th>From (child table)</th>
    <th>Column</th>
    <th>To (parent)</th>
    <th>Column</th>
    <th>Cardinality</th>
    <th>Status</th>
  </tr>
  <tr>
    <td>Compensation</td>
    <td>EmployeeID</td>
    <td>Employees</td>
    <td>EmployeeID</td>
    <td><span class="pill pill-many">Many</span> → <span class="pill pill-one">One</span></td>
    <td><span class="pill pill-active">Active</span></td>
  </tr>
  <tr>
    <td>Leave</td>
    <td>EmployeeID</td>
    <td>Employees</td>
    <td>EmployeeID</td>
    <td><span class="pill pill-many">Many</span> → <span class="pill pill-one">One</span></td>
    <td><span class="pill pill-active">Active</span></td>
  </tr>
  <tr>
    <td>Performance</td>
    <td>EmployeeID</td>
    <td>Employees</td>
    <td>EmployeeID</td>
    <td><span class="pill pill-many">Many</span> → <span class="pill pill-one">One</span></td>
    <td><span class="pill pill-active">Active</span></td>
  </tr>
  <tr>
    <td>Termination</td>
    <td>EmployeeID</td>
    <td>Employees</td>
    <td>EmployeeID</td>
    <td><span class="pill pill-many">Many</span> → <span class="pill pill-one">One</span></td>
    <td><span class="pill pill-active">Active</span></td>
  </tr>
</table>

<div class="tip">To create a relationship: in Model view, drag EmployeeID from Compensation onto EmployeeID in Employees. A line appears between the tables. Double-click it to verify the settings match the table above.</div>

<!-- ═══════════════════════════════════ STEP 3: DATE TABLE -->
<h2 id="step3">Step 3 &mdash; Create a Date Table</h2>

<p>A Date Table is the backbone of any time-based analysis in Power BI. Without it, measures like "Hires This Year" or "Monthly Attrition" won't work correctly.</p>

<p>We build this in <strong>Power Query</strong> using M language — this approach works regardless of your regional settings and avoids DAX locale issues entirely.</p>

<ol class="steps">
  <li>Click <strong>Home &rarr; Transform Data</strong> to open the Power Query Editor</li>
  <li>In Power Query, click <strong>Home &rarr; New Source &rarr; Blank Query</strong></li>
  <li>Click <strong>Home &rarr; Advanced Editor</strong></li>
  <li>Select all the text inside the editor, delete it, and paste the code below</li>
  <li>Click <strong>Done</strong></li>
  <li>In the <strong>Name</strong> box (top-left of the screen), rename the query from <em>Query1</em> to <strong>DateTable</strong></li>
  <li>Click <strong>Home &rarr; Close &amp; Apply</strong> to load it into Power BI</li>
</ol>

<div class="dax-block">
  <div class="label">Power Query (M Language) — Paste into Advanced Editor</div>
  <button class="copy-btn" onclick="copyCode(this)">Copy</button>
  <pre>let
    StartDate    = #date(2019, 1, 1),
    EndDate      = #date(2024, 12, 31),
    TotalDays    = Duration.Days(EndDate - StartDate) + 1,
    DateList     = List.Dates(StartDate, TotalDays, #duration(1, 0, 0, 0)),
    ToTable      = Table.FromList(DateList, Splitter.SplitByNothing(), {"Date"}),
    SetType      = Table.TransformColumnTypes(ToTable, {{"Date", type date}}),
    AddYear      = Table.AddColumn(SetType,      "Year",        each Date.Year([Date]),                          Int64.Type),
    AddMonthNum  = Table.AddColumn(AddYear,      "MonthNumber", each Date.Month([Date]),                         Int64.Type),
    AddMonthName = Table.AddColumn(AddMonthNum,  "MonthName",   each Date.ToText([Date], "MMMM"),                type text),
    AddMonthShort= Table.AddColumn(AddMonthName, "MonthShort",  each Date.ToText([Date], "MMM"),                 type text),
    AddQuarter   = Table.AddColumn(AddMonthShort,"Quarter",     each "Q" & Text.From(Date.QuarterOfYear([Date])),type text),
    AddYearMonth = Table.AddColumn(AddQuarter,   "YearMonth",   each Date.ToText([Date], "yyyy-MM"),             type text),
    AddWeekDay   = Table.AddColumn(AddYearMonth, "WeekDay",     each Date.ToText([Date], "dddd"),                type text)
in
    AddWeekDay</pre>
</div>

<div class="warn">Make sure you rename the query to <strong>DateTable</strong> (step 6 above) before clicking Close &amp; Apply. If you forget, right-click the table in the Fields pane and choose Rename.</div>

<ol class="steps" style="counter-reset: step 7;">
  <li>Back in Power BI, go to <strong>Model view</strong> and connect <strong>DateTable[Date]</strong> to <strong>Employees[HireDate]</strong></li>
  <li>Also connect <strong>DateTable[Date]</strong> to <strong>Termination[TerminationDate]</strong> — double-click the relationship line and set it to <strong>Inactive</strong></li>
  <li>Right-click <strong>DateTable</strong> in the Fields pane &rarr; <strong>Mark as date table</strong> &rarr; choose the <strong>Date</strong> column</li>
</ol>

<div class="tip">After loading the DateTable, sort MonthName by MonthNumber so months appear in calendar order in charts: click the <strong>MonthName</strong> column in Table view &rarr; Column Tools tab &rarr; Sort by Column &rarr; MonthNumber.</div>

<!-- ═══════════════════════════════════ STEP 4: MEASURES -->
<h2 id="step4">Step 4 &mdash; Create Your DAX Measures</h2>

<p>Measures are Power BI's calculation engine. They recalculate on the fly based on filters and slicers. Follow these steps for <strong>every measure</strong> in this guide:</p>

<ol class="steps">
  <li>In the <strong>Fields</strong> pane, right-click the <strong>Employees</strong> table &rarr; <strong>New measure</strong></li>
  <li>Delete the placeholder text, paste the DAX below, and press Enter</li>
  <li>Rename the measure if needed by clicking it in the Fields pane</li>
</ol>

<div class="warn">Create all measures inside the Employees table to keep them organized in one place. You can always move them later.</div>

<!-- ═══════════════════════════════════ CANVAS BASICS -->
<h2 id="page1">How the Power BI Canvas Works</h2>
<p>Before building anything, understand the three panes you will use constantly:</p>

<div class="pane-diagram">
  <div class="pane-box">
    <div class="pane-icon">📊</div>
    <strong>Visualizations Pane</strong>
    <p>Right side. Click an icon here to add a visual to the page. Below the icons are the <strong>field wells</strong> — the slots where you drag your data.</p>
  </div>
  <div class="pane-box">
    <div class="pane-icon">📋</div>
    <strong>Fields Pane</strong>
    <p>Far right. Lists all your tables and columns. Drag items from here into the field wells. Measures show with a calculator icon (&#x2211;).</p>
  </div>
  <div class="pane-box">
    <div class="pane-icon">🎨</div>
    <strong>Format Pane</strong>
    <p>Appears after clicking a visual. Use it to change colors, titles, font sizes, and borders. It is the paint roller icon above the field wells.</p>
  </div>
</div>

<p><strong>The basic workflow for every visual is always:</strong></p>
<ol class="steps">
  <li>Click the visual icon in the Visualizations pane &rarr; an empty box appears on the canvas</li>
  <li>Drag fields from the Fields pane into the correct field wells (Axis, Values, Legend, etc.)</li>
  <li>Click the paint roller icon to open the Format pane and make it look good</li>
  <li>Resize by dragging the edges of the visual. Move by dragging the title bar.</li>
</ol>

<div class="tip">Rename each page tab at the bottom: right-click the tab &rarr; Rename Page. Use names like "Headcount", "Compensation", "Attrition", "Performance", "Data Quality".</div>

<!-- ═══════════════════════════════════ PAGE 1 BUILD -->
<h2 style="margin-top:48px">Page 1 &mdash; Headcount Overview</h2>
<span class="page-badge">DASHBOARD PAGE 1</span>
<p>Your landing page. Build these visuals in order — cards first, then charts.</p>

<div class="build-visual">
  <div class="viztype-badge">Card Visual</div>
  <h4>Visual 1, 2, 3 &mdash; Three KPI Cards (Active Headcount / New Hires / Attrition Rate)</h4>
  <p>You will make three cards and place them in a row across the top. Repeat these steps 3 times:</p>
  <ol class="steps">
    <li>In the Visualizations pane, click the <strong>Card</strong> icon — it looks like a rectangle with the number "123" inside it</li>
    <li>An empty card appears on the canvas. Leave it for now.</li>
    <li>From the Fields pane, drag the measure into the <strong>Fields</strong> well:</li>
  </ol>
  <table class="field-table">
    <tr><th>Card</th><th>Drag this measure into Fields well</th></tr>
    <tr><td>Card 1</td><td>Active Headcount</td></tr>
    <tr><td>Card 2</td><td>New Hires Last 12 Months</td></tr>
    <tr><td>Card 3</td><td>Attrition Rate %</td></tr>
  </table>
  <p><strong>Format each card (paint roller icon):</strong></p>
  <ul class="format-list">
    <li><strong>General &rarr; Title &rarr; turn On</strong> — type the card label (e.g. "Active Employees")</li>
    <li><strong>Callout value &rarr; Font</strong> — set to size 36–40 so the number is big and readable</li>
    <li><strong>General &rarr; Effects &rarr; Background</strong> — pick a very light color (light blue or white) to make cards stand out from the canvas</li>
    <li>Drag the three cards to sit in a row along the top of the page. Make them the same size by dragging the edges.</li>
  </ul>
</div>

<div class="build-visual">
  <div class="viztype-badge">Clustered Bar Chart</div>
  <h4>Visual 4 &mdash; Headcount by Department</h4>
  <ol class="steps">
    <li>Click the <strong>Clustered bar chart</strong> icon (horizontal bars stacked side by side — it is usually in the top row of the Visualizations pane)</li>
    <li>Drag fields into the wells:</li>
  </ol>
  <table class="field-table">
    <tr><th>Field Well</th><th>What to drag in</th><th>From table</th></tr>
    <tr><td><span class="well-name">Y-axis</span></td><td>Department</td><td>Employees</td></tr>
    <tr><td><span class="well-name">X-axis</span></td><td>Active Headcount (measure)</td><td>Employees</td></tr>
  </table>
  <ul class="format-list">
    <li><strong>Bars &rarr; Color</strong> — pick a single color (dark blue works well)</li>
    <li><strong>Data labels &rarr; turn On</strong> — shows the count on each bar</li>
    <li><strong>General &rarr; Title</strong> — type "Headcount by Department"</li>
    <li>Sort bars: click the three dots (...) on the visual &rarr; Sort axis &rarr; Active Headcount &rarr; Sort descending (biggest department on top)</li>
  </ul>
</div>

<div class="build-visual">
  <div class="viztype-badge">Line Chart</div>
  <h4>Visual 5 &mdash; Monthly Hires Over Time</h4>
  <ol class="steps">
    <li>Click the <strong>Line chart</strong> icon (a line going up and to the right)</li>
    <li>Drag fields into the wells:</li>
  </ol>
  <table class="field-table">
    <tr><th>Field Well</th><th>What to drag in</th><th>From table</th></tr>
    <tr><td><span class="well-name">X-axis</span></td><td>YearMonth</td><td>DateTable</td></tr>
    <tr><td><span class="well-name">Y-axis</span></td><td>Monthly New Hires (measure)</td><td>Employees</td></tr>
  </table>
  <ul class="format-list">
    <li><strong>Lines &rarr; Color</strong> — pick a contrasting color to your bar chart</li>
    <li><strong>Markers &rarr; turn On</strong> — adds dots at each data point</li>
    <li><strong>General &rarr; Title</strong> — type "Monthly New Hires"</li>
  </ul>
  <div class="info" style="margin-top:10px">If the X-axis shows a date hierarchy instead of YearMonth, click the small arrow next to YearMonth in the well and select "YearMonth" (not the auto-hierarchy).</div>
</div>

<div class="build-visual">
  <div class="viztype-badge">Donut Chart</div>
  <h4>Visual 6 &mdash; Employment Type Breakdown</h4>
  <ol class="steps">
    <li>Click the <strong>Donut chart</strong> icon (a ring/circle with a hole in the middle)</li>
    <li>Drag fields into the wells:</li>
  </ol>
  <table class="field-table">
    <tr><th>Field Well</th><th>What to drag in</th><th>From table</th></tr>
    <tr><td><span class="well-name">Legend</span></td><td>EmploymentType</td><td>Employees</td></tr>
    <tr><td><span class="well-name">Values</span></td><td>Active Headcount (measure)</td><td>Employees</td></tr>
  </table>
  <ul class="format-list">
    <li><strong>Detail labels &rarr; turn On, show Percent</strong></li>
    <li><strong>General &rarr; Title</strong> — type "By Employment Type"</li>
  </ul>
</div>

<div class="build-visual">
  <div class="viztype-badge">Slicer</div>
  <h4>Visual 7 &mdash; Department Slicer (Filter)</h4>
  <p>A slicer is a clickable filter. When you click "Engineering" in the slicer, every other visual on the page updates to show only Engineering data.</p>
  <ol class="steps">
    <li>Click the <strong>Slicer</strong> icon (looks like a funnel)</li>
    <li>Drag <strong>Department</strong> from Employees into the <strong>Field</strong> well</li>
    <li>Format &rarr; Slicer settings &rarr; Style &rarr; change to <strong>Tile</strong> (makes it look like clickable buttons instead of a list)</li>
    <li>Place it somewhere easy to reach — the top right or left side of the page</li>
  </ol>
</div>

<!-- ═══════════════════════════════════ PAGE 2 BUILD -->
<h2 id="page2">Page 2 &mdash; Compensation Analysis</h2>
<span class="page-badge">DASHBOARD PAGE 2</span>

<div class="warn">Before adding any visuals, set a page-level filter so all compensation charts automatically show current salaries only: open the <strong>Filters</strong> pane (right side) &rarr; drag <strong>IsCurrentSalary</strong> from Compensation to the <strong>Filters on this page</strong> section &rarr; set value to <strong>TRUE</strong>. This applies to every visual on this page without you having to set it on each one individually.</div>

<div class="build-visual">
  <div class="viztype-badge">Card Visual</div>
  <h4>Visuals 1, 2, 3 &mdash; Three Salary KPI Cards</h4>
  <p>Same process as Page 1. Make three cards in a top row:</p>
  <table class="field-table">
    <tr><th>Card</th><th>Measure</th><th>Format as</th></tr>
    <tr><td>Card 1</td><td>Avg Current Salary</td><td>Currency $ — go to Measure Tools &rarr; Format &rarr; Currency</td></tr>
    <tr><td>Card 2</td><td>Median Current Salary</td><td>Currency $</td></tr>
    <tr><td>Card 3</td><td>Total Payroll</td><td>Currency $ — in the Format pane set Display units to Millions</td></tr>
  </table>
</div>

<div class="build-visual">
  <div class="viztype-badge">Clustered Bar Chart</div>
  <h4>Visual 4 &mdash; Average Salary by Department</h4>
  <table class="field-table">
    <tr><th>Field Well</th><th>What to drag in</th><th>From table</th></tr>
    <tr><td><span class="well-name">Y-axis</span></td><td>Department</td><td>Employees</td></tr>
    <tr><td><span class="well-name">X-axis</span></td><td>Avg Current Salary (measure)</td><td>Employees</td></tr>
  </table>
  <ul class="format-list">
    <li><strong>Data labels &rarr; On</strong> &mdash; shows the dollar amount on each bar</li>
    <li><strong>X-axis &rarr; Display units</strong> &mdash; set to Thousands so numbers read as "$95K" instead of "$95,000"</li>
    <li>Sort descending by salary (highest-paying dept on top)</li>
  </ul>
</div>

<div class="build-visual">
  <div class="viztype-badge">Clustered Column Chart</div>
  <h4>Visual 5 &mdash; Average Salary by Pay Grade</h4>
  <ol class="steps">
    <li>Click the <strong>Clustered column chart</strong> icon (vertical bars — different from bar chart which is horizontal)</li>
  </ol>
  <table class="field-table">
    <tr><th>Field Well</th><th>What to drag in</th><th>From table</th></tr>
    <tr><td><span class="well-name">X-axis</span></td><td>PayGrade</td><td>Compensation</td></tr>
    <tr><td><span class="well-name">Y-axis</span></td><td>Avg Current Salary (measure)</td><td>Employees</td></tr>
  </table>
  <div class="tip" style="margin-top:10px">Pay grades go G1, G2, G3, G4, M1, M2, D1, V1. The chart should show a clear staircase pattern — each grade earns more than the last. This is a great visual proof that your salary logic worked.</div>
</div>

<div class="build-visual">
  <div class="viztype-badge">Table Visual</div>
  <h4>Visual 6 &mdash; Salary Outliers Table</h4>
  <ol class="steps">
    <li>Click the <strong>Table</strong> icon (grid of rows and columns)</li>
    <li>Drag these columns into the <strong>Columns</strong> well one at a time:</li>
  </ol>
  <table class="field-table">
    <tr><th>Columns well (add in this order)</th><th>From table</th></tr>
    <tr><td>FullName</td><td>Employees</td></tr>
    <tr><td>Department</td><td>Employees</td></tr>
    <tr><td>BaseSalary</td><td>Compensation</td></tr>
    <tr><td>PayGrade</td><td>Compensation</td></tr>
    <tr><td>SalaryOutlier</td><td>Compensation</td></tr>
  </table>
  <ol class="steps" style="counter-reset: step 3;">
    <li>In the <strong>Filters</strong> pane, drag <strong>SalaryOutlier</strong> to <strong>Filters on this visual</strong> &rarr; set to <strong>TRUE</strong>. Now only outlier rows show.</li>
    <li>Format &rarr; Style presets &rarr; pick <strong>Minimal</strong> for a clean look</li>
    <li>Title: "Salary Records Needing Review"</li>
  </ol>
</div>

<!-- ═══════════════════════════════════ PAGE 3 BUILD -->
<h2 id="page3">Page 3 &mdash; Attrition &amp; Retention</h2>
<span class="page-badge">DASHBOARD PAGE 3</span>

<div class="build-visual">
  <div class="viztype-badge">Card Visual</div>
  <h4>Visuals 1, 2, 3 &mdash; Three Attrition KPI Cards</h4>
  <table class="field-table">
    <tr><th>Card</th><th>Measure</th><th>Format as</th></tr>
    <tr><td>Card 1</td><td>Total Terminated</td><td>Whole number</td></tr>
    <tr><td>Card 2</td><td>Voluntary Rate %</td><td>Measure Tools &rarr; Format &rarr; Percentage</td></tr>
    <tr><td>Card 3</td><td>Avg Tenure at Exit Years</td><td>Decimal number, 1 place — title it "Avg Tenure at Exit (Yrs)"</td></tr>
  </table>
</div>

<div class="build-visual">
  <div class="viztype-badge">Donut Chart</div>
  <h4>Visual 4 &mdash; Termination Reason Breakdown</h4>
  <table class="field-table">
    <tr><th>Field Well</th><th>What to drag in</th><th>From table</th></tr>
    <tr><td><span class="well-name">Legend</span></td><td>TerminationReason</td><td>Termination</td></tr>
    <tr><td><span class="well-name">Values</span></td><td>Total Terminated (measure)</td><td></td></tr>
  </table>
  <ul class="format-list">
    <li>Detail labels &rarr; On &rarr; show both Value and Percent</li>
    <li>Title: "Why Employees Left"</li>
  </ul>
</div>

<div class="build-visual">
  <div class="viztype-badge">Clustered Bar Chart</div>
  <h4>Visual 5 &mdash; Terminations by Department</h4>
  <table class="field-table">
    <tr><th>Field Well</th><th>What to drag in</th><th>From table</th></tr>
    <tr><td><span class="well-name">Y-axis</span></td><td>Department</td><td>Employees</td></tr>
    <tr><td><span class="well-name">X-axis</span></td><td>Total Terminated (measure)</td><td></td></tr>
  </table>
</div>

<div class="build-visual">
  <div class="viztype-badge">Line Chart</div>
  <h4>Visual 6 &mdash; Monthly Terminations Over Time</h4>
  <table class="field-table">
    <tr><th>Field Well</th><th>What to drag in</th><th>From table</th></tr>
    <tr><td><span class="well-name">X-axis</span></td><td>YearMonth</td><td>DateTable</td></tr>
    <tr><td><span class="well-name">Y-axis</span></td><td>Monthly Terminations (measure)</td><td></td></tr>
  </table>
  <div class="info" style="margin-top:10px">This measure uses USERELATIONSHIP to activate the inactive link between DateTable and Termination[TerminationDate], so it will show terminations plotted over time correctly.</div>
</div>

<div class="build-visual">
  <div class="viztype-badge">Slicer</div>
  <h4>Visual 7 &mdash; Termination Reason Slicer</h4>
  <ol class="steps">
    <li>Click the Slicer icon, drag <strong>TerminationReason</strong> from Termination into the <strong>Field</strong> well</li>
    <li>Format &rarr; Slicer settings &rarr; Style &rarr; <strong>Tile</strong></li>
    <li>Now clicking "Voluntary" in the slicer will filter all charts on this page to show only voluntary leavers</li>
  </ol>
</div>

<!-- ═══════════════════════════════════ PAGE 4 BUILD -->
<h2 id="page4">Page 4 &mdash; Performance &amp; Talent</h2>
<span class="page-badge">DASHBOARD PAGE 4</span>

<div class="build-visual">
  <div class="viztype-badge">Card Visual</div>
  <h4>Visuals 1, 2, 3 &mdash; Three Performance KPI Cards</h4>
  <table class="field-table">
    <tr><th>Card</th><th>Measure</th><th>Format as</th></tr>
    <tr><td>Card 1</td><td>Avg Performance Rating</td><td>Decimal, 1 place</td></tr>
    <tr><td>Card 2</td><td>High Performer Count</td><td>Whole number — title "High Performers (4-5)"</td></tr>
    <tr><td>Card 3</td><td>Promotion Eligible %</td><td>Percentage, 1 decimal</td></tr>
  </table>
</div>

<div class="build-visual">
  <div class="viztype-badge">Clustered Column Chart</div>
  <h4>Visual 4 &mdash; Rating Distribution</h4>
  <p>This shows how many employees got each rating (1 through 5). You want to see a bell curve shape centered around 3.</p>
  <table class="field-table">
    <tr><th>Field Well</th><th>What to drag in</th><th>From table</th></tr>
    <tr><td><span class="well-name">X-axis</span></td><td>RatingLabel</td><td>Performance</td></tr>
    <tr><td><span class="well-name">Y-axis</span></td><td>Count of EmployeeID</td><td>Performance — drag EmployeeID and Power BI will count it automatically</td></tr>
  </table>
  <ul class="format-list">
    <li>If the X-axis sorts alphabetically (1 before 5 but "2 - Below" before "1 - Needs"), click the three dots (...) &rarr; Sort axis &rarr; RatingLabel &rarr; ascending</li>
    <li>Data labels &rarr; On</li>
    <li>Title: "Rating Distribution"</li>
  </ul>
</div>

<div class="build-visual">
  <div class="viztype-badge">Clustered Bar Chart</div>
  <h4>Visual 5 &mdash; Average Rating by Department</h4>
  <table class="field-table">
    <tr><th>Field Well</th><th>What to drag in</th><th>From table</th></tr>
    <tr><td><span class="well-name">Y-axis</span></td><td>Department</td><td>Employees</td></tr>
    <tr><td><span class="well-name">X-axis</span></td><td>Avg Performance Rating (measure)</td><td></td></tr>
  </table>
  <ul class="format-list">
    <li>X-axis &rarr; set minimum to 1, maximum to 5 (so the scale is always anchored to the rating range)</li>
  </ul>
</div>

<div class="build-visual">
  <div class="viztype-badge">Table Visual</div>
  <h4>Visual 6 &mdash; Promotion Pipeline Table</h4>
  <p>Lists every high performer who is also marked eligible for promotion — your talent pipeline.</p>
  <ol class="steps">
    <li>Click the Table icon, drag these columns into the <strong>Columns</strong> well:</li>
  </ol>
  <table class="field-table">
    <tr><th>Column</th><th>From table</th></tr>
    <tr><td>FullName</td><td>Employees</td></tr>
    <tr><td>Department</td><td>Employees</td></tr>
    <tr><td>JobTitle</td><td>Employees</td></tr>
    <tr><td>Rating</td><td>Performance</td></tr>
    <tr><td>PromotionEligible</td><td>Performance</td></tr>
  </table>
  <ol class="steps" style="counter-reset: step 2;">
    <li>Filters pane &rarr; Filters on this visual &rarr; drag <strong>Rating</strong> &rarr; set filter type to <strong>is greater than or equal to</strong> &rarr; type <strong>4</strong></li>
    <li>Also filter <strong>PromotionEligible</strong> &rarr; is &rarr; <strong>Yes</strong></li>
    <li>Title: "Promotion Pipeline"</li>
  </ol>
</div>

<!-- ═══════════════════════════════════ PAGE 5 BUILD -->
<h2 id="page5">Page 5 &mdash; Data Quality Alerts</h2>
<span class="page-badge">DASHBOARD PAGE 5</span>
<p>This page highlights the intentional data problems from the raw dataset. In a real company, HR Ops would use this page to find and fix issues before they affect reporting.</p>

<div class="build-visual">
  <div class="viztype-badge">Card Visual</div>
  <h4>Visuals 1–4 &mdash; Four Data Quality Score Cards</h4>
  <p>Make four cards in a row. For the first three, a high number is <em>bad</em>. For the last one, a high number is <em>good</em>.</p>
  <table class="field-table">
    <tr><th>Card</th><th>Measure</th><th>Good sign</th></tr>
    <tr><td>Card 1</td><td>Missing Manager Count</td><td>Should be 0</td></tr>
    <tr><td>Card 2</td><td>Salary Outlier Count</td><td>Should be 0</td></tr>
    <tr><td>Card 3</td><td>Overlapping Leave Count</td><td>Should be 0</td></tr>
    <tr><td>Card 4</td><td>Data Quality Score %</td><td>Should be close to 100%</td></tr>
  </table>
  <p><strong>Color the alert cards red/orange:</strong> Select Card 1 &rarr; Format pane &rarr; Callout value &rarr; Font color &rarr; pick red. Do the same for Cards 2 and 3. Leave Card 4 green.</p>
</div>

<div class="build-visual">
  <div class="viztype-badge">Table Visual</div>
  <h4>Visual 5 &mdash; Employees Missing a Manager</h4>
  <ol class="steps">
    <li>Click the Table icon, add these columns to the Columns well: <strong>EmployeeID, FullName, Department, JobTitle</strong> — all from Employees</li>
    <li>Filters pane &rarr; Filters on this visual &rarr; drag <strong>ManagerID_Missing</strong> &rarr; set to <strong>TRUE</strong></li>
    <li>Title: "Employees with No Manager Assigned"</li>
  </ol>
</div>

<div class="build-visual">
  <div class="viztype-badge">Table Visual</div>
  <h4>Visual 6 &mdash; Salary Outlier Records</h4>
  <ol class="steps">
    <li>Click the Table icon, add: <strong>FullName</strong> (Employees), <strong>Department</strong> (Employees), <strong>BaseSalary, PayGrade, SalaryOutlier</strong> (Compensation)</li>
    <li>Filters pane &rarr; Filters on this visual &rarr; drag <strong>SalaryOutlier</strong> &rarr; set to <strong>TRUE</strong></li>
    <li>Title: "Salary Records Needing Review"</li>
  </ol>
</div>

<div class="build-visual">
  <div class="viztype-badge">Table Visual</div>
  <h4>Visual 7 &mdash; Overlapping Leave Records</h4>
  <ol class="steps">
    <li>Click the Table icon, add: <strong>EmployeeID, LeaveType, StartDate, EndDate, DurationDays, OverlapWithPriorLeave</strong> — all from Leave</li>
    <li>Filters pane &rarr; Filters on this visual &rarr; drag <strong>OverlapWithPriorLeave</strong> &rarr; set to <strong>TRUE</strong></li>
    <li>Title: "Overlapping Leave Windows"</li>
  </ol>
</div>

<!-- ═══════════════════════════════════ ALL MEASURES -->
<h2 id="all-measures">Quick Reference — All Measures</h2>

<table class="measure-table">
  <tr><th>Measure Name</th><th>Used On Page</th><th>Format</th></tr>
  <tr><td>Active Headcount</td><td>1, 3</td><td>Whole number</td></tr>
  <tr><td>Total Employees</td><td>1</td><td>Whole number</td></tr>
  <tr><td>New Hires Last 12 Months</td><td>1</td><td>Whole number</td></tr>
  <tr><td>Monthly New Hires</td><td>1</td><td>Whole number</td></tr>
  <tr><td>Attrition Rate %</td><td>1, 3</td><td>Percentage, 1 decimal</td></tr>
  <tr><td>Avg Current Salary</td><td>2</td><td>Currency $, 0 decimals</td></tr>
  <tr><td>Median Current Salary</td><td>2</td><td>Currency $, 0 decimals</td></tr>
  <tr><td>Total Payroll</td><td>2</td><td>Currency $M, 1 decimal</td></tr>
  <tr><td>Avg Bonus Target %</td><td>2</td><td>Whole number (it's already a %)</td></tr>
  <tr><td>Salary Outlier Count</td><td>2, 5</td><td>Whole number</td></tr>
  <tr><td>Total Terminated</td><td>3</td><td>Whole number</td></tr>
  <tr><td>Voluntary Terminations</td><td>3</td><td>Whole number</td></tr>
  <tr><td>Involuntary Terminations</td><td>3</td><td>Whole number</td></tr>
  <tr><td>Voluntary Rate %</td><td>3</td><td>Percentage, 1 decimal</td></tr>
  <tr><td>Avg Tenure at Exit Years</td><td>3</td><td>Decimal, 1 place</td></tr>
  <tr><td>Monthly Terminations</td><td>3</td><td>Whole number</td></tr>
  <tr><td>Avg Performance Rating</td><td>4</td><td>Decimal, 1 place</td></tr>
  <tr><td>High Performer Count</td><td>4</td><td>Whole number</td></tr>
  <tr><td>Promotion Eligible %</td><td>4</td><td>Percentage, 1 decimal</td></tr>
  <tr><td>Promotion Ready</td><td>4</td><td>Whole number</td></tr>
  <tr><td>Missing Manager Count</td><td>5</td><td>Whole number</td></tr>
  <tr><td>Overlapping Leave Count</td><td>5</td><td>Whole number</td></tr>
  <tr><td>Data Quality Score %</td><td>5</td><td>Percentage, 1 decimal</td></tr>
</table>

<!-- ═══════════════════════════════════ TIPS -->
<h2 id="tips">Final Tips Before You Start</h2>

<div class="tip">Save your .pbix file early and often. Power BI does not auto-save.</div>

<div class="tip">To format a measure as currency: click the measure in the Fields pane &rarr; Measure Tools tab &rarr; Format dropdown &rarr; Currency. Set symbol to $ and decimals to 0.</div>

<div class="tip">To add a page-level filter (e.g., IsCurrentSalary = TRUE on Page 2): open the Filters pane &rarr; drag the column to "Filters on this page" &rarr; set it to TRUE. This applies to every visual on the page automatically.</div>

<div class="tip">If a slicer shows blank values, add a filter on that slicer visual itself: Filters pane &rarr; Visual level filter &rarr; [column] is not blank.</div>

<div class="warn">Don't use SUMX or AVERAGEX unless you're confident — start with SUM and AVERAGE wrapped in CALCULATE. They're simpler and less likely to produce surprises.</div>

<div class="info">When your dashboard is done, use File &rarr; Export &rarr; Export to PDF to create a shareable version without sharing the .pbix file (which contains all your data).</div>

<br><br>
<p style="text-align:center; color:#999; font-size:0.8rem;">Generated by generate_pbi_guide.py &mdash; HRIS Portfolio Project</p>
<br>

</main>

<script>
function copyCode(btn) {
  const pre = btn.closest('.dax-block').querySelector('pre');
  const text = pre.innerText;
  navigator.clipboard.writeText(text).then(() => {
    btn.textContent = 'Copied!';
    setTimeout(() => btn.textContent = 'Copy', 1800);
  });
}

// Highlight active nav link on scroll
const sections = document.querySelectorAll('h2[id]');
const links    = document.querySelectorAll('.sidebar a');
window.addEventListener('scroll', () => {
  let current = '';
  sections.forEach(s => { if (window.scrollY >= s.offsetTop - 80) current = s.id; });
  links.forEach(l => {
    l.classList.toggle('active', l.getAttribute('href') === '#' + current);
  });
});
</script>

</body>
</html>
"""

output_path = "PowerBI_Build_Guide.html"
with open(output_path, "w", encoding="utf-8") as f:
    f.write(GUIDE_HTML)

print(f"Guide written to: {output_path}")
print("Open it in any browser (double-click the file) and keep it")
print("open alongside Power BI Desktop while you build.")
