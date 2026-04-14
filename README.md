# Excel Analysis and Comparison Web App

This project is an ASP.NET Core MVC web app that lets users upload an Excel file, choose an analysis goal, and get a clear outcome dashboard.

## What the app does

- Upload a `.xlsx` or `.xls` workbook.
- Read the first worksheet and detect headers automatically.
- Let the user choose one analysis goal:
  - `Overview` (auto insights and charts),
  - `Compare two groups` (A vs B values),
  - `Breakdown by category`,
  - `Trend over time`.
- Build a results dashboard with:
  - KPI cards,
  - charts,
  - column profile,
  - preview rows,
  - comparison table and actionable narrative.
- Export processed data and column profile to Excel.

## Main user flow

1. Upload an Excel file on the home page.
2. Select the analysis setup (goal + relevant columns).
3. Run analysis.
4. Review results and download export if needed.

## Tech stack

- .NET 8
- ASP.NET Core MVC
- EPPlus (Excel parsing/export)
- Bootstrap + Chart.js

## Run locally

```bash
dotnet restore
dotnet run
```

Then open the local URL shown in the terminal.
