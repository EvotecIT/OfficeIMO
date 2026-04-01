---
title: "Automating Weekly Reports with OfficeIMO.Excel and GitHub Actions"
description: "A practical sample showing one way to build automated Excel reports in CI/CD with OfficeIMO.Excel and GitHub Actions."
date: 2026-01-15
tags: [excel, ci-cd, automation]
categories: [Workflow]
author: "Przemyslaw Klys"
---

This walkthrough uses GitHub issues as a concrete sample input, but the same pattern works for any repeatable report source your team already has: API responses, database exports, CSV snapshots, or repository metadata. The goal is not a special OfficeIMO-only workflow; it is a simple CI job that turns structured data into an `.xlsx` artifact on a schedule.

## Architecture

The workflow has three stages:

1. **Fetch data** from the GitHub API using the `gh` CLI.
2. **Generate the Excel report** with a small C# console app that uses OfficeIMO.Excel.
3. **Upload the artifact** so it can be downloaded or forwarded.

## The C# Report Generator

Create a console project:

```bash
dotnet new console -n WeeklyReport
cd WeeklyReport
dotnet add package OfficeIMO.Excel
dotnet add package System.Text.Json
```

The generator reads a JSON file produced by the workflow and writes an Excel workbook. The `Issue` record below matches the shape selected by the sample `gh issue list --json ...` call, so adjust it if you fetch different fields:

```csharp
using System.Text.Json;
using OfficeIMO.Excel;

var issues = JsonSerializer.Deserialize<List<Issue>>(
    File.ReadAllText("issues.json"));

using var workbook = ExcelDocument.Create("WeeklyReport.xlsx");
var sheet = workbook.AddSheet("Open Issues");

// Header row
sheet.SetRow(0, new object[] { "Number", "Title", "Labels", "Created", "Age (days)" });

int row = 1;
foreach (var issue in issues!)
{
    var age = (DateTime.UtcNow - issue.CreatedAt).Days;
    sheet.SetRow(row++, new object[]
    {
        issue.Number,
        issue.Title,
        string.Join(", ", issue.Labels),
        issue.CreatedAt.ToString("yyyy-MM-dd"),
        age
    });
}

sheet.AutoFitColumns();
workbook.Save();
Console.WriteLine($"Report generated with {issues.Count} issues.");

record Issue(int Number, string Title, List<string> Labels, DateTime CreatedAt);
```

## The GitHub Actions Workflow

```yaml
name: Weekly Report
on:
  schedule:
    - cron: "0 7 * * 1"  # Every Monday at 07:00 UTC
  workflow_dispatch:       # Allow manual trigger

permissions:
  contents: read
  issues: read

jobs:
  generate:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4

      - name: Setup .NET
        uses: actions/setup-dotnet@v4
        with:
          dotnet-version: "8.0.x"

      - name: Fetch open issues
        env:
          GH_TOKEN: ${{ secrets.GITHUB_TOKEN }}
        run: |
          gh issue list \
            --repo ${{ github.repository }} \
            --state open \
            --limit 500 \
            --json number,title,labels,createdAt \
            > WeeklyReport/issues.json

      - name: Generate Excel report
        run: dotnet run --project WeeklyReport

      - name: Upload report artifact
        uses: actions/upload-artifact@v4
        with:
          name: weekly-report-${{ github.run_number }}
          path: WeeklyReport/WeeklyReport.xlsx
          retention-days: 30
```

After each run, the Excel file is available as a downloadable artifact on the workflow run page. If you already have a downstream delivery step, you can hand the workbook off there instead of storing it as a long-lived artifact.

## Sending the Report via Email

If you need email delivery, add a mail action as a separate step:

```yaml
      - name: Send report
        uses: dawidd6/action-send-mail@v3
        with:
          server_address: smtp.office365.com
          server_port: 587
          username: ${{ secrets.MAIL_USER }}
          password: ${{ secrets.MAIL_PASS }}
          subject: "Weekly Issue Report - ${{ github.run_number }}"
          to: team@contoso.com
          from: reports@contoso.com
          body: "Attached is this week's open-issue report."
          attachments: WeeklyReport/WeeklyReport.xlsx
```

## Extending the Report

Once the skeleton is working, you can add more sheets:

- **PR Merge Times**: fetch `gh pr list --state merged --json mergedAt,createdAt` and compute duration.
- **Label Distribution**: pivot issues by label into a summary table.
- **Trend Sheet**: append a row to a persistent CSV in the repo and chart it over time.

Each sheet is just another call to `workbook.AddSheet()` with its own data. In a real pipeline you will usually move the report-building code into its own project, keep the schema classes under source control, and version the workbook layout like any other deliverable.

## Operational Note

Runner availability, retention limits, and billing depend on your GitHub plan and repository type. Check the current GitHub Actions pricing and usage docs for the exact limits that apply to your environment.

Automating repeatable reports like this takes the manual export step out of the loop while keeping the output in a format teams already know how to consume. OfficeIMO.Excel and GitHub Actions are one straightforward way to build that pipeline with ordinary .NET code.

## Continue with

- [OfficeIMO.Excel](/products/excel/) for the package overview and reporting-focused capabilities.
- [Excel documentation](/docs/excel/) for workbook structure, tables, formulas, and formatting patterns.
- [Building Excel Reports with Parallel Compute](/blog/excel-parallel-reports/) if your scheduled reports are large enough to benefit from parallel writes.
- [PSWriteOffice](/products/pswriteoffice/) if you want a PowerShell-driven automation path instead of a C# console app.
