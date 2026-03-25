---
title: "Automating Weekly Reports with OfficeIMO.Excel and GitHub Actions"
description: "A real-world guide to building automated Excel reports in CI/CD using OfficeIMO.Excel and GitHub Actions, complete with YAML workflow and C# code."
date: 2026-01-15
tags: [excel, ci-cd, automation]
categories: [Workflow]
author: "Przemyslaw Klys"
---

Every Monday morning, your team lead asks for the same report: open issues by label, pull request merge times, and a trend chart. Instead of spending 30 minutes clicking through dashboards, let GitHub Actions build the report automatically and drop it in a Teams channel or email inbox.

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

The generator reads a JSON file produced by the workflow and writes an Excel workbook:

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

After each run, the Excel file is available as a downloadable artifact on the workflow run page.

## Sending the Report via Email

Add a step using a mail action:

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

Each sheet is just another call to `workbook.AddSheet()` with its own data.

## Cost

GitHub Actions provides 2,000 free minutes per month for private repositories and unlimited minutes for public repositories. A single report run takes about 30 seconds, so even daily generation stays well within the free tier.

Automating the boring reports frees your team to focus on the work that matters. OfficeIMO.Excel and GitHub Actions make it surprisingly simple.
