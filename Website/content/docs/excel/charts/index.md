---
title: "Add Charts to Excel Workbooks"
description: Add native, editable Excel charts from objects, cell ranges, or tables, and style titles, axes, legends, and data labels with OfficeIMO.Excel.
order: 26
packages:
  - OfficeIMO.Excel
  - OfficeIMO.Excel.Pdf
limits: "Native chart fidelity is listed on the [Excel format card](/compatibility/#excel)."
search_terms:
  - excel chart
  - xlsx chart
  - line chart
  - pie chart
verified: 2026-11-14
---

OfficeIMO.Excel writes native Excel charts. They stay editable in Excel and keep a link to their source cells, so they update when the data changes. The chart type comes from the shared `OfficeChartType` enum (`Column`, `ColumnClustered`, `Bar`, `Line`, `Pie`, `Area`, `Scatter`, and more).

## Chart a list of objects

Pass the objects, the category property, and one or more value properties. The data is written to the sheet and the chart references it.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Excel;

Sale[] sales = [new("Q1", 10_500, 12_000), new("Q2", 19_800, 18_000), new("Q3", 25_300, 28_000)];

using var book = ExcelDocument.Create("Sales.xlsx");
ExcelSheet sheet = book.AddSheet("Sales");
sheet.AddTable(sales, new OfficeTableOptions { Name = "Sales" });

sheet.AddChart(OfficeChartType.ColumnClustered, sales, nameof(Sale.Quarter),
    new ExcelChartOptions { Title = "Revenue vs target", Row = 2, Column = 6 },
    nameof(Sale.Revenue), nameof(Sale.Target));

book.Save();

record Sale(string Quarter, double Revenue, double Target);
```

### PowerShell

```powershell
$trend = Import-OfficeCsv -Path .\Data\Ticket-Trend.csv
New-OfficeExcel -Path .\Reports\Helpdesk-Trend.xlsx {
    ExcelSheet 'Trend' {
        ExcelChart -InputObject $trend -Category Week -Value Opened, Closed -Type Line -Address A1 -Title 'Tickets per week'
    }
}
```

## Chart explicit series and style it

Use `ExcelChartData` when the values do not come from a list of objects. The returned `ExcelChart` has fluent methods for the legend, axis titles, number formats, and data labels. Combination charts (series on primary and secondary axes), trendlines, data-label templates, and scatter charts are shown in `OfficeIMO.Examples/Excel/ChartsExcel.cs` in the repository.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;

using var book = ExcelDocument.Create("Quarterly.xlsx");
ExcelSheet sheet = book.AddSheet("Summary");

var data = new ExcelChartData(
    new[] { "Q1", "Q2", "Q3", "Q4" },
    new[] {
        new ExcelChartSeries("Sales", new[] { 10d, 20d, 25d, 30d }),
        new ExcelChartSeries("Target", new[] { 12d, 22d, 24d, 32d })
    });

ExcelChart chart = sheet.AddChart(OfficeChartType.ColumnClustered, data, new ExcelChartOptions {
    Row = 2, Column = 2,
    Width = OfficeLength.Pixels(640), Height = OfficeLength.Pixels(360),
    Title = "Quarterly sales"
});
chart.SetLegend(OfficeChartLegendPosition.Bottom)
     .SetCategoryAxisTitle("Quarter")
     .SetValueAxisTitle("Revenue (k)")
     .SetValueAxisNumberFormat("#,##0", sourceLinked: false);

book.Save();
```

## Chart an existing table

When the data is already an Excel table, reference it by name.

### PowerShell

```powershell
New-OfficeExcel -Path .\Reports\Regions.xlsx {
    ExcelSheet 'Summary' {
        ExcelTable (Import-OfficeCsv -Path .\Data\Region-Revenue.csv) -Name RegionSummary
        ExcelChart -Table RegionSummary -Category Region -Value Revenue -Address F2 -Type Column
    }
}
```
