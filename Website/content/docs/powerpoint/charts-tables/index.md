---
title: Add Charts and Tables to PowerPoint Slides
description: "Add native, editable PowerPoint tables and charts from .NET objects, merge cells, continue long tables across slides, and style chart axes and labels."
order: 32
packages:
  - OfficeIMO.PowerPoint
cmdlets:
  - Add-OfficePowerPointTable
  - Add-OfficePowerPointChart
  - Set-OfficePowerPointTableCell
limits: Long tables are not clipped or shrunk automatically; use AddTableSlides to continue rows on new slides.
search_terms:
  - powerpoint chart
  - pptx table
  - slide chart from data
  - editable chart
  - table across slides
verified: 2026-11-14
---

Tables and charts written by `OfficeIMO.PowerPoint` are native PowerPoint objects: tables stay editable cell by cell, and charts carry an embedded workbook, so **Edit Data** works in PowerPoint. Chart types come from the shared `OfficeChartType` enum (`ColumnClustered`, `BarStacked`, `Line`, `Pie`, `Doughnut`, `Scatter`, `Bubble`, `Radar`, and more), the same one Excel and Word use.

## Add a table from a list of objects

Each public property becomes a column. The callback controls header text and column order.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Data;
using OfficeIMO.PowerPoint;

ProductSales[] rows = [
    new("Alpha", 12, 15, 18, 20),
    new("Beta", 9, 11, 13, 14),
    new("Gamma", 6, 9, 12, 16)
];

using PowerPointDocument presentation = PowerPointDocument.Create("sales.pptx");
PowerPointSlide slide = presentation.AddSlide();
slide.AddTitle("Sales by product");

PowerPointTable table = slide.AddTable(rows,
    options => {
        options.HeaderCase = HeaderCase.Title;
        options.PinFirst(nameof(ProductSales.Product));
    },
    includeHeaders: true,
    left: OfficeLength.Centimeters(2), top: OfficeLength.Centimeters(4),
    width: OfficeLength.Centimeters(22), height: OfficeLength.Centimeters(5));
table.SetColumnWidthsEvenly();

presentation.Save();

record ProductSales(string Product, int Q1, int Q2, int Q3, int Q4);
```

### PowerShell

```powershell
$rows = Import-Csv .\product-sales.csv
New-OfficePowerPoint -Path .\sales.pptx {
    Add-OfficePowerPointSlide -Title 'Sales by product' {
        Add-OfficePowerPointTable -InputObject $rows -X 2cm -Y 4cm -Width 22cm -Height 5cm
    }
}
```

## Fill cells and merge them

`GetCell(row, column)` returns a cell you can write and format. `MergeCells` takes the first and last row and column of the range.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;

using PowerPointDocument presentation = PowerPointDocument.Create("merged.pptx");
PowerPointSlide slide = presentation.AddSlide();

PowerPointTable table = slide.AddTable(3, 4, OfficeLength.Centimeters(2), OfficeLength.Centimeters(3),
    OfficeLength.Centimeters(20), OfficeLength.Centimeters(4));
table.GetCell(0, 0).Text = "2026 sales";
table.MergeCells(0, 0, 0, 3);

string[] quarters = ["Q1", "Q2", "Q3", "Q4"];
for (int c = 0; c < quarters.Length; c++) {
    PowerPointTableCell header = table.GetCell(1, c);
    header.Text = quarters[c];
    header.Bold = true;
    header.HorizontalAlignment = PowerPointTextAlignment.Center;
    table.GetCell(2, c).Text = (100 + c * 20).ToString();
}

presentation.Save();
```

### PowerShell

`Set-OfficePowerPointTableCell` writes one cell; its row and column numbers start at 1.

```powershell
New-OfficePowerPoint -Path .\quarters.pptx {
    $slide = Add-OfficePowerPointSlide -Title '2026 sales' -PassThru
    $table = Add-OfficePowerPointTable -Slide $slide -Rows 2 -Columns 4 -X 2cm -Y 4cm -Width 20cm -Height 3cm -PassThru
    $quarters = 'Q1', 'Q2', 'Q3', 'Q4'
    for ($c = 1; $c -le 4; $c++) {
        $table | Set-OfficePowerPointTableCell -Row 1 -Column $c -Text $quarters[$c - 1]
    }
}
```

## Continue a long table on new slides

`AddTableSlides` adds as many slides as the rows need, measured from the table box and the minimum row height, and repeats the header row on each one. `ConfigureSlide` runs once per page so you can add a title.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;

Device[] devices = Enumerable.Range(1, 60)
    .Select(i => new Device($"PC-{i:000}", i % 3 == 0 ? "Pending" : "Compliant"))
    .ToArray();

using PowerPointDocument presentation = PowerPointDocument.Create("inventory.pptx");

PowerPointPaginatedTableResult result = presentation.AddTableSlides(devices,
    new[] {
        PowerPointTableColumn<Device>.Create("Device", d => d.Name),
        PowerPointTableColumn<Device>.Create("State", d => d.State)
    },
    new PowerPointTablePaginationOptions {
        MinimumRowHeight = OfficeLength.Points(20),
        ConfigureSlide = (slide, page) =>
            slide.AddTitle($"Device inventory ({page.PageIndex + 1} of {page.PageCount})")
    });

Console.WriteLine($"{result.SourceRowCount} rows on {result.PageCount} slides");
presentation.Save();

record Device(string Name, string State);
```

## Add a chart from data

`OfficeChartData` holds the categories and one `OfficeSeries` per data series. `PowerPointChartOptions` places the chart on the slide.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;

string[] products = ["Alpha", "Beta", "Gamma"];
var data = new OfficeChartData(products, new[] {
    new OfficeSeries("Q1", new double[] { 12, 9, 6 }),
    new OfficeSeries("Q2", new double[] { 15, 11, 9 }),
    new OfficeSeries("Q3", new double[] { 18, 13, 12 })
});

using PowerPointDocument presentation = PowerPointDocument.Create("chart.pptx");
PowerPointSlide slide = presentation.AddSlide();
slide.AddTitle("Quarterly performance");

slide.AddChart(OfficeChartType.ColumnClustered, data, new PowerPointChartOptions {
    Left = OfficeLength.Centimeters(2), Top = OfficeLength.Centimeters(4),
    Width = OfficeLength.Centimeters(22), Height = OfficeLength.Centimeters(11),
    Title = "Units sold"
});

presentation.Save();
```

### PowerShell

```powershell
$rows = @(
    [pscustomobject]@{ Month = 'Jan'; Sales = 42; Profit = 9 }
    [pscustomobject]@{ Month = 'Feb'; Sales = 55; Profit = 13 }
    [pscustomobject]@{ Month = 'Mar'; Sales = 61; Profit = 15 }
)
New-OfficePowerPoint -Path .\chart.pptx {
    $slide = Add-OfficePowerPointSlide -Title 'Monthly performance' -PassThru
    Add-OfficePowerPointChart -Slide $slide -InputObject $rows -Category Month -Value Sales, Profit -Type ColumnClustered -Title 'Sales and profit'
}
```

## Style a chart and update its data

The chart methods return the chart, so settings chain. `UpdateData` replaces the values in both the chart and its embedded workbook.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;

using PowerPointDocument presentation = PowerPointDocument.Create("trend.pptx");
PowerPointSlide slide = presentation.AddSlide();

var data = new OfficeChartData(new[] { "Jan", "Feb", "Mar" }, new[] {
    new OfficeSeries("Tickets", new double[] { 120, 98, 87 })
});

PowerPointChart chart = slide.AddChart(OfficeChartType.Line, data, new PowerPointChartOptions {
    Left = OfficeLength.Centimeters(2), Top = OfficeLength.Centimeters(2),
    Width = OfficeLength.Centimeters(22), Height = OfficeLength.Centimeters(12)
});
chart.SetTitle("Open tickets")
     .SetLegend(OfficeChartLegendPosition.Bottom)
     .SetCategoryAxisTitle("Month")
     .SetValueAxisTitle("Tickets")
     .SetDataLabels(showValue: true);

chart.UpdateData(new OfficeChartData(new[] { "Jan", "Feb", "Mar", "Apr" }, new[] {
    new OfficeSeries("Tickets", new double[] { 120, 98, 87, 74 })
}));

presentation.Save();
```

## Add a bubble chart

Bubble series take explicit X, Y, and size values. Sizes must be finite and not negative, with one size per point.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;

var data = new OfficeChartData(new[] { "1", "2", "3" }, new[] {
    OfficeSeries.CreateBubble("Portfolio",
        xValues: new double[] { 1, 2, 3 },
        yValues: new double[] { 12, 18, 15 },
        bubbleSizes: new double[] { 9, 36, 16 })
});

using PowerPointDocument presentation = PowerPointDocument.Create("bubble.pptx");
PowerPointSlide slide = presentation.AddSlide();
slide.AddChart(OfficeChartType.Bubble, data, new PowerPointChartOptions {
    Left = OfficeLength.Centimeters(2), Top = OfficeLength.Centimeters(2),
    Width = OfficeLength.Centimeters(22), Height = OfficeLength.Centimeters(12),
    Title = "Risk and value"
});
presentation.Save();
```
