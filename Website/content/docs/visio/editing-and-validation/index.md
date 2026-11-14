---
title: "Edit and Validate Visio Diagrams"
description: "Load a VSDX file, find and update shapes by text or Shape Data, work with layers, and validate the result with OfficeIMO.Visio."
meta.seo_title: "Edit and validate existing VSDX Visio diagrams in .NET"
order: 54
packages:
  - OfficeIMO.Visio
cmdlets:
  - Open-OfficeVisio
  - Save-OfficeVisio
limits: "Validation checks package and diagram rules that OfficeIMO knows about; it does not guarantee identical rendering in every Visio viewer."
search_terms:
  - edit vsdx
  - visio shape data
  - visio layers
  - validate vsdx
verified: 2026-11-14
---

A loaded diagram uses the same page, shape, and connector API as a new one. Select shapes by text, layer, master, or Shape Data, and apply one change to the whole selection.

## List pages and shapes

### C#

```csharp
using OfficeIMO.Visio;

using VisioDocument document = VisioDocument.Load("operations.vsdx");

foreach (VisioPage page in document.Pages) {
    Console.WriteLine($"{page.Name}: {page.Shapes.Count} shapes, {page.Connectors.Count} connectors");
    foreach (VisioShape shape in page.Shapes) {
        Console.WriteLine($"  {shape.Id} {shape.Text}");
    }
}
```

### PowerShell

```powershell
$diagram = Open-OfficeVisio -Path .\operations.vsdx
$diagram.Pages | ForEach-Object { '{0}: {1} shapes' -f $_.Name, $_.Shapes.Count }
```

## Update shapes that match text or data

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Visio;

using VisioDocument document = VisioDocument.Load("operations.vsdx");
VisioPage page = document.GetPage("Operations");

page.SelectWithData("Owner", "Ops")
    .Fill(OfficeColor.LightBlue)
    .ShapeData("Reviewed", "Yes", "Reviewed", VisioShapeDataType.Boolean);

page.SelectContainingText("Legacy")
    .Text(shape => shape.Text!.Replace("Legacy", "Production", StringComparison.Ordinal));

document.Save("operations.updated.vsdx");
```

### PowerShell

```powershell
$diagram = Open-OfficeVisio -Path .\operations.vsdx
$diagram.Pages[0].Shapes | Where-Object Text -like '*Legacy*' | ForEach-Object {
    $_.Text = $_.Text -replace 'Legacy', 'Production'
}
$diagram | Save-OfficeVisio -Path .\operations.updated.vsdx
```

## Add layers and Shape Data

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Visio;

using VisioDocument document = VisioDocument.Create("architecture.vsdx");
VisioPage page = document.AddPage("Architecture", OfficeLength.Inches(11), OfficeLength.Inches(8.5));

page.AddLayer("Infrastructure");
page.AddLayer("Annotations").Print = false;

VisioShape server = page.AddRectangle(OfficeLength.Inches(2), OfficeLength.Inches(5), OfficeLength.Inches(2), OfficeLength.Inches(1), "API server");
server.SetShapeData("Owner", "Platform", "Owner", VisioShapeDataType.String, "Owning support team");

VisioShape note = page.AddRectangle(OfficeLength.Inches(6), OfficeLength.Inches(5), OfficeLength.Inches(2), OfficeLength.Inches(1), "Internal note");

page.AddToLayer("Infrastructure", server)
    .AddToLayer("Annotations", note);

document.Save();
```

## Validate a diagram

Validate the document in memory before saving, or a file on disk. Both return an empty list when nothing is wrong.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Visio;

using VisioDocument document = VisioDocument.Create("check.vsdx");
VisioPage page = document.AddPage("Page 1");
page.AddRectangle(OfficeLength.Inches(2), OfficeLength.Inches(2), OfficeLength.Inches(1.5), OfficeLength.Inches(0.75), "Start");

IReadOnlyList<string> inMemoryIssues = document.Validate();
document.Save();

IReadOnlyList<string> savedIssues = VisioValidator.Validate("check.vsdx");
Console.WriteLine($"{inMemoryIssues.Count} in memory, {savedIssues.Count} on disk");
```
