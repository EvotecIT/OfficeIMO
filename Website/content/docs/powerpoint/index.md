---
title: PowerPoint Presentations
description: Overview of the OfficeIMO.PowerPoint package for generating .pptx decks with slides, text, tables, charts, and notes.
order: 30
---

# PowerPoint Presentations

`OfficeIMO.PowerPoint` lets you build `.pptx` presentations from a higher-level object model instead of stitching raw Open XML parts together yourself. It is designed for generated decks, reporting pipelines, demo content, and repeatable presentation output that needs to stay readable in source control.

## Good fit scenarios

- Create status decks, release reviews, and customer reports from application data.
- Generate training material or runbook slides as part of CI, build, or scheduled jobs.
- Reuse the same source data across Word, Excel, PowerPoint, and Reader pipelines.
- Produce presentations on servers, containers, or developer machines without PowerPoint installed.

## Core building blocks

| Type | Purpose |
|------|---------|
| `PowerPointPresentation` | Root document for creating, loading, saving, and organizing slides |
| `PowerPointSlide` | Individual slide surface for titles, text boxes, tables, charts, images, and notes |
| `PowerPointTextBox` | Rich text container with bullets, formatted runs, margins, and auto-fit |
| `PowerPointTable` | Strongly typed table surface with styles, column widths, row heights, and merged cells |
| `PowerPointChart` | Chart host for series-driven visualizations inside a slide |
| `PowerPointNotes` | Speaker notes attached to a slide |
| `PowerPointLayoutBox` | Measurement helper for content regions, columns, and spacing |

## Quick start

```csharp
using OfficeIMO.PowerPoint;

using var presentation = PowerPointPresentation.Create("status-update.pptx");
const double marginCm = 1.5;

var content = presentation.SlideSize.GetContentBoxCm(marginCm);
var slide = presentation.AddSlide();

var title = slide.AddTitleCm(
    "Weekly Status",
    content.LeftCm,
    content.TopCm,
    content.WidthCm,
    1.4);

if (title.Paragraphs.Count > 0)
{
    PowerPointTextStyle.Title.WithColor("1F4E79").Apply(title.Paragraphs[0]);
}

var agenda = slide.AddTextBoxCm(
    string.Empty,
    content.LeftCm,
    content.TopCm + 2.0,
    content.WidthCm,
    content.HeightCm - 2.0);

agenda.AddBullets(new[]
{
    "Deployment completed successfully",
    "Customer onboarding toolkit updated",
    "Performance metrics improved week over week"
});

presentation.Save();
```

## Recommended workflow

1. Create a presentation and get a content box from `SlideSize` so positioning stays consistent.
2. Add slides and establish structure with title, body, and supporting shapes or tables.
3. Populate content from your domain data, not from hard-coded presentation markup.
4. Add speaker notes, sections, theme tweaks, or layout helpers when the deck grows.
5. Save the file and ship it as a build artifact, report attachment, or generated deliverable.

## Layout and content model

- Use `SlideSize.GetContentBoxCm(...)` to reserve consistent margins around the live slide area.
- Use `GetColumnsCm(...)` when you want balanced two- or three-column compositions.
- Work in centimeters, points, or EMUs through the built-in unit helpers depending on the scenario.
- Add titles, text boxes, tables, charts, images, and notes to the same `PowerPointSlide` surface.

## Related workflows

- Pair `OfficeIMO.PowerPoint` with `OfficeIMO.Excel` when chart/table data already exists in workbook form.
- Use `OfficeIMO.Reader` when the same pipeline later needs extraction or chunking from generated decks.
- Use PSWriteOffice when you want the same presentation automation from PowerShell rather than C#.

## Further reading

- [Slides](/docs/powerpoint/slides) -- Creating slides with text boxes, shapes, images, and charts.
- [PSWriteOffice PowerPoint Cmdlets](/docs/pswriteoffice/powerpoint/) -- Build decks from PowerShell with DSL aliases and cmdlets.
- [OfficeIMO.PowerPoint product page](/products/powerpoint/) -- Package-level overview, install command, and positioning.
