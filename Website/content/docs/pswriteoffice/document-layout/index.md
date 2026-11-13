---
title: "Readable document layouts"
description: "Build readable PSWriteOffice reports with paragraph spacing, table-cell padding, consistent colours, clear chart units, and page-aware layouts."
layout: docs
---

A report is easier to scan when headings, tables and charts have their own space. Start with a restrained palette, a readable body size and a clear distinction between the headline result and its supporting detail.

## Separate a table from a chart

Place the chart inside a `WordParagraph` block and set spacing on that paragraph. `WordChart` uses the active paragraph, so the spacing belongs to the chart's position in the document. Avoid inserting blank paragraphs to create a gap.

```powershell
WordParagraph {
    Set-OfficeWordParagraph -SpacingBefore 14 -SpacingAfter 12
    WordChart -Type Bar -InputObject $rows -Category Site `
        -Value Patched, Pending -Title 'Patched and pending by site' `
        -Legend -FitToPageWidth -WidthFraction 1 -Height 260px
}
```

Keep headings with the content that follows them:

```powershell
WordParagraph {
    Set-OfficeWordParagraph -Style Heading1 -SpacingBefore 16 `
        -SpacingAfter 8 -KeepWithNext
    WordText 'Compliance by site' -Bold -Size 15pt -Color '17365D'
}
```

## Give table rows room

Use `-PassThru` to retain the table object, select its cells and set their margins. `Set-OfficeWordTableCell` accepts explicit physical units for margins; bare values and paragraph spacing use points.

```powershell
$table = WordTable -InputObject $rows -Style GridTable4Accent1 `
    -Layout AutoFitToWindow -PassThru
$table | Get-OfficeWordTableCell | Set-OfficeWordTableCell `
    -MarginTop 0.12cm -MarginBottom 0.12cm `
    -MarginLeft 0.16cm -MarginRight 0.16cm
```

Each margin is independent. Omitted parameters retain the existing margin. Passing `$null` removes that explicit margin so the document's table style can supply it. Zero is an explicit zero margin.

Choose the columns for the audience rather than sending every object property to the table. A long action description often belongs in an action section instead of a narrow scorecard column.

## Keep charts interpretable

Label units and the observation period. Counts and percentages need separate charts or a clearly explained axis arrangement. A series of availability percentages near 100 and a series of incident counts near zero do not share a useful scale.

Use the chart height that supports the message. A five-category chart in a one-page report rarely needs the same height as a full-page analysis. Recheck labels and legends after changing the size.

## Check the delivered file

Run the [patch compliance recipe](https://github.com/EvotecIT/OfficeIMO/blob/d3d39792d6c1c0cc3e4945e9ec5767e83822c239/PowerShell/PSWriteOffice/Examples/Word/Recipe-Word-PatchCompliance.ps1), open the DOCX and inspect the exported PDF. Check every page for clipped text, split headings, table widths and chart labels. Native PDF conversion can paginate differently from a desktop Word editor; the PDF is the delivery copy when its layout is the agreed result.
