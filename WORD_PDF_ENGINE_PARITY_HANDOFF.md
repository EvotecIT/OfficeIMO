# OfficeIMO Word-to-PDF Engine Parity Handoff

## Context

OfficeIMO PR #1976 was merged into `master` on 2026-06-19.

The goal of that PR was to improve the native `OfficeIMO.Pdf` / `OfficeIMO.Word.Pdf` Word-to-PDF path so it behaves more like Microsoft Word's own "Save as PDF" output. The work was intentionally engine-level and reusable. It was not meant to encode anything specific to TestimoX, Active Directory, IAM, audit reports, or any other domain.

TestimoX was used only as a real-world stress case because it produces complex Word documents with:

- Headers and footers
- Table of contents
- Charts
- Dense tables
- Lists and assertions
- Section/page layout
- Paragraph and table styles
- Repeated report structures

The larger product requirement is generic: OfficeIMO should be able to produce beautiful Word documents and beautiful PDFs from the same Word/document model, with native PDF output closely matching what Word itself would export.

## Core Problem

The native OfficeIMO Word-to-PDF engine does not yet fully match Word's layout semantics.

A `.docx` can look correct in Word, and Word's own "Save as PDF" will preserve that layout. OfficeIMO, however, must interpret WordprocessingML and independently recreate the visual result in PDF. Mismatches appear when Word behavior depends on subtle layout rules, especially around spacing, table sizing, list indentation, style inheritance, headers/footers, and pagination.

The next work should continue improving the reusable engine, not the TestimoX report generator.

## What PR #1976 Improved

- Document defaults flow into PDF more accurately.
- Paragraph style defaults and table style defaults are mapped more accurately.
- Native PDF layout handles more Word-like text metrics, wrapping, spacing, hyphenation, and table sizing.
- Header and footer offsets now preserve caller-provided `PdfOptions.HeaderOffsetY` and `PdfOptions.FooterOffsetY`.
- Table cell widths such as `2400` twips are treated as explicit widths instead of being ignored.
- Additional focused tests and visual baselines were added around Word `SaveAsPdf` behavior.
- The implementation stayed generic inside `OfficeIMO.Word.Pdf` and `OfficeIMO.Pdf`.

## Baseline Artifacts

Primary DOCX originally used for visual inspection:

```text
C:\Support\GitHub\TestimoX-exporter-powerpoint-excel-proof\Artifacts\LiveExporterComparison\20260617-125016\TestimoX-live-10.docx
```

That exact `20260617-125016` folder contains DOCX/HTML/PPTX/XLSX and metadata, but no PDF at the top level.

Later comparison runs do include PDF exports. The latest PDF-bearing run found during handoff was:

```text
C:\Support\GitHub\TestimoX-exporter-powerpoint-excel-proof\Artifacts\LiveExporterComparison\20260617-201357\TestimoX-live-10.docx
C:\Support\GitHub\TestimoX-exporter-powerpoint-excel-proof\Artifacts\LiveExporterComparison\20260617-201357\TestimoX-live-10-excel-save-as.pdf
```

Important caveat: the PDF filename says `excel-save-as.pdf`. Before treating it as the canonical Word Save-as-PDF reference, verify how that artifact was actually produced.

Other PDF-bearing comparison folders found:

```text
C:\Support\GitHub\TestimoX-exporter-powerpoint-excel-proof\Artifacts\LiveExporterComparison\20260617-193610
C:\Support\GitHub\TestimoX-exporter-powerpoint-excel-proof\Artifacts\LiveExporterComparison\20260617-194452
C:\Support\GitHub\TestimoX-exporter-powerpoint-excel-proof\Artifacts\LiveExporterComparison\20260617-194917
C:\Support\GitHub\TestimoX-exporter-powerpoint-excel-proof\Artifacts\LiveExporterComparison\20260617-195709
C:\Support\GitHub\TestimoX-exporter-powerpoint-excel-proof\Artifacts\LiveExporterComparison\20260617-200153
C:\Support\GitHub\TestimoX-exporter-powerpoint-excel-proof\Artifacts\LiveExporterComparison\20260617-200549
C:\Support\GitHub\TestimoX-exporter-powerpoint-excel-proof\Artifacts\LiveExporterComparison\20260617-200954
C:\Support\GitHub\TestimoX-exporter-powerpoint-excel-proof\Artifacts\LiveExporterComparison\20260617-201357
```

## OfficeIMO Visual Baselines

Merged visual baselines live under:

```text
C:\Support\GitHub\OfficeIMO\OfficeIMO.Tests\Pdf\VisualBaselines
```

Useful starting baselines:

```text
officeimo-pdf-native-word-report.page1.poppler.png
officeimo-pdf-native-word-daily-layout.page1.poppler.png
officeimo-pdf-lists-tables.page1.poppler.png
officeimo-pdf-headers-footers.page1.poppler.png
officeimo-pdf-headers-footers.page2.poppler.png
officeimo-pdf-professional-report.page1.poppler.png
officeimo-pdf-native-word-table-cell-picture-control.page1.poppler.png
```

Relevant Word `SaveAsPdf` test files live under:

```text
C:\Support\GitHub\OfficeIMO\OfficeIMO.Tests\Pdf
```

High-value files to inspect first:

```text
Word.SaveAsPdf.ParagraphFormatting.cs
Word.SaveAsPdf.TableRowsAndCells.cs
Word.SaveAsPdf.TableStyleMapping.cs
Word.SaveAsPdf.HeaderFooterVariants.cs
Word.SaveAsPdf.Sections.Basic.cs
Word.SaveAsPdf.Charts.cs
Word.SaveAsPdf.Lists.cs
Word.SaveAsPdf.ListTests.cs
Word.SaveAsPdf.TablePlacementWidth.cs
Word.SaveAsPdf.TableMarginsSpacing.cs
Word.SaveAsPdf.TableHeaders.cs
```

## Next Agent Focus

Focus on generic Word layout fidelity:

- Paragraph spacing: `before`, `after`, line spacing, collapsed spacing, style inheritance, and interactions with headings/lists/tables.
- Tables: fixed vs auto-fit layout, preferred widths in twips/percent/auto, cell margins, borders, padding, merged cells, repeating headers, and overflow handling.
- Lists: bullet/number indentation, hanging indents, nested list levels, and text alignment.
- Pagination: keep-with-next, keep-lines, explicit page breaks, section breaks, table row splitting, and body frame interaction with headers/footers.
- Style cascade: document defaults -> paragraph/table/list styles -> direct formatting.
- Visual verification: compare native OfficeIMO PDF output against Word's own Save-as-PDF output where possible.

## Recommended Workflow

1. Start from current `OfficeIMO` `master`.
2. Use the TestimoX artifacts only as a broad visual stress baseline.
3. Do not add TestimoX/domain-specific branches in OfficeIMO.
4. For every observed mismatch, create or reuse a small generic Word fixture that isolates one Word behavior.
5. Compare:
   - Word document appearance in Word
   - Word's own Save-as-PDF output
   - OfficeIMO native `SaveAsPdf` output
6. Fix the reusable layout engine in `OfficeIMO.Word.Pdf` / `OfficeIMO.Pdf`.
7. Add focused tests and update visual baselines only when they represent the intended generic behavior.

## Short Handoff Prompt

```text
Focus on OfficeIMO.Pdf native Word-to-PDF layout parity, not TestimoX-specific rendering. TestimoX is only a real-world stress fixture. The engine must generically interpret WordprocessingML so PDF output matches Word Save as PDF behavior for spacing, tables, lists, style inheritance, headers/footers, and pagination.

Start with the TestimoX DOCX at 20260617-125016, and compare against later PDF-export artifacts such as 20260617-201357 after confirming how the PDF was produced. Then add smaller generic OfficeIMO Word fixtures for each mismatch: spacing, tables, list indentation, headers/footers, style cascade, and pagination.

Do not add domain-specific logic for TestimoX, AD, IAM, audit reports, invoices, legal docs, or any specific document type. The engine should work for all of them by honoring Word layout semantics.
```
