---
title: "Space tables and charts in a Word report"
description: "Create a patch compliance report with padded table cells, a separated chart and a PDF delivery copy."
layout: docs
---

The [patch compliance recipe](https://github.com/EvotecIT/OfficeIMO/blob/master/PowerShell/PSWriteOffice/Examples/Word/Recipe-Word-PatchCompliance.ps1) turns sample inventory objects into a one-page Word report and PDF. It keeps the headline, per-site table, chart and next steps visually separate.

```powershell
./Examples/Word/Recipe-Word-PatchCompliance.ps1 -OutputDirectory ./Output
```

Replace the inventory objects first. Keep the `Site`, `Servers`, `Patched`, `Pending` and `Status` properties, or update the selected columns and chart properties together.

The table uses conditional shading for sites that need attention and explicit cell margins for readable rows. The chart sits inside a paragraph with 14 points of space before it and 12 after it. Section headings use `KeepWithNext` so a heading stays with the following content.

The complete script includes ordinary module loading, the sample data, the DOCX generation and PDF export. See [readable document layouts](/docs/pswriteoffice/document-layout/) for the smaller composition patterns and their units.
