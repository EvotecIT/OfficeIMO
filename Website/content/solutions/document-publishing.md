---
title: "Create and Publish Documents from .NET"
description: "Generate Word, Excel, PowerPoint, PDF, HTML, Markdown, CSV, RTF, and OpenDocument output from application data without Microsoft Office."
meta.eyebrow: "Document publishing"
meta.outcome: "Application data turned into dependable deliverables"
meta.primary_label: "See publishing workflows"
meta.primary_url: "/docs/workflows/content-publishing/"
meta.powershell: true
---

Document publishing begins with a user need, not a file format. A finance team may need an editable workbook and a signed PDF. A customer may need an accessible Word report. A knowledge pipeline may need HTML and Markdown from the same source material.

OfficeIMO lets the application use focused document models for those outputs instead of automating desktop Office.

## Pick an authoring model that fits the deliverable

- Use [OfficeIMO.Word](/products/word/) for reports, letters, contracts, templates, review, comments, and flowing pages.
- Use [OfficeIMO.Excel](/products/excel/) for data tables, formulas, charts, worksheets, and interactive analysis.
- Use [OfficeIMO.PowerPoint](/products/powerpoint/) for briefings, data-driven decks, notes, charts, and slide layouts.
- Use [OfficeIMO.Pdf](/products/pdf/) when PDF is the primary document rather than only an export.
- Use Markdown, HTML, CSV, RTF, or OpenDocument packages when those formats are the actual contract.

The focused packages can be combined without forcing every application to carry the whole ecosystem.

## Separate authoring from delivery

Keep business data and document construction in the application. Let storage, authentication, email, signing, and retention remain explicit boundaries. That makes the workflow testable and prevents a template engine from becoming an accidental system of record.

For Word or Excel output that also needs PDF, add the corresponding first-party adapter:

- [Word to PDF](/convert/word-to-pdf/)
- [Excel to PDF](/convert/excel-to-pdf/)

## Validate what the reader experiences

Package validity is necessary but not sufficient. Test representative output for readable pagination, fonts, table widths, image quality, accessibility, links, and the data that matters. OfficeIMO includes validators, diagnostics, generated samples, and visual baselines across the project, while your application can add domain-specific checks.

## Run where the application runs

OfficeIMO targets cross-platform .NET and supports service, worker, desktop, container, CI, and NativeAOT scenarios across most production projects. No Microsoft Office installation is required. See [platform support](/docs/getting-started/platform-support/) and [NativeAOT guidance](/docs/advanced/aot-trimming/) before selecting a deployment profile.
