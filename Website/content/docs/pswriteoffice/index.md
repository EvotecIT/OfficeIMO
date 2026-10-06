---
title: "PSWriteOffice Documentation"
description: "Workflow guides for the manifest-derived PowerShell command surface across Office, PDF, Reader, Confluence Cloud, Visio, and open formats."
layout: docs
---

PSWriteOffice is the PowerShell surface for OfficeIMO. Use these guides to choose a workflow, compose or inspect a document, and then move into generated command reference for exact parameters and examples.

## Start here

- [What PSWriteOffice covers](/docs/pswriteoffice/overview/)
- [Install and verify](/docs/pswriteoffice/install/)
- [Choose a workflow](/docs/pswriteoffice/choosing-a-workflow/)
- [Pipeline, object, and DSL workflows](/docs/pswriteoffice/object-workflows/)
- [Use-case index](/docs/pswriteoffice/use-case-index/)
- [DSL cookbook](/docs/pswriteoffice/dsl-cookbook/)
- [Command families](/docs/pswriteoffice/command-families/)
- [PSWriteOffice vs ImportExcel vs ExcelFast](/docs/pswriteoffice/compare-importexcel-excelfast/)
- [PSWriteOffice vs Office Interop, Graph, and LibreOffice](/docs/pswriteoffice/compare-office-automation-options/)
- [Migrate from PSWriteWord, PSWriteExcel, or PSWritePDF](/docs/pswriteoffice/migrate-from-legacy-modules/)

## Document workflows

- [Word automation](/docs/pswriteoffice/word/)
  - [Read and inspect Word](/docs/pswriteoffice/word-read-inspect/)
  - [Update existing Word documents](/docs/pswriteoffice/word-update-existing/)
  - [Merge documents and generate letters](/docs/pswriteoffice/word-merge-mailmerge/)
- [Excel automation](/docs/pswriteoffice/excel/)
  - [Read and import Excel data](/docs/pswriteoffice/excel-read-import/)
  - [Update, merge, and compare workbooks](/docs/pswriteoffice/excel-merge-compare/)
- [PowerPoint automation](/docs/pswriteoffice/powerpoint/)
  - [Inspect and update presentations](/docs/pswriteoffice/powerpoint-read-inspect/)
  - [Reuse and combine slides](/docs/pswriteoffice/powerpoint-reuse-slides/)
- [PDF automation](/docs/pswriteoffice/pdf/)
  - [Read and preflight PDFs](/docs/pswriteoffice/pdf-read-preflight/)
  - [Merge and split PDF pages](/docs/pswriteoffice/pdf-merge-split-pages/)
  - [Position PDF text and graphics](/docs/pswriteoffice/pdf-positioned-content/)
  - [Redact, sanitize, and optimize](/docs/pswriteoffice/pdf-redact-sanitize/)
- [Reader, extraction, and OCR](/docs/pswriteoffice/reader/)
  - [Search a mixed document corpus](/docs/pswriteoffice/reader-search-corpus/)
  - [Extract chunks, tables, and assets](/docs/pswriteoffice/reader-chunks-tables/)
- [Confluence Cloud publishing](/docs/pswriteoffice/confluence/)
- [Visio diagrams](/docs/pswriteoffice/visio/)
- [Markdown, RTF, CSV, OpenDocument, email, AsciiDoc, and LaTeX](/docs/pswriteoffice/open-text-formats/)
  - [Read and parse Markdown](/docs/pswriteoffice/markdown-read-parse/)
  - [Convert and publish Markdown](/docs/pswriteoffice/markdown-convert-publish/)
  - [Convert between Markdown and Word](/docs/pswriteoffice/markdown-word-roundtrip/)
- [Automation patterns](/docs/pswriteoffice/automation-patterns/)
- [Troubleshooting and diagnostics](/docs/pswriteoffice/troubleshooting/)
- [Migrate from PSWriteWord](/docs/pswriteoffice/migrate-from-pswriteword/)
- [Migrate from PSWriteExcel](/docs/pswriteoffice/migrate-from-pswriteexcel/)
- [Migrate from PSWritePDF](/docs/pswriteoffice/migrate-from-pswritepdf/)
- [PSWriteOffice project hub](https://evotec.xyz/projects/pswriteoffice/)

## Notes

- The [command reference](/api/powershell/) is generated from the module manifest and external help.
- The [scenario-driven example library](https://github.com/EvotecIT/PSWriteOffice/blob/v3.0.6/Examples/README.md) maps practical recipes and larger showcases to each format.
- The family totals shown on the site come from `PSWriteOffice.psd1`; documentation validation fails if an exported command is left uncategorized.
