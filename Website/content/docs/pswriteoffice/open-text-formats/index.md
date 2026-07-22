---
title: "Open Formats and Text Automation"
description: "Use the smaller PSWriteOffice families for text, interchange, open formats, and managed message workflows."
layout: docs
---

PSWriteOffice is not limited to the three desktop Office formats. Smaller command families expose the same managed-engine approach for text, interchange, open document, and email workflows.

## Markdown

Twenty-five commands build and inspect typed Markdown. Add headings, paragraphs, lists, task lists, tables, code, callouts, details, front matter, images, quotes, definition lists, and tables of contents. Reader commands expose headings, nodes, tables, and front matter; converters bridge HTML and Word workflows.

## RTF

Five canonical commands create, load, update, convert, and inspect Rich Text Format documents. Use RTF when a lightweight rich-text interchange file is the required source or destination, and keep loss-aware conversion diagnostics for complex content.

## CSV

Five commands convert, import, export, and inspect CSV through OfficeIMO.CSV. Use the CSV family for delimited-data contracts; use Excel when worksheet formatting, formulas, charts, or workbook structure are part of the outcome.

## OpenDocument

Five commands create, read, convert, and save ODT, ODS, and ODP artifacts through OfficeIMO.OpenDocument. These are native managed workflows rather than LibreOffice automation.

## Email

Four commands load and save messages and mailbox artifacts through OfficeIMO.Email. The underlying engine covers multiple message, personal-information, store, and address-book families; exact support and diagnostics belong to the generated command/API reference.

## AsciiDoc and LaTeX

Each family provides four bounded interoperability commands for reading, saving, and bridging through Markdown. These are explicit profiles, not a claim to implement every extension or package in the wider AsciiDoc or TeX ecosystems.

## HTML review

Office format families provide focused HTML conversion commands, and `Export-OfficeHtmlImage` supports extracted assets. Use the [HTML review examples](https://github.com/EvotecIT/PSWriteOffice/tree/main/Examples) to build a browser-review step without discarding the source document.
