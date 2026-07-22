---
title: "Automate PDF Workflows"
description: "Author, inspect, combine, annotate, sign, preflight, optimize, redact, and extract PDF files."
layout: docs
---

The PDF family exports 74 commands over the first-party OfficeIMO.Pdf engine. It includes document composition, metadata, pages, text, images, tables, lists, headings, panels, headers and footers, backgrounds, watermarks, forms, annotations, attachments, bookmarks, signatures, compliance, extraction, merge/split, redaction, optimization, and diagnostics.

## Compose a fixed-layout document

Use `New-OfficePdf` with the `Add-OfficePdf*` commands when the script owns the PDF. Apply page setup and a theme once, then add semantic headings, paragraphs, lists, tables, images, panels, stamps, and page breaks.

## Inspect before transforming

Read-only commands expose document information, text, images, fonts, attachments, form fields, annotations, signatures, compliance, interactions, optimization data, preflight results, append-only mutation state, redaction plans, and text/layout diagnostics.

This enables a safer pipeline:

1. preflight the source;
2. collect compliance and signature evidence;
3. plan redaction, flattening, optimization, or page changes;
4. write to a new artifact;
5. reopen and test the result.

## Document operations

- combine and separate files with `Join-OfficePdf` and `Split-OfficePdf`;
- copy, move, and remove pages;
- import and export XFDF annotations;
- flatten annotations or forms deliberately;
- sanitize, optimize, or redact with explicit conversion commands;
- add or update metadata, forms, annotations, signatures, compliance, electronic invoice data, page appearance, and themes;
- extract images, text, and layout overlays for downstream review.

PDF features vary by document structure and preservation requirement. Use `Test-OfficePdfRewrite`, diagnostics, and preflight evidence for complex inputs instead of assuming every rewrite is lossless.

See the [PDF examples](https://github.com/EvotecIT/PSWriteOffice/tree/main/Examples/Pdf) and search the [command reference](/api/powershell/) for `OfficePdf`.
