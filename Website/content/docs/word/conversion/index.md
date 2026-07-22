---
title: "Word Conversion and Rendering"
description: "Choose Word-to-HTML, Markdown, PDF, OpenDocument, RTF, Google Docs, or image workflows with explicit fidelity and resource policies."
layout: docs
---

Word conversion is split into focused adapters so applications do not acquire every renderer and destination model. Start with `OfficeIMO.Word`, then add the package for the destination you ship.

## Routes

| Destination | Package | What to validate |
|---|---|---|
| HTML and MHTML | `OfficeIMO.Word.Html` | CSS, images, links, headers/footers, lists, tables, review markup, and external-resource policy |
| Markdown | `OfficeIMO.Word.Markdown` | Heading/list semantics, tables, code, images, links, and features with no Markdown equivalent |
| PDF | `OfficeIMO.Word.Pdf` | Pagination, fonts, shaping, fields, tables, images, headers/footers, links, and renderer diagnostics |
| OpenDocument | `OfficeIMO.Word.OpenDocument` | DOCX/ODT model mapping and round-trip expectations |
| RTF | `OfficeIMO.Word.Rtf` | Rich-text semantics, lists, tables, and bounded legacy-format behavior |
| Google Docs | `OfficeIMO.Word.GoogleDocs` | Conversion loss policy, authentication, Drive ownership, and remote update behavior |
| Images | `OfficeIMO.Word` imaging APIs | Page size, resolution, fonts, and final-revision view |

## Use diagnostics as output

A successful method call only proves that an artifact was produced. Treat warnings as part of the conversion result and decide which codes are acceptable. For regulated or high-value output, store the source hash, adapter version, selected options, diagnostics, and destination hash together.

## Review-aware conversion

Decide whether the destination should represent the original view, final revision view, visible comments, or a clean approved document. Inspect review state before conversion and fail the job when unresolved comments or unsupported metadata violate policy.

## Resource policy

HTML, images, and PDF output can depend on fonts, linked resources, embedded assets, and network access. Prefer explicit local or allow-listed resolvers. Avoid a converter silently reaching arbitrary URLs in a server or build environment.

Use the [conversion map](/docs/capabilities/conversions/) for other source families and the generated [Word API reference](/api/word/) for exact option types.
