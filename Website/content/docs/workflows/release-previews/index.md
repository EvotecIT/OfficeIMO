---
title: "Capabilities in Release Preview"
description: "See which OfficeIMO and PSWriteOffice capabilities are implemented in open pull requests but not yet available from released packages."
order: 9
meta.seo_title: "Upcoming OfficeIMO and PSWriteOffice capabilities"
---

This page helps evaluators distinguish repository code from installable releases. Do not build a production plan around these commands until the linked pull request is merged and the required package version is published.

## PSWriteOffice previews

| Capability | What the implementation adds | Release state |
|---|---|---|
| Mixed document search | `Search-OfficeDocument` across Word, Excel, PowerPoint, PDF, PST, OST, Markdown, and Reader formats, with bounded concurrency and as-completed results | [PSWriteOffice #151](https://github.com/EvotecIT/PSWriteOffice/pull/151); waiting on OfficeIMO.Reader package publication |
| Authenticated PDF automation | Password-aware reading, merge sources, reports, canvas stamping and overlays, rendering options, tables, and header/footer zones | [PSWriteOffice #152](https://github.com/EvotecIT/PSWriteOffice/pull/152); waiting on OfficeIMO PDF package publication |
| Confluence publishing | Sessions, page read/publish/delete, managed sections, attachments, and a database-reporting example | [PSWriteOffice #153](https://github.com/EvotecIT/PSWriteOffice/pull/153); waiting on OfficeIMO Confluence packages |

The Confluence reporting example also depends on [DbaClientX #215](https://github.com/EvotecIT/DbaClientX/pull/215), which adds the Azure Tables provider and PowerShell commands.

## Resource-boundary hardening

The current OfficeIMO pull-request stack adds deterministic limits around resource-heavy paths. These changes protect existing workflows; they do not create new format claims:

- [OfficeIMO #2144](https://github.com/EvotecIT/OfficeIMO/pull/2144): PDF and Excel resource expansion.
- [OfficeIMO #2145](https://github.com/EvotecIT/OfficeIMO/pull/2145): raster, SVG, Reader, and presentation processing.
- [OfficeIMO #2146](https://github.com/EvotecIT/OfficeIMO/pull/2146): remaining parser and export paths including CSV and image export.

## How to evaluate preview code

1. Check the pull request head and dependency notes.
2. Build from source with the linked OfficeIMO projects rather than substituting an older published package.
3. Run the included examples and focused tests.
4. Recheck NuGet or PowerShell Gallery before switching an installation guide to the new public version.

Stable workflows remain documented in the normal product and guide pages. This preview page is the only place where the site intentionally describes commands that are not yet generally installable.
