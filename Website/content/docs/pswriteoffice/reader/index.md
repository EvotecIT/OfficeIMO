---
title: "Reader, Document Extraction, and OCR"
description: "Detect formats and extract normalized documents, chunks, hierarchy, tables, visuals, assets, OCR text, and ingest results. Includes examples and cmdlet links."
layout: docs
---

Reader is the cross-format path for search, review, ingestion, migration, OCR, and AI preparation. Its 11 exported commands expose the normalized OfficeIMO.Reader model and the easy local image OCR entry point without forcing the caller to switch on every native document type.

## What Reader returns

- `Read-OfficeDocument` returns the normalized document envelope.
- `Get-OfficeDocumentInfo` reports detected format and routing evidence.
- `Get-OfficeDocumentCapability` explains the registered adapters' supported output.
- `Read-OfficeDocument -As Hierarchy` exposes section and structural relationships.
- `-As Chunks` returns bounded text/content chunks; `-As Structure` exposes structured content.
- `-As Tables`, `-As Visuals`, and `-As Assets` return non-text material explicitly.
- `Read-OfficeDocument -Path .\Knowledge -Recurse -As Ingest` returns a corpus ingestion summary.
- `Search-OfficeDocument` searches the normalized result.
- `Read-OfficeDocument -Path .\scan.png -As Text -Ocr` performs local image OCR; OCR settings come from `New-OfficeOcrOptions`.

## Choose Reader versus a native family

Use Reader when downstream behavior is common across file formats: indexing, chunking, extraction, discovery, classification, previews, and bulk ingest. Use Word, Excel, PowerPoint, PDF, Visio, or another native family when the script must modify format-specific objects and preserve their native semantics.

Reader adapters are modular. A deployment can carry only the adapters it needs, while broader presets compose multiple local formats. Optional OCR, web, and process-backed integrations have separate dependency and trust boundaries.

## Task guides

- [Search across Office and document files](/docs/pswriteoffice/reader-search-corpus/)
- [Extract chunks, tables, and assets](/docs/pswriteoffice/reader-chunks-tables/)
- [Build bounded ingestion pipelines](/docs/pswriteoffice/reader-ingest-pipelines/)

## Operational pattern

Detect first, inspect capabilities, extract the required surfaces, and record warnings/provenance with the output. Do not silently reduce a document to plain text when tables, visuals, or assets matter to the consumer.

See the [document Reader examples](https://github.com/EvotecIT/OfficeIMO/tree/f04b1719d4344666e44f213727dc28b053dd7493/PowerShell/PSWriteOffice/Examples/Reader) and search the [command reference](/api/powershell/) for `OfficeDocument`.
