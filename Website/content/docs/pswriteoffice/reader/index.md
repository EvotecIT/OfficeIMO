---
title: "Reader and Document Extraction"
description: "Detect formats and extract normalized documents, chunks, hierarchy, tables, visuals, assets, and ingest results."
layout: docs
---

Reader is the cross-format path for search, review, ingestion, migration, and AI preparation. Its 13 exported commands expose the normalized OfficeIMO.Reader model without forcing the caller to switch on every native document type.

## What Reader returns

- `Get-OfficeDocument` returns the normalized document envelope.
- `Get-OfficeDocumentDetection` reports detected format and routing evidence.
- `Get-OfficeDocumentCapability` explains the selected adapter's supported output.
- `Get-OfficeDocumentHierarchy` exposes section and structural relationships.
- `Get-OfficeDocumentChunk` returns bounded text/content chunks.
- `Get-OfficeDocumentStructured` exposes structured content.
- `Get-OfficeDocumentTable`, `Get-OfficeDocumentVisual`, and `Get-OfficeDocumentAsset` return non-text material explicitly.
- `Get-OfficeDocumentBatch` and `Get-OfficeDocumentIngest` support pipeline-scale processing.
- `Search-OfficeDocument` searches the normalized result.

## Choose Reader versus a native family

Use Reader when downstream behavior is common across file formats: indexing, chunking, extraction, discovery, classification, previews, and bulk ingest. Use Word, Excel, PowerPoint, PDF, Visio, or another native family when the script must modify format-specific objects and preserve their native semantics.

Reader adapters are modular. A deployment can carry only the adapters it needs, while broader presets compose multiple local formats. Optional OCR, web, and process-backed integrations have separate dependency and trust boundaries.

## Operational pattern

Detect first, inspect capabilities, extract the required surfaces, and record warnings/provenance with the output. Do not silently reduce a document to plain text when tables, visuals, or assets matter to the consumer.

See the [document Reader examples](https://github.com/EvotecIT/PSWriteOffice/tree/main/Examples/Documents) and search the [command reference](/api/powershell/) for `OfficeDocument`.
