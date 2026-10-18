---
title: "Extract Document Chunks, Tables, and Assets"
description: "Produce semantic chunks, structured table exports, page Markdown, hierarchy, and embedded assets for downstream processing."
layout: docs
---

Reader extraction turns supported files into stable intermediate structures. That is useful for search indexes, retrieval systems, review sidecars, data migration, and document inventories.

## Chunk content

`Get-OfficeDocumentChunk` emits bounded text chunks with location and provenance. Format-specific options control Word notes, PowerPoint notes, Excel headers and row groups, Markdown heading chunks, hashes, and maximum content sizes.

## Materialize tables

`Get-OfficeDocumentTable` returns structured tables. `-AsExport` produces CSV, Markdown, and JSON representations in memory; `-OutputDirectory` writes deterministic sidecars. `Get-OfficeDocumentAsset` similarly exposes or materializes embedded images and other supported assets.

The [extract-tables recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Reader/Recipe-Reader-ExtractTables.ps1) creates a Markdown table, reads it as structured data, writes its sidecars, and returns semantic chunks.

Choose limits from the downstream contract. A search preview, an embedding pipeline, and a complete archival export need different chunk and table bounds.
