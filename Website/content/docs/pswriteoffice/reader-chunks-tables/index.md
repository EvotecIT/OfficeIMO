---
title: "Extract Document Chunks, Tables, and Assets"
description: "Produce semantic chunks, structured table exports, page Markdown, hierarchy, and embedded assets for downstream processing."
layout: docs
---

Reader extraction turns supported files into stable intermediate structures. That is useful for search indexes, retrieval systems, review sidecars, data migration, and document inventories.

## Chunk content

`Read-OfficeDocument -As Chunks` emits bounded text chunks with location and provenance. Format-specific options control Word notes, PowerPoint notes, Excel headers and row groups, Markdown heading chunks, hashes, and maximum content sizes. Set `-MaximumChunkCharacters` and `-MaximumTableRows` to the downstream limits.

## Materialize tables

`Read-OfficeDocument -As Tables` returns structured tables. Add `-Bundle` for CSV, Markdown, and JSON representations in memory. `Export-OfficeDocumentAsset -Type Table -OutputDirectory .\Tables` writes deterministic sidecars. Use `Read-OfficeDocument -As Assets` to inspect embedded assets and `Export-OfficeDocumentAsset -OutputDirectory .\Assets` to materialize them.

The [extract-tables recipe](https://github.com/EvotecIT/OfficeIMO/blob/e71f09178f1e1231bd7901cbda405bd4f6b4718b/PowerShell/PSWriteOffice/Examples/Reader/Recipe-Reader-ExtractTables.ps1) creates a Markdown table, reads it as structured data, writes its sidecars, and returns semantic chunks.

Choose limits from the downstream contract. A search preview, an embedding pipeline, and a complete archival export need different chunk and table bounds.
