---
title: "Build Bounded Document Ingestion Pipelines"
description: "Scan selected document formats, collect ingestion summaries, and feed structured chunks into downstream indexing or analysis systems."
layout: docs
---

`Get-OfficeDocumentIngest` combines folder discovery and document reading into one bounded operation. It reports files scanned, files parsed, chunks produced, and errors instead of forcing every caller to rebuild the same orchestration.

## Select the corpus

```powershell
$result = Get-OfficeDocumentIngest -FolderPath '.\Knowledge' `
    -Extension docx,pdf,md,html,json,yaml `
    -MaxFiles 1000 -MaxTotalBytes 2GB
```

Use `-NoRecurse` for a single controlled directory. Extension filters keep unrelated files outside the parser path. Reader options and content bounds should match the same policy used by search and chunk extraction.

The [ingest-folder recipe](https://github.com/EvotecIT/OfficeIMO/blob/25f36d5a7c1c92c35e66c295444f3d4fb7b29ec0/PowerShell/PSWriteOffice/Examples/Reader/Recipe-Reader-IngestFolder.ps1) ingests supported documents from a simple `.\Documents` folder.

Keep indexing, embeddings, databases, or network publication outside the Reader command. Reader owns deterministic document extraction; the downstream system owns storage and retrieval policy.
