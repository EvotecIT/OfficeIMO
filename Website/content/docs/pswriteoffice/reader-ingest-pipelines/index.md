---
title: "Build Bounded Document Ingestion Pipelines"
description: "Scan selected document formats, collect ingestion summaries, and feed structured chunks into downstream indexing or analysis systems."
layout: docs
---

`Read-OfficeDocument -As Ingest` combines folder discovery and document reading into one bounded operation. It reports files scanned, files parsed, chunks produced, and errors instead of forcing every caller to rebuild the same orchestration.

## Select the corpus

```powershell
$result = Read-OfficeDocument -As Ingest -Path '.\Knowledge' -Recurse `
    -Extension docx,pdf,md,html,json,yaml `
    -MaximumFiles 1000 -MaximumTotalBytes 2GB
```

Omit `-Recurse` for a single controlled directory. Extension filters keep unrelated files outside the parser path. Reader options and content bounds should match the same policy used by search and chunk extraction.

The [ingest-folder recipe](https://github.com/EvotecIT/OfficeIMO/blob/f04b1719d4344666e44f213727dc28b053dd7493/PowerShell/PSWriteOffice/Examples/Reader/Recipe-Reader-IngestFolder.ps1) ingests supported documents from a simple `.\Documents` folder.

Keep indexing, embeddings, databases, or network publication outside the Reader command. Reader owns deterministic document extraction; the downstream system owns storage and retrieval policy.
