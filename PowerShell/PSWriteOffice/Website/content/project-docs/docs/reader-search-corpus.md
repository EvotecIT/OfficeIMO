---
title: "Search Across Office and Document Files"
description: "Search a bounded folder or explicit file set across Word, Excel, PowerPoint, PDF, Markdown, email, and other Reader-supported formats."
layout: docs
---

Use Reader search when the question spans formats: find a policy term across reports, spreadsheets, presentations, PDFs, Markdown notes, and mail stores without writing one parser per file type.

## Bound the search

```powershell
$matches = Search-OfficeDocument -Path '.\Evidence' -Recurse `
    -Query 'Retention policy' -Extension docx,xlsx,pptx,pdf,md `
    -MaxDocuments 500 -MaximumResults 200 -MaxDegreeOfParallelism 4
```

Document, result, store-item, and concurrency limits are safety and predictability controls. Use `-NoDocumentLimit`, `-AllResults`, or `-AllStoreItems` only when the input scope is already controlled. Extension filters are applied before files consume the document ceiling.

The [search-folder recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Reader/Recipe-Reader-SearchFolder.ps1) searches a mixed `.\Documents` folder with one command and selects the useful result fields.

Use `-IncludePageLocations` when page-aware PDF or paged-document results matter. For downstream retrieval or indexing, continue with [chunks and tables](/docs/pswriteoffice/reader-chunks-tables/).
