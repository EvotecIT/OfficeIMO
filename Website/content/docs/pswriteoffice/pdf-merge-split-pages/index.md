---
title: "Merge, Split, and Reorder PDF Pages"
description: "Combine PDF files, normalize page sizes, split by page count, range, or bookmarks, and copy, move, or remove selected pages."
layout: docs
---

PDF page operations work on completed files. They are appropriate for delivery packs, scanned batches, extracted chapters, and reordered evidence bundles.

## Merge and normalize

```powershell
Join-OfficePdf -Path '.\Cover.pdf','.\Report.pdf','.\Appendix.pdf' `
    -OutputPath '.\Pack.pdf' -PageSize A4 -ResizeMode Fit -ResizeMargin 18
```

Page normalization is optional. Keep original sizes when fidelity matters; use a fixed target size when the delivery contract requires consistent paper geometry.

## Split by intent

`Split-OfficeDocument` splits PDFs into consecutive page groups, explicit ranges (`-Page '1-3','4,6'`), or top-level bookmarked chapters (`-ByBookmark`). Nested sections stay within their chapter; `-Bookmark` selects exact chapter titles. Outputs use the source base name or `-Prefix`, followed by a part index. Padded indexes make filenames sort predictably. Use `-NoClobber` to keep existing parts; other files in the output folder remain intact.

```powershell
Split-OfficeDocument -Path '.\Pack.pdf' -OutputDirectory '.\Pages' `
    -PageCount 1 -Prefix page -PadIndex -IndexWidth 3
```

Use `Copy-OfficePdfPage`, `Move-OfficePdfPage`, and `Remove-OfficePdfPage` for targeted page assembly. The [merge-and-split recipe](https://github.com/EvotecIT/OfficeIMO/blob/894f02ac68ee967efc3105c566a1d67ba0dbe288/PowerShell/PSWriteOffice/Examples/Pdf/Recipe-Pdf-MergeAndSplit.ps1) creates two small inputs, combines them, and splits the result into one-page files.
