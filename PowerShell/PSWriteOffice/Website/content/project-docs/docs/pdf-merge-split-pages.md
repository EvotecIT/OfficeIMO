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

`Split-OfficePdf` can split every N pages, explicit page ranges, named bookmarks, or bookmark structure. Padded indexes make generated filenames sort predictably.

```powershell
Split-OfficePdf -Path '.\Pack.pdf' -OutputDirectory '.\Pages' `
    -PagesPerDocument 1 -Prefix page -PadIndex -IndexWidth 3
```

Use `Copy-OfficePdfPage`, `Move-OfficePdfPage`, and `Remove-OfficePdfPage` for targeted page assembly. The [merge-and-split recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Pdf/Recipe-Pdf-MergeAndSplit.ps1) creates two small inputs, combines them, and splits the result into one-page files.
