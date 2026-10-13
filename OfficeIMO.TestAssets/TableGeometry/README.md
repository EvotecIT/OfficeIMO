# Native table geometry fixtures

These two independently authored ReportLab PDFs contain the same four-column service ledger at page rotations of 0 and 90 degrees. The ruled columns measure 240, 72, 96, and 120 points. Both files use a 612 by 792 point MediaBox and an identical unrotated content stream; the second file changes the PDF `/Rotate` property to 90. These inputs cover page rotation while retaining native content coordinates. The existing Pango/Cairo `latin-90` and `latin-270` inputs in `MultilingualLayout` cover rotated content whose restored visual column strips progress vertically and share the same X projection.

`manifest.json` records the authored ruling coordinates, text anchors and alignment, expected ordered cells, producer version, and PDF SHA-256 hashes. Boundary coordinates refer to the centre of each 0.5-point ruling in default PDF user space. Each row is 35 points high. Helvetica text uses a 12-point size, a baseline 22 points below the row's upper ruling, and an 8-point inset from the selected cell edge. The numeric columns align right. The repeated `89.00` values and negative credit retain meaningful numeric and duplicate-text checks.

The labels come from the generator's explicit coordinates and cell values. OfficeIMO does not generate these inputs or derive their labels. The manifest's normalized-width tolerances apply to propagation from the canonical PDF reader's detected column progression lengths into saved and reopened XLSX/ODS documents. The ODS tolerance allows its existing conversion between Excel character widths and physical widths. These limits do not qualify recovery of the authored ruled grid. Compare detected widths with the authored grid separately and retain that difference as evidence; `sourceGridRecoveryQualified` is false.

Create and activate a disposable Python virtual environment outside the repository, using configured development scratch storage where available. From the repository root, install ReportLab 4.4.9 into that environment and generate the native-only inputs:

```sh
python -m pip install --no-cache-dir reportlab==4.4.9
python OfficeIMO.TestAssets/TableGeometry/generate.py
```

The generator checks the ReportLab version and uses invariant PDF metadata and uncompressed streams. Running it again with the pinned version reproduces the PDF and manifest bytes. No rasterization, OCR provider, or external font file is required.

The generator, labels, and generated PDFs use the repository's [MIT license](../../LICENSE). ReportLab uses its BSD license and remains contributor-only generation tooling. Helvetica is an unembedded PDF standard 14 font; these fixtures do not redistribute a font binary. None of the generator's Python dependencies enters an OfficeIMO runtime package.

`ruling-controls-manifest.json` labels a separate 19-case admission corpus. Generate it with the same disposable ReportLab environment:

```sh
python OfficeIMO.TestAssets/TableGeometry/generate_ruling_controls.py
```

`PdfClosedRulingGridTests` compares recovered geometry directly with the authored coordinates. Its two original ledger cases qualify the complete visible closed-grid profile at page rotations of 0 and 90 degrees. The controls cover clipping, incomplete and competing boundaries, conflicting text assignments, content rotation, ordinary decoration, invisible strokes and later overpaint. Rejected controls must retain the text-only table's cells and geometry. The original manifest's propagation tolerances and `sourceGridRecoveryQualified: false` remain scoped to the spreadsheet-width propagation contract; they are not the acceptance criteria for these separate grid-recovery tests. Ruling-grid recovery combined with `/UserUnit`, merged or spanning cells, and the native Numbers table-count discrepancy remain outside this qualification.
