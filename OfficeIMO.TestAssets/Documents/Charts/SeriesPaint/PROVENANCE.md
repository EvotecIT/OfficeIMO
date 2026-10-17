# Independent series fill and outline paint

`generate.py` reuses the [automatic-colour cohort's package writer](../AutomaticColors/PROVENANCE.md)
to create the same sixteen charts in DOCX, XLSX and PPTX: area, bubble, column and line, each
with all four combinations of fill and outline. Fill is blue `156082`; outline is orange
`E97132`, 3 points wide. Line charts explicitly disable markers and smoothing. Grid lines
are explicitly `E2E8F0`, 0.5 points, to isolate series paint from automatic chart-frame defaults.

Microsoft Word, Excel and PowerPoint 16.113.4 on macOS 27.0.1 opened these exact inputs and
exported the checked-in PDFs using **Best for printing**, with the whole workbook selected
in Excel. Each Word page and PowerPoint slide holds one chart. Each Excel worksheet spans
two PDF pages. `office-reference.json` records the source and PDF hashes and producer builds.
The source packages and references use the repository's MIT licence.

Run `python3 OfficeIMO.TestAssets/Documents/Charts/SeriesPaint/generate.py` from the repository
root to reproduce the packages. Open each package in its Office app and export the matching
`microsoft-<app>-16.113-series-paint.pdf`. Then run `verify-native.py`, which inspects vector
paint using MuPDF's `mutool draw -F trace`. Its checks require blue area/bubble/column fills
only when fill is enabled, orange outlines only when outline is enabled, and no series paint
when both are disabled. A line uses outline paint regardless of its fill setting.

The shared-renderer and adapter regressions check these same combinations. Visual comparison
confirms retained unfilled area and bubble outlines and absent paint when both are disabled.
This cohort qualifies series paint, not automatic axis scaling, native bubble diameter
normalization, chart frame layout, font metrics or page placement. Area point segmentation
remains separately diagnosed and tracked in the product roadmap; the [LibreOffice area-point
reference](../LibreOffice/PROVENANCE.md#area-point-paint) paints one series area despite its
stored point fills.
