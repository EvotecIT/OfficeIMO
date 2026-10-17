# Automatic chart colours

`generate.py` writes three packages with the same 19 charts, `automatic-colors.pptx`
(one chart per slide), `automatic-colors.docx` (one chart per page) and
`automatic-colors.xlsx` (one chart per worksheet, beside the cells it plots), plus
`line-markers.pptx`. No series or data point stores a colour, a marker symbol or a
line, so the producer chooses them. The charts cover classic chart styles 1, 2 and 4, three to twelve
series, column, bar, line, area, scatter, pie, doughnut and stock groups, and the
modern `cs:chartStyle` (style 201, copied from a PowerPoint-written chart in
`chart-style-201.xml`) and `cs:colorStyle` parts (Colourful palette 1 and
Monochromatic palette 1).

Microsoft PowerPoint, Word and Excel 16.113.4 (16.113.26100421) on macOS 27.0.1
exported the PDFs in this folder. `office-reference.json` lists, for each chart,
the distinct fills Office drew in first-drawn order, read from the PDFs with
`mutool trace`. PowerPoint and Word draw identical fills for every chart; Excel
spreads each worksheet over two pages and draws the same fills.

The exports show that:

- Office ignores the `cs:chartStyle` and `cs:colorStyle` parts when a series has no
  colour of its own. The colourful and monochromatic colour styles render exactly
  like the classic chart without them.
- Classic style 2 (and every eighth style from it) cycles accents 1 to 6, style 1
  cycles six tints of text 1, and styles 3 to 8 use one accent each (style 4 is
  accent 2).
- Every pass through the pattern is shaded or tinted in linear light by an amount
  spread between -70% and +70% over `count / 6 + 1` passes: six or eight series are
  darkened by 23.3% (`115472` for accent 1 `156082`), and twelve series darken the
  first six by 35% and leave the second six unchanged. A single-accent style
  spreads the same range over all of its series.
- Line and scatter series without a symbol get size 7 diamond, square and
  triangle markers, and pie and doughnut slices take the colour sequence by point,
  on every ring.

`line-markers.pptx` holds two line charts whose series store no marker: one group
has no `c:marker` flag and the other sets it to 0. PowerPoint draws the automatic
markers on both (`microsoft-powerpoint-16.113-line-markers.pdf`).

`OfficeIMO.SharedSource/OpenXml/OfficeOpenXmlChartAutomaticColors.cs` implements
these rules. The export reproduces only the producer's colours; layout, text and
axis scaling are outside this evidence.

Run `python3 OfficeIMO.TestAssets/Documents/Charts/AutomaticColors/generate.py`
from the repository root to recreate the packages. It uses only the Python standard
library and fixed timestamps, so the bytes are reproducible. To refresh the
references, open each package in its Office app, export it as PDF ("Best for
printing"; Word's Save as PDF), and regenerate the fills with `mutool trace`.

| File | SHA-256 |
| --- | --- |
| `automatic-colors.pptx` | `f07aefe0cc3ca06a61d23aec0ceb48fd0d63f6582acc2e634e91a9b9b8d465a5` |
| `automatic-colors.docx` | `dece3dd8f40d160ad0896bddeec01ab0ca3afb7364866a981daa8882553f21cb` |
| `automatic-colors.xlsx` | `2388f11887fa0a5f1bccbff8b0cd29a8c2e156c6d08bda4a93d843465e748cd5` |
| `line-markers.pptx` | `8e3f6a9f8bf44c95e7f274ac41ce9fd76d75c5cfcce145c0ebd6749f631a0691` |

The generator, packages and references use the repository's [MIT license](../../../../LICENSE).
