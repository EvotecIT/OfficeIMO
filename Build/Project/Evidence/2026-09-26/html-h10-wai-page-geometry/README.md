# WAI print-sheet geometry probe

At clean source `757d44c4f`, the unchanged 14-resource WAI archive (SHA-256 `e412cfda033d9aa02b7abd455ae0fab917c2e44ba0ad0f4bc83f233877a4397d`) replayed offline in 13 PDF operations with no runner failures. The opt-in report now records every output page's measured width and height in PDF points (schema 3), in addition to page counts. All ten OfficeIMO PDFs are byte-identical to the preceding clean replay at `732da240d`; all three OfficeIMO and Chromium print rasters at 96 dpi are pixel-identical to the previously inspected page pairs. Chromium and OfficeIMO zero-margin local-font print remain three pages, while default OfficeIMO print remains four.

| WAI print output | Measured page size, pt | First paragraph line ending |
| --- | ---: | --- |
| Chromium `format: A4` | 595.92 × 842.88 | `need HTML` |
| Chromium explicit 210 × 297 mm | 595.92 × 841.92 | `need HTML` |
| Chromium explicit 209.7 × 297 mm | 594.96 × 841.92 | `need HTML` |
| Chromium explicit 209.5 × 297 mm | 594.00 × 841.92 | `need` |
| OfficeIMO zero-margin A4 | 595.276 × 841.89 | `need` |

The Chromium 151.0.7922.34 probe used the same frozen MHTML, print media and Playwright 1.62.0. Chromium's measured PDF sheet is not OfficeIMO's exact 210 mm A4 sheet, even when the request uses explicit millimeters. That difference matters for a borderline line break, but sheet width alone does not explain the result: Chromium still fits `HTML` at a measured width narrower than OfficeIMO's. Text measurement or the line-break threshold remains a live fidelity question. Changing the global renderer wrap tolerance from this page alone would be unjustified; the other visible line, list and resource differences remain open in the [H10 roadmap](../../../../../Docs/ROADMAP.md#unfamiliar-page-conversion-qualification).

The exact-head report and print page rasters are under `Ignore/HtmlUnknownPageQualification/h10-wai-geometry-clean-757d44c4f/`. The standalone browser probe was task-owned scratch output; its measured sizes and line endings are recorded above, and its binaries and duplicated PDFs were removed after comparison. The report's per-operation page geometry makes subsequent print, screen-media and snapshot comparisons explicit rather than assuming that an `A4` label means identical canvases.
