# Ruled tables and vector figures

`word-business-summary.pdf` is the frozen OfficeIMO PDF export of the repository's business-delivery-summary Word fixture at `c000d5ce86`. Its second page has a ruled Worksheet/Purpose table that was mistaken for a vector figure, hiding its editable text behind a picture. The manifest records its exact bytes. The tests also use the existing independent Keynote table export under `OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports`.

These are test-only inputs. The tests qualify figure classification and retained editable text, separately from complete ruled-grid reconstruction.

`word-business-grids-rotated.pdf` repeats source page 2 at page rotations 0, 90, 180 and 270 degrees. It was produced with qpdf page extraction followed by `--rotate=+90:2 --rotate=+180:3 --rotate=+270:4`; text, embedded fonts and grid paint are unchanged. These pages qualify wrapped row values, table reading order and the reading-frame coordinate transform. The original second-page fixture also qualifies distinct labels beside empty response cells.

`word-business-grids-fragmented.pdf` repeats source page 2 with four authored grid variants: adjacent vertical segments, vertical segments with a 1 pt gap, adjacent horizontal segments, and horizontal segments with a 1 pt gap. Near-axis segments differ by 0.1 pt. Only the Worksheet/Purpose stroke operators were edited in uncompressed QDF; source text, fonts, clipping and the lower table are unchanged. `fix-qdf` repaired stream lengths and offsets, then qpdf combined the four pages. This qualifies near-axis interval merging without bridging disconnected rules.
