# Browser-balanced prose columns

The adjacent HTML files are synthetic inputs printed by headless Microsoft Edge
on macOS with `--headless=new --no-pdf-header-footer --print-to-pdf`. They use
Arial at 11 pt. The two-column case has a 4 mm gutter; the three-column case has
a 12 mm gutter. Both contain three complete paragraphs per column, in ascending
`ParaN` order.

The PDFs are independent-producer fixtures. The first column starts one line
below the others because the browser keeps its first paragraph margin. The
two-column case protects narrow-gutter segmentation; the three-column case
protects wrapped prose from false table detection.
