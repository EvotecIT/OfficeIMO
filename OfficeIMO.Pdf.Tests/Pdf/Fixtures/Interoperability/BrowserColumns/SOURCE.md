# Browser-balanced prose columns

The adjacent HTML files are synthetic inputs printed by headless Microsoft Edge
on macOS with `--headless=new --no-pdf-header-footer --print-to-pdf`. They use
Arial at 10 or 11 pt. The two-column case has a 4 mm gutter; the three-column cases have
4 mm and 12 mm gutters. Each contains three complete paragraphs per column, in ascending
`ParaN` order.

The PDFs are independent-producer fixtures. The first column starts one line
below the others because the browser keeps its first paragraph margin. The
two-column case protects narrow-gutter segmentation; the three-column case
protects wrapped prose from false table detection.
The narrow three-column case protects empty gutters from being replaced with
wider strips that cut through ragged line ends.
