# Reverse-conversion fixtures

Small PDFs used by the PDF to Word, Excel, PowerPoint and Markdown regression tests.

- `article-two-column.pdf` and `invoices.pdf` were exported by LibreOffice 26.8 from flat OpenDocument
  sources written for these tests. The article is a two-column page with a spanning title, a figure and a
  bulleted list; the invoice list is a ruled table with ISO dates and comma-decimal amounts. Both embed
  subsets of the Liberation fonts (SIL Open Font License).
- `Pdf.Showcase.Statement.pdf` and `Pdf.Showcase.Dashboard.pdf` are produced by the OfficeIMO.Pdf showcase
  examples. The statement uses `31,80 PLN` style amounts; the dashboard has filled table headers with white
  text and large KPI values.

The text is synthetic. The files contain no scripts, forms or external references.
