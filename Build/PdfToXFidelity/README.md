# PDF-to-X fidelity evaluation

An opt-in scorecard for OfficeIMO's PDF to editable and semantic conversions: DOCX, XLSX, PPTX, RTF, ODT,
ODS, ODP, HTML and Markdown. It is not part of `OfficeIMO.slnx`, CI or any package. Fixed-layout image
exports (PNG, JPEG, TIFF, WebP, SVG) are covered by the renderer suites and are not scored here.

## How the ground truth is built

The ground truth never comes from OfficeIMO:

1. **Round trip through an independent PDF producer.** DOCX, XLSX and PPTX sources with known structure are
   printed to PDF by LibreOffice (headless, in WSL). The PDF is converted back with OfficeIMO and the result is
   compared with the structure read from the original source file.
2. **Second producer.** The `WordPageFootnotes` fixtures also ship PDFs saved by Microsoft Word; those are
   scored against the same sources.
3. **Tagged PDFs.** For tagged PDFs, Poppler's `pdfinfo -struct-text` supplies the logical structure (headings,
   paragraphs, lists, tables, figures) in structure order.

Sources come from `OfficeIMO.TestAssets` (Word-made documents, Apple Pages/Numbers/Keynote exports, the footnote
fixtures) and from `generate_sources.py`, which writes feature-focused files with python-docx, openpyxl and
python-pptx: heading levels, nested bullet and numbered lists, merged and empty cells, two- and three-column
sections, images, hyperlinks, bold/italic/size runs, and numbers, currencies, percentages and dates in en-US,
de-DE, fr-FR, pl-PL, de-CH, ja-JP and ISO styles. Files produced by OfficeIMO are excluded from the corpus.

All readers in `fidelity/` parse the package XML, HTML and Markdown directly with the Python standard library,
so a reader quirk cannot hide or invent a converter difference. OpenDocument and RTF outputs are opened in
LibreOffice and saved as DOCX/XLSX/PPTX first, so they are measured as a LibreOffice user would see them.

## Metrics

All scores are 0–1 and averaged over the documents where the metric applies (a source without tables has no
table score). A failed or unreadable conversion is scored as an empty document.

| Metric | Definition |
|---|---|
| `char_recall`, `char_precision` | Multiset overlap of non-space characters of the word tokens. Header and footer text in the source neither counts as required nor as noise. |
| `word_recall`, `word_precision` | Multiset overlap of word tokens (punctuation and list markers dropped). |
| `reading_order` | Of the recovered words, the share that also appear in source order (longest matching blocks). |
| `heading_recall`, `heading_precision`, `heading_level_acc` | Headings matched by text (similarity ≥ 0.8); level accuracy is the share of matched heading pairs (within 60 positions) whose relative level — higher, same or lower — agrees with the source. |
| `list_recall`, `list_precision`, `list_level_acc`, `list_type_acc` | List items matched by text; nesting depth relative to the shallowest item; ordered versus bulleted. |
| `table_recall`, `table_precision` | Tables paired by cell-text overlap (F1 ≥ 0.3); tables with fewer than two cells are ignored. |
| `teds`, `teds_struct`, `teds_all` | Tree edit similarity over table/row/cell trees, with rows aligned to rows (an upper bound on unconstrained tree edit distance). Cell rename cost is span mismatch, numeric inequality within the displayed precision, or normalised edit distance. `teds_struct` ignores text; `teds_all` counts missed tables as 0. |
| `table_shape_acc`, `merged_cells`, `empty_cells` | Same grid size; non-empty merged cells reproduced with the same span; empty grid positions still empty. |
| `images`, `links` | Share of source pictures present; share of external hyperlink targets present. |
| `footnote_text`, `footnote_native` | Footnote text anywhere in the output; footnotes carried as real notes. |
| `bold_f1`, `italic_f1`, `font_size_acc` | On words aligned between source and output: F1 of the bold and italic flags; sizes within 0.6 pt. |
| `xl_number_value`, `xl_number_typed` | Source numeric cells whose value (within displayed precision) appears in the output; and appears as a typed number. |
| `xl_date_value`, `xl_date_typed`, `xl_text` | The same for dates; text cells reproduced exactly. |
| `visual_layout`, `visual_ssim` | Output re-rendered to PDF by LibreOffice, first 8 pages rasterised by `pdftoppm` at 36 DPI. `visual_layout` is the Dice overlap of ink density on a 22×16 grid per page (blank or missing page 0, same layout 1); `visual_ssim` is mean SSIM, reported with `blank_ssim`, the score of a white page, because SSIM favours empty pages over reflowed text. |

## Compared tools

These run only on the evaluation machine, from an isolated Python virtual environment, and never enter OfficeIMO
projects or packages:

| Tool | Targets | Licence |
|---|---|---|
| [pdf2docx](https://github.com/ArtifexSoftware/pdf2docx) | DOCX | MIT (depends on PyMuPDF, AGPL) |
| [Docling](https://github.com/docling-project/docling) with its layout and table models, OCR off | Markdown, HTML | MIT (models: see the Docling model cards) |
| [PyMuPDF4LLM](https://github.com/pymupdf/pymupdf4llm) | Markdown | AGPL-3.0, benchmark use only |
| LibreOffice PDF import (`writer_pdf_import`, `impress_pdf_import`) | DOCX, PPTX | MPL-2.0 |
| [pdf_oxide](https://github.com/yfedoseev/pdf_oxide) through the `PdfOxide` NuGet package (`ToMarkdownAll`, `ToHtmlAll`) | Markdown, HTML | MIT or Apache-2.0 |

pdf_oxide runs from a small console kept outside the repository; its outputs go to `<work>/outputs/pdfoxide/<id>.md`
and `.html`, and `score` picks them up when that folder exists.

## Running

Requirements: .NET 10 SDK, Python 3.12, WSL with LibreOffice and poppler-utils. Keep all output in a scratch folder.

```powershell
$work = 'D:\Dev\Scratch\pdf-to-x-fidelity'
py -3.12 -m venv "$work\venv"
& "$work\venv\Scripts\python.exe" -m pip install pdf2docx docling pymupdf4llm numpy pillow
$py = "$work\venv\Scripts\python.exe"
dotnet build Build/PdfToXFidelity -c Release
$tool = 'Build/PdfToXFidelity/bin/Release/net10.0/OfficeIMO.PdfToXFidelity.Tool.dll'

& $py -I Build/PdfToXFidelity/run.py corpus --work $work --venv-python $py --tagged <pdf.js>/test/pdfs
& $py -I Build/PdfToXFidelity/run.py officeimo --work $work --tool $tool
& $py -I Build/PdfToXFidelity/run.py tools --work $work --venv-python $py
& $py -I Build/PdfToXFidelity/run.py normalize --work $work
& $py -I Build/PdfToXFidelity/run.py score --work $work
& $py -I Build/PdfToXFidelity/run.py visual --work $work
& $py -I Build/PdfToXFidelity/run.py report --work $work
```

`--tagged` is optional and points at a folder of PDFs; only files with a structure tree and at least 30 words
of tagged text are used. The run writes `results.csv` (one row per document, tool and target), `visual.csv` and
`scorecard.md` into the work folder. Generated corpora and outputs are not committed.

## Limits

- LibreOffice is the PDF producer for most of the corpus, so its layout choices (fonts, line breaks, page
  setup for spreadsheets) are part of every test case; the Word-produced PDFs check that results are not
  specific to one producer.
- The source document is the truth even where the PDF cannot carry it: a style the PDF only shows visually
  (a heading level, an ordered list) has to be inferred by every converter in the same way.
- Text-box positions, colours, borders and fonts other than size are not scored.
