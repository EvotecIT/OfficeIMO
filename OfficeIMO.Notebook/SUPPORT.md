# Notebook format support

OfficeIMO opens saved Jupyter notebooks from a path, a caller-owned stream or
UTF-8 memory. The model retains cells, source, metadata, attachments and all
saved MIME alternatives. It is suitable for offline inspection and editing.

| Input or operation | Supported behavior |
| --- | --- |
| nbformat 4.x read | Markdown, code, raw and opaque extension cells. String and string-array fields are joined for text access. Original JSON remains accessible for saved MIME items. |
| nbformat 4.x save | Writes at least 4.5. Keeps future v4 minor versions, unknown properties and payloads. Retains existing valid cell IDs and creates deterministic IDs for missing or invalid entries. Structural preservation permits changes to whitespace, serialization shape and property order. |
| nbformat 3 read | Projects worksheet cells in order into the typed model. Headings, inputs, execution counts and saved outputs are normalized. A v4 save requires explicit acceptance of the lost worksheet structure. |
| Reader | Projects text and Markdown using the existing bounded Reader contract. Use the notebook model or the conversion packages for rich saved output. |
| PDF | Projects ordered notebook content and supported saved text, Markdown, HTML tables, images, attachments and TeX math into fixed pages. Code wraps, tables repeat available headers and oversized figures scale to the page. A figure stays with its caption when both fit. Rendering reports unsupported MIME values, blocked resources and fidelity changes. |
| HTML | Writes offline HTML with supported images, semantic tables and MathML. Stored script and widget payloads are treated as opaque data. |
| Markdown | Writes standalone Markdown with supported embedded images and saved outputs. The typed projection is also shared by the document exports. |
| DOCX | Writes editable text, code, headings and supported tables through the Word document engine. Saved figures use the Word adapter's supported image formats. |
| CSV and XLSX | Exports a selected inline structured data resource with an ordered field schema. |

The notebook codec owns its bounded JSON storage for the lifetime of the model.
Media are decoded on demand. The parser enforces byte, depth, cell and output
budgets. Use the PDF conversion report to identify source cells and outputs
that need attention, and the save options to control acceptance of native
normalization.

The [producer corpus](../OfficeIMO.TestAssets/Notebook/Producers/README.md) covers
JupyterLab, Colab, VS Code, a Kaggle compatibility fixture, .NET Interactive,
Deepnote, Databricks, and marimo. Tests check typed source and saved results
and verify structural preservation of metadata and outputs after an edit.
When a Deepnote Jupyter export duplicates its block group as a cell property
and in metadata, save keeps the metadata copy to satisfy the v4.5 cell schema.
Use [OfficeIMO.Notebook.Import](../OfficeIMO.Notebook.Import/README.md) for
script, publishing, and vendor-native notebook inputs.

Examples and target-specific controls live in [PDF](../OfficeIMO.Notebook.Pdf/README.md),
[HTML](../OfficeIMO.Notebook.Html/README.md), [Markdown](../OfficeIMO.Notebook.Markdown/README.md),
[Word](../OfficeIMO.Notebook.Word/README.md) and
[Spreadsheet](../OfficeIMO.Notebook.Spreadsheet/README.md).
Use the [notebook benchmark suite](../OfficeIMO.Notebook.Benchmarks/README.md)
to rerun the independent fixture checks and matched performance cases.
