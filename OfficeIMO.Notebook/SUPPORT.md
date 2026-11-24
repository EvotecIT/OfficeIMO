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
| HTML | Writes offline HTML with supported images, semantic tables and MathML. An opt-in profile enables supported controls from saved Jupyter widget state. |
| Markdown | Writes standalone Markdown with supported embedded images and saved outputs. The typed projection is also shared by the document exports. |
| DOCX | Writes editable text, code, headings and supported tables through the Word document engine. Saved figures use the Word adapter's supported image formats. |
| CSV and XLSX | Exports a selected inline structured data resource with an ordered field schema. |

For chart outputs, PDF, HTML, Markdown and Word first try saved SVG, PNG and
JPEG representations. When a bundle contains only a supported chart
specification, export can draw a vector preview.

| Saved chart specification | Static preview |
| --- | --- |
| Plotly Cartesian | Single panel with categorical grouped bars or numeric scatter markers and lines. Supported `scattergl` traces draw vector geometry from the saved data. Numeric arrays can use Plotly.py's one-dimensional typed encoding. |
| Plotly 3D | One scene with supported `scatter3d` traces or a regular `surface` grid. Uses the saved camera, projects depth-ordered vector geometry, and accepts supported Plotly.py typed arrays. |
| Vega-Lite | Single-unit categorical bars or numeric points and lines using inline literal rows, including a named dataset stored with the specification. |
| Other advanced charts | Layers, facets, meshes and maps use a supported saved image when available. The original chart data remains in the notebook on save. |

Saved images and reconstructed previews report the reduction to static output.
Plotly scenes use OfficeIMO's camera and drawing model. They approximate visual
effects such as lighting and styling. Each figure has a bounded data and image
budget. The [advanced chart fixture](../OfficeIMO.TestAssets/Notebook/AdvancedOutputs/Charts/README.md) records the tested camera, surface and WebGL inputs.

The opt-in HTML profile renders supported saved text, selection, toggle, range,
progress, Play and box models. It resolves browser-side widget links when both
endpoints can be rendered. Kernel callbacks and notebook-supplied JavaScript
remain inert. Conversion reports identify unsupported saved views. See the
[original widget fixtures](../OfficeIMO.TestAssets/Notebook/AdvancedOutputs/Widgets/README.md).

The notebook codec owns its bounded JSON storage for the lifetime of the model.
Media are decoded on demand. The parser enforces byte, depth, cell and output
budgets. Use the PDF conversion report to identify source cells and outputs
that need attention, and the save options to control acceptance of native
normalization.

The [producer corpus](../OfficeIMO.TestAssets/Notebook/Producers/README.md) covers
JupyterLab, Colab, VS Code, Kaggle, .NET Interactive, Deepnote, Databricks,
and marimo. Tests check typed source and saved results
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
