# OfficeIMO.Notebook.Html

Export stored Jupyter notebook content to a standalone HTML document or embeddable fragment. The default export uses embedded styles and images with native MathML for supported equations. Notebook code stays inert.

```csharp
using OfficeIMO.Notebook;
using OfficeIMO.Notebook.Html;

using var notebook = NotebookDocument.Load("analysis.ipynb");
await notebook.SaveAsHtmlAsync("analysis.html");
```

Set `ProjectionOptions` to control visibility and add a title or author. To inspect fidelity findings without writing a file, use `ToHtml(options, out report)` or `ToHtmlAsync`. Set `FailOnLoss` to require successful conversion without reported fidelity loss.

Set `IncludeInteractiveWidgets` to enable supported controls from saved Jupyter
widget state. The exported document works offline with native form controls and
a built-in script for supported client-side links. Sliders, selections, text
fields, toggles and Play controls can change values locally. Python callbacks
require a live kernel. The export reports unsupported models and styles and
leaves notebook-supplied widget code inert.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Notebook.Html` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
