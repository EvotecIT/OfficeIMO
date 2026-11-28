# OfficeIMO.Notebook

Open and write Jupyter notebooks with a managed document model. Cells, metadata,
attachments, and saved output MIME alternatives remain available without running
a kernel.

## Install

```sh
dotnet add package OfficeIMO.Notebook
```

```csharp
using OfficeIMO.Notebook;

using NotebookDocument notebook = NotebookDocument.Create();
notebook.AddMarkdownCell("# Lab results");
NotebookCell cell = notebook.AddCodeCell("print(42)");
cell.ExecutionCount = 1;
cell.AddOutput(NotebookOutput.CreateStream("stdout", "42\n"));
notebook.Save("lab.ipynb");

using NotebookDocument loaded = NotebookDocument.Load("lab.ipynb");
foreach (NotebookCell item in loaded.Cells) {
    Console.WriteLine(item.Source);
}
```

`Load` accepts paths, streams and UTF-8 memory. `LoadAsync` accepts paths and
streams. Caller-supplied streams remain open. For atomic file replacement, use a
path overload. Stream saves start at the current write position and can leave
partial output if interrupted. The notebook owns the parsed JSON
backing store, so attached `JsonElement` values are available until disposal.
Payloads are retained in their original JSON representation and decoded when
requested by the caller. The parser reads the whole bounded JSON document into
memory. Configure `NotebookLoadOptions` to set the input, nesting, cell and
output-count limits.

Saving produces nbformat 4.5 or retains a higher v4 minor version. Unknown v4
properties, cell and output types and MIME alternatives are carried forward.
This is structural preservation. The writer can change formatting and property
order, gives older cells IDs, and repairs missing, invalid or duplicate IDs.
Legacy v3 worksheets load as one ordered cell list. Saving their normalized form
requires `NotebookSaveOptions.AllowLoss`, since v4 has no worksheet boundaries.
`NotebookDocument.Diagnostics` describes these normalization findings.

Cell edits use zero-based indexes in current document order. Use
`MoveCell(fromIndex, toIndex)` with the destination's final index and
`RemoveCell(index)` to remove a cell. Moving a cell preserves its ID, metadata,
attachments, and saved outputs. Use `AddCodeCell(source, language)` to set a
per-cell language. `ClearOutputs()` removes saved outputs from a code cell.
Setting `ExecutionCount` to `null` clears the count independently.
`FilePath` identifies the default native save path when one exists.

The package handles stored notebook content offline. It does not invoke a
kernel, load widgets or access remote resources.

Use the [support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Notebook/SUPPORT.md) to choose a native lifecycle or export
route and check its fidelity limits.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 1 | 0 | 0 | 0 | 0 | 0 |
| Read | 1 | 0 | 0 | 0 | 0 | 0 |
| Edit | 1 | 0 | 0 | 0 | 0 | 0 |
| Preserve | 0 | 1 | 0 | 0 | 0 | 0 |
| Inspect | 1 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Notebook` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
