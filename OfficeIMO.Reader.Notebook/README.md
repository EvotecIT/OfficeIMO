# OfficeIMO.Reader.Notebook

`OfficeIMO.Reader.Notebook` adds bounded Jupyter `.ipynb` ingestion to an isolated `OfficeDocumentReader`.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Notebook;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddNotebookHandler(new ReaderNotebookOptions {
        IncludeOutputs = true,
        MaxOutputCharactersPerCell = 20_000
    })
    .Build();

string markdown = reader.ConvertToMarkdown("analysis.ipynb");
```

Markdown, raw, and code cells are projected in source order. Code is emitted as fenced Markdown using the notebook language when available. Stream, error, `text/markdown`, and `text/plain` outputs can be included; binary display outputs such as images, HTML, and widgets are deliberately not decoded or executed.

The adapter uses the runtime JSON API already present in the Reader graph. It does not run kernels, execute cells, start processes, load models, or contact notebook services. `OfficeIMO.Reader.All` registers this handler.

## Registration

| Item | Value |
|---|---|
| Handler id | `officeimo.reader.notebook` |
| Extensions | `.ipynb` |
| Reader kind | `ReaderInputKind.Json` |
| Default input limit | 64 MiB when `ReaderOptions.MaxInputBytes` is not set |

The notebook must be a JSON object with a `cells` array; otherwise reading fails with `FormatException`. A UTF-8 byte-order mark is accepted, and JSON nesting is limited to depth 128.

## Options

| `ReaderNotebookOptions` | Default | Range | Effect |
|---|---|---|---|
| `IncludeCodeCells` | `true` | | When `false`, code cells and their outputs are skipped. |
| `IncludeOutputs` | `true` | | Projects code-cell outputs under an `#### Output` heading. |
| `MaxCells` | `10000` | 1 to 100,000 | Cells inspected before stopping. |
| `MaxOutputsPerCell` | `100` | 0 to 10,000 | Outputs inspected per code cell. |
| `MaxCellCharacters` | `1000000` | 1 to 16,000,000 | Longer cell source is truncated. |
| `MaxOutputCharactersPerCell` | `100000` | 0 to 16,000,000 | Total output text kept per cell. |

Out-of-range values throw `ArgumentOutOfRangeException` when the handler is registered. `ReaderOptions.MaxChars` (default `8000`) splits a large cell into continuation chunks.

## Output

- **Markdown cells** keep their Markdown source.
- **Code cells** become `### Code cell N`, a fence tagged with the notebook language (from `metadata.kernelspec.language`, then `metadata.language_info.name`), and optional outputs. `text/markdown` outputs are inserted as Markdown; stream, error tracebacks, and `text/plain` outputs are fenced. Fences grow longer when the content itself contains backticks.
- **Raw cells** become `### Raw cell N` with a `text` fence.
- **Chunks:** one per included cell, with `Id` `notebook-cell-0000`, `notebook-cell-0001`, and so on, matching the cell's position in the notebook. `Location.SourceBlockKind` is `notebook-markdown-cell`, `notebook-code-cell` or `notebook-raw-cell`. Split cells continue as `notebook-cell-0003-part-0001` with `ContinuesPreviousChunk = true`. `Text` holds the cell source and output text without Markdown decoration.
- **Document metadata:** `notebook-nbformat`, `notebook-nbformat-minor`, `notebook-language`, `notebook-cell-count` and `notebook-included-cell-count`.

## Limits and diagnostics

| Condition | Result |
|---|---|
| More cells than `MaxCells` | Inspection stops; recoverable `notebook-cell-limit` diagnostic with `maxCells` and `totalCells` attributes. |
| Cell source longer than `MaxCellCharacters` | Source truncated; chunk warning `Notebook cell source was truncated at MaxCellCharacters.` |
| More outputs than `MaxOutputsPerCell` | Remaining outputs dropped; chunk warning. |
| Output text over `MaxOutputCharactersPerCell` | Output truncated; chunk warning. |
| Cell split by `MaxChars` | Chunk warning `Notebook content was split due to MaxChars.` |
| Cell with no projectable content after truncation | Cell omitted; `notebook-content-truncated` diagnostic with the cell index and type. |

## Example output

A Python notebook with one Markdown cell and one code cell that printed `42` produces:

````markdown
# Revenue check

### Code cell 2

```python
print(6 * 7)
```

#### Output

```
42
```
````

The code cell's chunk is `notebook-cell-0001`, and the document metadata reports `Language = python`, `CellCount = 2` and `IncludedCellCount = 2`.

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.Reader.Core`.
