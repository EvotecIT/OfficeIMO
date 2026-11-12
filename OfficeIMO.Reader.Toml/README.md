# OfficeIMO.Reader.Toml - TOML reader adapter

`OfficeIMO.Reader.Toml` adds TOML ingestion to `OfficeIMO.Reader.Core` through the [OfficeIMO.Toml](../OfficeIMO.Toml/README.md) engine. It reads `pyproject.toml`, `Cargo.toml`, `uv.lock`, Hugo and other TOML 1.1 files into path, type and value rows.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Toml
```

## Configure

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Toml;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddTomlHandler(new TomlReadOptions { ChunkRows = 100 })
    .Build();

foreach (ReaderChunk chunk in reader.Enumerate("pyproject.toml")) {
    Console.WriteLine(chunk.Markdown);
}
```

`AddAllOfficeIMOHandlers()` from `OfficeIMO.Reader.All` registers it too, for `.toml` files.

## Output

Every value becomes one row with three columns:

| Column | Content |
| --- | --- |
| `Path` | The value's path, such as `$.project.dependencies[0]` or `$.package[3].name`. Keys that are not identifiers are quoted: `$.tool.uv.sources["my.pkg"]`. The grammar is the one `TomlDocument.Select` accepts, so a path can be fed back to edit the same value. |
| `Type` | `string`, `integer`, `float`, `boolean`, `datetime`, `datetime-local`, `date-local` or `time-local`; empty tables and arrays appear as `table` (`{}`) and `array` (`[]`). |
| `Value` | The decoded string, or the value as written for other types (`0xFF`, `1979-05-27`). |

Rows are emitted in document order, `ChunkRows` per chunk, with plain-text and Markdown table projections and the source lines each chunk covers. Projected into the neutral `OfficeDocumentModel`, for example for PDF output, a TOML result reports `OfficeDocumentFormat.Toml`.

## Options

| Option | Default | Meaning |
| --- | --- | --- |
| `ChunkRows` | 200 | Rows per emitted chunk. |
| `IncludeMarkdown` | `true` | Adds a Markdown table projection to each chunk. |
| `MaxDepth` | `null` (64) | Maximum nesting of tables, arrays and inline tables, counting each key of a dotted key or header. |
| `MaxNodes` | `null` (200,000) | Maximum tables, arrays and values in the document. |
| `MaxScalarLength` | `null` (1,048,576) | Maximum source length of one string or key in UTF-16 code units, including its quotes (`TomlLoadOptions.MaxStringLength`). |
| `LoadOptions` | `null` | Engine settings used as the base for the limits above, and the place to set `MaxInputBytes`. |

## Limits and diagnostics

Limits start from `LoadOptions`, or from `TomlLoadOptions.UntrustedDefaults` when it is `null`; the defaults in brackets come from there. `MaxDepth`, `MaxNodes` and `MaxScalarLength` replace the matching base value when set, so `new TomlReadOptions { LoadOptions = custom, MaxNodes = 500_000 }` keeps the other values of `custom`. `ReaderOptions.MaxInputBytes` applies as well and wins when it is lower. Values below 1 throw `ArgumentOutOfRangeException`.

Exceeding a limit throws `ReaderResourceLimitException` whose `LimitName` is the `TomlReadOptions` property (`MaxDepth`, `MaxNodes`, `MaxScalarLength`) or `MaxInputBytes`, with the engine's `TomlLimitExceededException`, including its position, as `InnerException`. Raise the limit for trusted larger files. Invalid TOML produces a single warning chunk with the error code, line and column instead of throwing.
