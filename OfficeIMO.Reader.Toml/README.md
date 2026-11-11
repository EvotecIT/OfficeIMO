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

Rows are emitted in document order, `ChunkRows` per chunk, with plain-text and Markdown table projections and the source lines each chunk covers. Invalid TOML produces a single warning chunk with the error code, line and column instead of throwing.

Files are loaded with `TomlLoadOptions.UntrustedDefaults` unless `TomlReadOptions.LoadOptions` sets other limits; `ReaderOptions.MaxInputBytes` applies as well.
