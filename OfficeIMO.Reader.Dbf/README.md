# OfficeIMO.Reader.Dbf

`OfficeIMO.Reader.Dbf` reads bounded DBF/xBase tables into OfficeIMO Reader chunks.
The [DbaClientX DBF codec](https://github.com/EvotecIT/DbaClientX/tree/master/DbaClientX.Dbf)
owns the file format, code pages, typed values and DBT/FPT memo handling. This
adapter projects those values into Reader tables without adding a database driver.

## Read a table

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Dbf;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddDbfHandler(new ReaderDbfOptions {
        ChunkRows = 100,
        AllowMemoSidecarReads = true
    })
    .Build();

foreach (ReaderChunk chunk in reader.Read("contacts.dbf")) {
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

`OfficeIMO.Reader.All` includes this handler. Configure it through
`ReaderAllOptions.Dbf`. Registration takes a copy of the options; subsequent
changes to the supplied object do not affect an existing reader.

## Input and output contracts

- Supported profiles are dBASE III-compatible `03`/`83`, FoxPro 2 `F5`, and the
  bounded Visual FoxPro `30`/`31`/`32` field profiles described by the codec.
  Unsupported dialects, field types, encryption and malformed records fail
  explicitly. This is structured reading, not text salvage.
- Path reads use the table alone by default. `AllowMemoSidecarReads` enables the
  codec's same-stem `.dbt` or `.fpt` sidecar lookup. A nonempty memo reference
  without its sidecar fails when the value is read.
- Stream reads never resolve a filename against the filesystem, including when
  sidecar reads are enabled. The caller's stream remains open. For explicit
  table and memo streams, use `DbfDataReader` directly.
- `ReadOptions` controls byte, record, field and memo budgets, deleted-record
  inclusion and an optional encoding override. Reader's `MaxInputBytes` can
  tighten the table byte budget. `MaxTableRows` limits rows per chunk.
- Columns retain their source names. Numbers use invariant text, dates use ISO
  dates, timestamps retain milliseconds, and binary data becomes Base64. Reader
  cells represent nulls as empty strings and emit a warning; the typed reader
  retains `DBNull`. Markdown escapes source text as literal cell content.
- `SourceBlockIndex` identifies the first physical record in a chunk, counting
  deleted records. DBF indexes, database backlinks and embedded objects remain
  inert. No DBF writing or database traversal is provided.
- A table-only byte hash cannot identify memo-dependent content, so the adapter
  leaves `SourceHash` unset and reports that limitation. Reader chunk hashes
  identify the emitted projection.

## Convert typed rows to CSV or XLSX

Use the codec's standard `DbDataReader` with the existing exporters. CSV and Excel
remain optional packages; this Reader adapter does not reference either exporter.

```csharp
using DBAClientX.Dbf;
using OfficeIMO.CSV;
using OfficeIMO.Excel;

using (var rows = DbfDataReader.Open("contacts.dbf"))
using (var output = File.CreateText("contacts.csv")) {
    CsvDocument.WriteDataReader(output, rows, new CsvSaveOptions {
        DateTimeFormat = "O",
        NullValue = "<null>"
    });
}

using (var rows = DbfDataReader.Open("contacts.dbf"))
using (var output = File.Create("contacts.xlsx")) {
    ExcelDocument.WriteDataReader(output, rows,
        new ExcelTabularWriteOptions { RequireStreaming = true });
}
```

The codec's path overload owns its input streams and resolves supported memo
sidecars. Each exporter consumes the reader without closing it. Binary values
become Base64 text in both destinations. CSV does not retain native column types
or a distinct null value unless a null marker is selected. XLSX follows Excel's
numeric precision and cell text length limits; large memo values can exceed the
32,767-character cell limit. These conversions do not preserve DBF indexes,
deletion flags, code-page metadata or save-back semantics.

## Qualification and dependencies

The focused tests use independently generated Python `dbf` tables and `dbfread`
decoding: plain dBASE III, dBASE III with DBT, FoxPro 2 with FPT, and Visual FoxPro
with FPT, nulls, typed numbers, timestamps and binary values. Tests reopen CSV and
XLSX output, validate the Open XML package, and exercise stream ownership,
cancellation, registration snapshots, chunking and memo access policy. See the
[fixture provenance](../OfficeIMO.Reader.Dbf.Tests/Fixtures/README.md).

This evidence does not establish native dBASE/FoxPro application acceptance,
Visual FoxPro `31`/`32` generation coverage, Windows .NET Framework runtime
acceptance or NativeAOT qualification.

Runtime dependencies are `OfficeIMO.Reader.Core` and `DBAClientX.Dbf`. Older
target frameworks use the codec's Microsoft code-page compatibility package.
Python tools are fixture-generation tools only and are absent from shipped
packages and normal tests. Targets are `netstandard2.0`, `net8.0`, `net10.0`, and
`net472` on Windows. License: MIT.
