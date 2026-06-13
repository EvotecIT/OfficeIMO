# OfficeIMO.Reader.Csv - CSV reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Csv)](https://www.nuget.org/packages/OfficeIMO.Reader.Csv)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Csv?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Csv)

`OfficeIMO.Reader.Csv` registers a modular CSV/TSV ingestion adapter for `OfficeIMO.Reader`.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Csv
```

## Register

```csharp
using OfficeIMO.Reader.Csv;

DocumentReaderCsvRegistrationExtensions.RegisterCsvHandler(replaceExisting: true);
```

## What it emits

- CSV/TSV chunks with table-aware output.
- Path and stream dispatch.
- Deterministic chunk IDs and row-based locations.
- `MaxInputBytes` enforcement through shared `ReaderInputLimits`.

## Boundaries

- Reader adapter registration belongs here.
- CSV parsing and document modeling belongs in `OfficeIMO.CSV`.
- Shared extraction contracts belong in `OfficeIMO.Reader`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
