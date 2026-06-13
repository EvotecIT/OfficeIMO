# OfficeIMO.Reader.Visio - Visio reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Visio)](https://www.nuget.org/packages/OfficeIMO.Reader.Visio)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Visio?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Visio)

`OfficeIMO.Reader.Visio` registers a Visio adapter for `OfficeIMO.Reader` using `OfficeIMO.Visio` inspection snapshots.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Visio
```

## Register

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Visio;

DocumentReaderVisioRegistrationExtensions.RegisterVisioHandler();

IReadOnlyList<ReaderChunk> chunks = DocumentReader
    .Read("diagram.vsdx")
    .ToList();
```

## What it emits

- Page-aware chunks for `.vsdx`, `.vsdm`, `.vstx`, and `.vstm` files.
- Shape Data as `ReaderTable` rows.
- Pages, shapes, connectors, hyperlinks, and optional preview asset metadata in the shared read result envelope.

## Boundaries

- Reader adapter registration belongs here.
- Visio package parsing and inspection belongs in `OfficeIMO.Visio`.
- Shared extraction contracts belong in `OfficeIMO.Reader`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
