# OfficeIMO.Reader.Json - JSON reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Json)](https://www.nuget.org/packages/OfficeIMO.Reader.Json)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Json?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Json)

`OfficeIMO.Reader.Json` registers a modular JSON ingestion adapter for `OfficeIMO.Reader`.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Json
```

## Register

```csharp
using OfficeIMO.Reader.Json;

DocumentReaderJsonRegistrationExtensions.RegisterJsonHandler(replaceExisting: true);
```

## What it emits

- AST traversal through `System.Text.Json`.
- Path/type/value rows.
- Chunked structured output with optional Markdown tables.
- Path and stream dispatch.
- Warning chunks for malformed JSON.

## Boundaries

- Reader adapter registration belongs here.
- Shared extraction contracts belong in `OfficeIMO.Reader`.
- `OfficeIMO.Reader.Text` exists only as a compatibility orchestrator for structured text adapters.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
