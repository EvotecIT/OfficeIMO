# OfficeIMO.Reader.Epub - EPUB reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Epub)](https://www.nuget.org/packages/OfficeIMO.Reader.Epub)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Epub?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Epub)

`OfficeIMO.Reader.Epub` bridges `OfficeIMO.Epub` output into `OfficeIMO.Reader` chunk contracts.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Epub
```

## Register

```csharp
using OfficeIMO.Reader.Epub;

DocumentReaderEpubRegistrationExtensions.RegisterEpubHandler();
```

## What it emits

- Chapter-to-chunk projection.
- Max-character chunk splitting.
- Markdown and text chunk payloads.
- Warning chunks propagated from EPUB parser warnings.
- Virtual source paths such as `.epub::chapter.xhtml` for traceability.
- Path and stream dispatch, including non-seekable stream support.

## Boundaries

- Reader adapter registration belongs here.
- EPUB parsing belongs in `OfficeIMO.Epub`.
- Shared extraction contracts belong in `OfficeIMO.Reader`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
