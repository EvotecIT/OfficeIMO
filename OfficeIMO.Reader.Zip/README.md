# OfficeIMO.Reader.Zip - ZIP reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Zip)](https://www.nuget.org/packages/OfficeIMO.Reader.Zip)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Zip?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Zip)

`OfficeIMO.Reader.Zip` registers a ZIP traversal adapter for `OfficeIMO.Reader`.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Zip
```

## Register

```csharp
using OfficeIMO.Reader.Zip;

DocumentReaderZipRegistrationExtensions.RegisterZipHandler();
```

## What it emits

- Safe entry enumeration through `OfficeIMO.Zip`.
- Best-effort entry extraction into `ReaderChunk`.
- Warning chunks for skipped or failed entries.
- Bounded nested ZIP traversal with `ReaderZipOptions`.
- Path and stream dispatch, including non-seekable stream support.

## Boundaries

- Reader adapter registration belongs here.
- ZIP traversal policy belongs in `OfficeIMO.Zip`.
- Shared extraction contracts belong in `OfficeIMO.Reader`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
