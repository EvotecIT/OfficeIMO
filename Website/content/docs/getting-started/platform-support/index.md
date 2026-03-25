---
title: Platform Support
description: Supported target frameworks, operating systems, and AOT/trimming compatibility for OfficeIMO packages.
order: 3
---

# Platform Support

OfficeIMO is a fully cross-platform library suite. It does **not** require Microsoft Office to be installed. All document manipulation is performed through the Open XML SDK and pure .NET code.

## Target Frameworks

| Package | .NET 8 | .NET 10 | .NET Standard 2.0 | .NET Framework 4.7.2 |
|---------|--------|---------|--------------------|-----------------------|
| OfficeIMO.Word | Yes | Yes | Yes | Yes (Windows only) |
| OfficeIMO.Excel | Yes | Yes | Yes | Yes (Windows only) |
| OfficeIMO.Markdown | Yes | Yes | Yes | Yes (Windows only) |
| OfficeIMO.CSV | Yes | Yes | Yes | Yes |
| OfficeIMO.Word.Html | Yes | Yes | Yes | Yes (Windows only) |
| OfficeIMO.Word.Markdown | Yes | Yes | Yes | Yes (Windows only) |

The .NET Framework 4.7.2 target is only included when building on Windows. All other targets are fully cross-platform.

.NET Standard 2.0 provides broad compatibility with .NET Core 2.0+, .NET 5+, Mono, Xamarin, and Unity.

## Operating System Support

| OS | Status | Notes |
|----|--------|-------|
| **Windows** | Full support | All target frameworks including .NET Framework 4.7.2 |
| **Linux** | Full support | .NET 8, .NET 10, .NET Standard 2.0 |
| **macOS** | Full support | .NET 8, .NET 10, .NET Standard 2.0 |
| **Docker / CI** | Full support | No native dependencies required |

OfficeIMO has no native OS dependencies. It uses `SixLabors.ImageSharp` for image processing, which is a pure managed library.

## AOT Compilation

OfficeIMO packages enable trimming polyfills (`EnableTrimmingPolyfills`) to prepare for Native AOT scenarios. However, full AOT compatibility depends on the upstream Open XML SDK.

Key considerations:

- **OfficeIMO.Markdown** and **OfficeIMO.CSV** have no external dependencies and are the most AOT-friendly packages.
- **OfficeIMO.Word** and **OfficeIMO.Excel** depend on `DocumentFormat.OpenXml` which uses reflection in some code paths. Test your specific usage with `PublishAot` to verify compatibility.
- See [AOT and Trimming](/docs/advanced/aot-trimming) for detailed guidance.

## IL Trimming

All packages set `GenerateDocumentationFile` and use `LangVersion Latest` with nullable annotations, which helps the trimmer analyze code paths. When publishing trimmed applications:

```xml
<PropertyGroup>
    <PublishTrimmed>true</PublishTrimmed>
    <TrimMode>link</TrimMode>
</PropertyGroup>
```

Test thoroughly, as the Open XML SDK may produce trimming warnings for reflection-based deserialization paths.

## Minimum Visual Studio Version

- Visual Studio 2022 (17.0+) is recommended for full .NET 8 / .NET 10 support.
- Visual Studio 2019 can be used when targeting only .NET Standard 2.0 or .NET Framework 4.7.2.
- JetBrains Rider and VS Code with the C# Dev Kit are also fully supported.

## Architecture Support

OfficeIMO is architecture-independent (AnyCPU). It runs on x64, x86, and ARM64 without modification.
