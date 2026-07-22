---
title: Platform Support
description: Supported target frameworks, operating systems, and AOT/trimming compatibility for OfficeIMO packages.
order: 3
---

OfficeIMO is designed for COM-free document automation and does **not** require Microsoft Office to be installed for the workflows covered by this repo. `OfficeIMO.Word.Pdf` and `OfficeIMO.Excel.Pdf` use the first-party `OfficeIMO.Pdf` engine; PDF workloads should still be tested on the target OS with the fonts and templates you plan to ship. The framework matrix below is taken from the current project files in this repo rather than from package-marketing copy.

## Target Frameworks

| Package | .NET 8 | .NET 10 | .NET Standard 2.0 | .NET Framework 4.7.2 |
|---------|--------|---------|--------------------|-----------------------|
| OfficeIMO.Word | Yes | Yes | Yes | Yes (Windows build) |
| OfficeIMO.Excel | Yes | Yes | Yes | Yes (Windows build) |
| OfficeIMO.PowerPoint | Yes | Yes | Yes | Yes (Windows build) |
| OfficeIMO.Markdown | Yes | Yes | Yes | Yes (Windows build) |
| OfficeIMO.CSV | Yes | Yes | Yes | Yes |
| OfficeIMO.Reader | Yes | Yes | Yes | Yes (Windows build) |
| OfficeIMO.Visio | Yes | Yes | Yes | Yes (Windows build) |
| OfficeIMO.Word.Html | Yes | Yes | Yes | Yes (Windows build) |
| OfficeIMO.Word.Markdown | Yes | Yes | Yes | Yes (Windows build) |
| OfficeIMO.Word.Pdf | Yes | Yes | Yes | Yes (Windows build) |
| OfficeIMO.Excel.Pdf | Yes | Yes | Yes | Yes (Windows build) |
| OfficeIMO.Pdf | Yes | Yes | Yes | Yes (Windows build) |

The `.NET Framework 4.7.2` target is included for some packages only when building on Windows. The `netstandard2.0`, `net8.0`, and `net10.0` targets are the main cross-platform story.

## Operating System Support

| OS / Environment | Status | Notes |
|----|--------|-------|
| **Windows** | Supported | Full package set, including Windows-only build targets such as `net472`. |
| **Linux** | Supported | Core packages run cross-platform; PDF conversion should be tested with your font setup. |
| **macOS** | Supported | Core packages run cross-platform; PDF conversion should be tested with your font setup. |
| **Docker / CI** | Supported | Well-suited for COM-free document generation; PDF workloads benefit from explicit font provisioning. |

## Native Dependencies

For the core document packages, OfficeIMO mainly relies on managed libraries such as the Open XML SDK and first-party drawing helpers. The main PDF caveat is layout fidelity: `OfficeIMO.Word.Pdf` and `OfficeIMO.Excel.Pdf` are dependency-light, but host fonts and source templates still affect the rendered result.

## NativeAOT

OfficeIMO supports NativeAOT for its standard in-process document workflows. The repository publishes and executes native applications on `win-x64` and `linux-x64` for Word, typed Excel tables, PowerPoint charts and relationship cloning, Markdown, CSV, the complete 30-handler local Reader adapter preset, and HTML/PDF/image rendering.

All production projects are also built with the .NET trimming and AOT analyzers. Optional cloud authentication, external OCR executables, and WPF/WebView2 keep their own provider or desktop-runtime deployment requirements.

See [NativeAOT Deployment](/docs/advanced/aot-trimming/) for the publish settings, tested workflows, typed-API guidance, and integration boundaries.

## IL Trimming

When publishing trimmed applications, start with a conservative configuration and test your real workflows:

```xml
<PropertyGroup>
    <PublishTrimmed>true</PublishTrimmed>
    <TrimMode>link</TrimMode>
</PropertyGroup>
```

Resolve trimming warnings at the call site. Prefer typed OfficeIMO APIs; preserve members explicitly only when your application intentionally selects runtime types or plug-ins dynamically.

## Minimum Visual Studio Version

- Visual Studio 2022 (17.0+) is recommended for full .NET 8 / .NET 10 support.
- Visual Studio 2019 can still be used for older-target scenarios such as .NET Standard 2.0 or .NET Framework 4.7.2.
- JetBrains Rider and VS Code with the C# Dev Kit are also viable.

## Architecture Support

Most OfficeIMO packages are architecture-neutral from the application perspective, but the repo does not publish one exhaustive architecture-validation matrix for every package/OS combination. Test the exact runtime/OS combination you intend to ship, especially for PDF workloads and trimmed deployments.
