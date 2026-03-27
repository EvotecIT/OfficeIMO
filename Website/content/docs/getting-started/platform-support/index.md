---
title: Platform Support
description: Supported target frameworks, operating systems, and AOT/trimming compatibility for OfficeIMO packages.
order: 3
---

# Platform Support

OfficeIMO is designed for COM-free document automation and does **not** require Microsoft Office to be installed for the workflows covered by this repo. Most packages are pure managed code; one important exception is `OfficeIMO.Word.Pdf`, which adds QuestPDF/SkiaSharp and should be tested on the target OS with the fonts you plan to ship.

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

The `.NET Framework 4.7.2` target is included for some packages only when building on Windows. The `netstandard2.0`, `net8.0`, and `net10.0` targets are the main cross-platform story.

## Operating System Support

| OS / Environment | Status | Notes |
|----|--------|-------|
| **Windows** | Strong support | Full package set, including Windows-only build targets such as `net472`. |
| **Linux** | Strong support | Core packages run cross-platform; PDF conversion should be tested with your font setup. |
| **macOS** | Strong support | Core packages run cross-platform; PDF conversion should be tested with your font setup. |
| **Docker / CI** | Supported | Well-suited for COM-free document generation; PDF workloads benefit from explicit font provisioning. |

## Native Dependencies

For the core document packages, OfficeIMO mainly relies on managed libraries such as the Open XML SDK and ImageSharp. The main caveat is `OfficeIMO.Word.Pdf`, which uses QuestPDF and SkiaSharp, so runtime packaging and host fonts matter more there than they do for the rest of the suite.

## AOT Compilation

OfficeIMO does **not** have one identical AOT story across every package.

- **Best candidates:** `OfficeIMO.Markdown` and `OfficeIMO.CSV`
- **Requires scenario testing:** `OfficeIMO.Word`, `OfficeIMO.Excel`, `OfficeIMO.PowerPoint`, and `OfficeIMO.Reader`
- **Treat separately:** `OfficeIMO.Word.Pdf`

Some projects in the repo enable trimming polyfills, but full NativeAOT success still depends on the dependency graph and the code paths your application exercises.

See [AOT and Trimming](/docs/advanced/aot-trimming) for more detailed guidance.

## IL Trimming

When publishing trimmed applications, start with a conservative configuration and test your real workflows:

```xml
<PropertyGroup>
    <PublishTrimmed>true</PublishTrimmed>
    <TrimMode>link</TrimMode>
</PropertyGroup>
```

If your workload uses Open XML-heavy paths, expect to validate trimming warnings rather than assuming they are harmless.

## Minimum Visual Studio Version

- Visual Studio 2022 (17.0+) is recommended for full .NET 8 / .NET 10 support.
- Visual Studio 2019 can still be used for older-target scenarios such as .NET Standard 2.0 or .NET Framework 4.7.2.
- JetBrains Rider and VS Code with the C# Dev Kit are also viable.

## Architecture Support

Most OfficeIMO packages are architecture-neutral from the application perspective and run on x64, x86, and ARM64. As always, test the exact runtime/OS combination you intend to ship, especially for PDF workloads and trimmed deployments.
