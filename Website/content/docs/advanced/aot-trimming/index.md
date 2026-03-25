---
title: AOT and Trimming
description: Guidance for ahead-of-time compilation and IL trimming with OfficeIMO packages.
order: 80
---

# AOT Compilation and Trimming

OfficeIMO packages are designed with modern .NET deployment scenarios in mind, including Native AOT compilation and IL trimming. This guide covers the current state of compatibility, configuration recommendations, and known limitations.

## Overview

| Package | Trimming Ready | AOT Ready | Notes |
|---------|---------------|-----------|-------|
| OfficeIMO.Markdown | Yes | Yes | Small dependency surface and the strongest AOT fit in the repo |
| OfficeIMO.CSV | Yes | Yes | Lightweight package and a strong fit for trimmed deployments |
| OfficeIMO.Word | Partial | Partial | Depends on Open XML SDK reflection patterns |
| OfficeIMO.Excel | Partial | Partial | Depends on Open XML SDK reflection patterns |
| OfficeIMO.Word.Html | Partial | Partial | AngleSharp uses reflection for CSS parsing |
| OfficeIMO.Word.Markdown | Partial | Partial | Inherits limitations from Word and Html packages |

## Trimming Polyfills

Some OfficeIMO projects enable trimming polyfills in their project files:

```xml
<PropertyGroup>
    <EnableTrimmingPolyfills>true</EnableTrimmingPolyfills>
</PropertyGroup>
```

That helps where those compatibility shims are used, but it should not be read as a blanket guarantee that every package is trimming-safe in every scenario.

## Configuring Trimmed Publishing

To publish a trimmed application that uses OfficeIMO:

```xml
<PropertyGroup>
    <PublishTrimmed>true</PublishTrimmed>
    <TrimMode>link</TrimMode>
</PropertyGroup>
```

### Handling Trimming Warnings

When trimming an application that uses `OfficeIMO.Word` or `OfficeIMO.Excel`, you may see warnings from the Open XML SDK related to:

- `DocumentFormat.OpenXml.Framework.SchemaAttrAttribute`
- XML deserialization reflection paths
- Generic type instantiation in part readers

To suppress known-safe warnings, add trim roots for the types your application actually uses:

```xml
<ItemGroup>
    <TrimmerRootAssembly Include="DocumentFormat.OpenXml" />
</ItemGroup>
```

Alternatively, mark specific types as preserved:

```xml
<ItemGroup>
    <TrimmerRootDescriptor Include="TrimmerRoots.xml" />
</ItemGroup>
```

With a `TrimmerRoots.xml` file:

```xml
<linker>
    <assembly fullname="DocumentFormat.OpenXml">
        <type fullname="DocumentFormat.OpenXml.Wordprocessing.*" preserve="all" />
        <type fullname="DocumentFormat.OpenXml.Spreadsheet.*" preserve="all" />
    </assembly>
</linker>
```

## Native AOT Compilation

To publish with Native AOT:

```xml
<PropertyGroup>
    <PublishAot>true</PublishAot>
</PropertyGroup>
```

### AOT-Safe Packages

**OfficeIMO.Markdown** and **OfficeIMO.CSV** are the best starting points for AOT-sensitive workloads because they:

- Have zero external dependencies
- Use no reflection
- Use no dynamic code generation
- Use nullable annotations and latest C# features for analyzer compatibility

Example Markdown generator for AOT experiments:

```csharp
// Program.cs
using OfficeIMO.Markdown;

var doc = MarkdownDoc.Create()
    .H1("AOT Report")
    .P("Generated from a native AOT binary.")
    .Table(t => t
        .Headers("Metric", "Value")
        .Row("Startup", "< 10ms")
        .Row("Binary Size", "~8 MB")
    );

File.WriteAllText("report.md", doc.ToString());
```

### AOT Limitations for Word/Excel

The Open XML SDK (`DocumentFormat.OpenXml`) uses reflection-based deserialization internally. As of version 3.x, the SDK team has been making progress toward AOT compatibility, but full support is not guaranteed for all code paths.

Known issues:

1. **Part deserialization** -- The SDK uses `Activator.CreateInstance` to instantiate some OpenXml element types during document loading.
2. **Schema validation** -- The `OpenXmlValidator` class uses reflection to discover element schemas.
3. **Generic part readers** -- Some typed readers use generic instantiation patterns that the AOT compiler may not discover.

**Recommendation**: Test your specific usage scenario thoroughly with AOT. Document creation (write-only) paths tend to work better than read/modify paths.

## Self-Contained Single-File Deployment

For a single-file deployment without full AOT:

```xml
<PropertyGroup>
    <PublishSingleFile>true</PublishSingleFile>
    <SelfContained>true</SelfContained>
    <RuntimeIdentifier>win-x64</RuntimeIdentifier>
    <IncludeNativeLibrariesForSelfExtract>true</IncludeNativeLibrariesForSelfExtract>
</PropertyGroup>
```

This approach is often easier than full AOT, but you should still validate the exact package combination and code paths your application uses.

## Best Practices

1. **Prefer OfficeIMO.Markdown and OfficeIMO.CSV** for AOT scenarios -- they are the lowest-risk packages in the repo for that style of deployment.
2. **Test early** -- If you plan to publish with AOT or trimming, test from the start of your project.
3. **Pin Open XML SDK version** -- OfficeIMO pins to `[3.3.0, 4.0.0)`. Stay within this range for tested compatibility.
4. **Use `TrimMode=link`** -- This is more aggressive than `copyused` but gives smaller binaries. Test thoroughly.
5. **Check for trimming warnings** -- Build with `<TrimmerSingleWarn>false</TrimmerSingleWarn>` to see per-assembly warnings.
6. **Consider ReadyToRun** as an alternative -- For faster startup without full AOT, use `<PublishReadyToRun>true</PublishReadyToRun>`.

## Future Outlook

As the Open XML SDK improves its AOT annotations (tracking issue: [dotnet/Open-XML-SDK#1424](https://github.com/dotnet/Open-XML-SDK/issues/1424)), the Open XML-based OfficeIMO packages should benefit as well. Until then, treat trimming and AOT as scenario-driven validation work rather than a blanket promise.
