---
title: "AOT and Trimming: Making Office Libraries NativeAOT-Ready"
description: "A technical deep dive into how OfficeIMO achieves NativeAOT compatibility and full trimming support across its package family."
date: 2025-11-01
tags: [aot, trimming, performance]
categories: [Deep Dive]
author: "Przemyslaw Klys"
---

.NET 8 made NativeAOT a first-class deployment target. Applications compile to a single native binary with no JIT, no IL, and startup times under 10 milliseconds. But AOT compilation is ruthless: any code path that relies on unconstrained reflection, `Activator.CreateInstance`, or runtime code generation will either fail at compile time or blow up at runtime. This post explains how OfficeIMO achieves AOT and trimming compatibility across its package family.

## Why AOT Matters for Document Libraries

Document generation often runs in serverless functions (AWS Lambda, Azure Functions) or short-lived containers. In those environments, cold-start latency directly affects user experience and cost. A traditional .NET library that takes 800 ms to JIT-compile on first use wastes both time and money. A NativeAOT binary starts in single-digit milliseconds and reaches peak throughput immediately.

## Package Compatibility Matrix

| Package | Trimming Safe | NativeAOT Safe | Notes |
|---|---|---|---|
| OfficeIMO.Word | Yes | Yes | No reflection in hot paths |
| OfficeIMO.Excel | Yes | Yes | Parallel compute uses `Task`-based APIs only |
| OfficeIMO.Markdown | Yes | Yes | Zero dependencies, no reflection at all |
| OfficeIMO.Reader | Yes | Yes | Format detection uses magic bytes, not `Type.GetType` |
| OfficeIMO.Word.Pdf | Partial | Partial | Font discovery uses limited reflection; see below |

## Design Principles We Follow

### 1. No `dynamic` or `ExpandoObject`

Every data structure in OfficeIMO is a concrete type. We never use `dynamic` to paper over weakly-typed XML. This means the trimmer can see every type reference at compile time and preserve exactly what is needed.

### 2. Source Generators over Reflection

Where serialisation is required, for example when reading Open XML attributes into C# objects, we use source-generated mappers:

```csharp
[GeneratedMapper]
public partial class ParagraphProperties
{
    [XmlAttribute("w:val")]
    public string Alignment { get; set; }

    [XmlAttribute("w:sz")]
    public int? FontSize { get; set; }
}
```

The `[GeneratedMapper]` attribute triggers a Roslyn source generator that emits a `Deserialize` method at build time. No `System.Reflection.Emit`, no expression trees.

### 3. Explicit Type Preservation

In the few places where types must be discovered at runtime (plugin loading in OfficeIMO.Reader for custom format handlers), we annotate with `[DynamicallyAccessedMembers]`:

```csharp
public static IFormatReader CreateReader(
    [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicConstructors)]
    Type readerType)
{
    return (IFormatReader)Activator.CreateInstance(readerType)!;
}
```

This tells the trimmer to preserve the constructors of any type passed to this method, preventing `MissingMethodException` at runtime.

### 4. Conditional Font Discovery

OfficeIMO.Word.Pdf needs to enumerate installed fonts, which on .NET traditionally uses `System.Drawing` and reflection-heavy APIs. We isolate this behind a platform-specific facade:

```csharp
#if NET8_0_OR_GREATER
[RequiresUnreferencedCode("Font enumeration uses system APIs")]
#endif
public static class FontResolver
{
    public static string[] GetInstalledFonts() { /* ... */ }
}
```

If you do not call `GetInstalledFonts` and instead supply fonts explicitly via `PdfConverter.Options.FontPaths`, the trimmer removes the reflection-dependent code entirely.

## How to Publish with AOT

```xml
<PropertyGroup>
  <PublishAot>true</PublishAot>
  <TrimMode>full</TrimMode>
</PropertyGroup>
```

```bash
dotnet publish -c Release -r linux-x64
```

The resulting binary for a minimal Word-generation console app is around 12 MB and starts in 8 ms.

## Testing for Trim Correctness

We run the full test suite under `PublishTrimmed` in CI:

```bash
dotnet test -p:PublishTrimmed=true -p:TrimMode=full
```

Any trimmer warning that appears in the build log is treated as a build error. This catches regressions before they reach a release.

## Conclusion

AOT and trimming readiness is not an afterthought in OfficeIMO; it is a design constraint that shapes every API decision. If you are building for serverless, edge, or embedded targets, OfficeIMO gives you Office document capabilities without sacrificing modern deployment patterns.
