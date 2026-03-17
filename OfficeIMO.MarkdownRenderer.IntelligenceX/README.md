# OfficeIMO.MarkdownRenderer.IntelligenceX

`OfficeIMO.MarkdownRenderer.IntelligenceX` is the first-party IntelligenceX plugin pack for `OfficeIMO.MarkdownRenderer`.

It keeps the generic OfficeIMO markdown renderer neutral while exposing the IntelligenceX visual aliases and transcript-oriented preset helpers from a dedicated package boundary.

## What It Adds

- `ix-chart`, `ix-network`, and `ix-dataview` visual aliases
- transcript compatibility wiring for IntelligenceX-hosted markdown/chat output
- convenience helpers that build on top of `OfficeIMO.MarkdownRenderer`

## Example

```csharp
using OfficeIMO.MarkdownRenderer;
using OfficeIMO.MarkdownRenderer.IntelligenceX;

var options = IntelligenceXMarkdownRenderer.CreateTranscriptDesktopShell();
string html = MarkdownRenderer.RenderBodyHtml(markdown, options);
```

## Compatibility Pack

```csharp
using OfficeIMO.MarkdownRenderer;
using OfficeIMO.MarkdownRenderer.IntelligenceX;

var options = MarkdownRendererPresets.CreateStrict();
options.ApplyFeaturePack(IntelligenceXMarkdownRenderer.TranscriptCompatibilityPack);
```

Use the transcript compatibility pack when you want to stay generic-first but still layer in the IntelligenceX visual aliases, transcript reader/AST contract, and legacy transcript cleanup behavior as one reusable host-level contract.

## Transcript Plugin

```csharp
using OfficeIMO.MarkdownRenderer;
using OfficeIMO.MarkdownRenderer.IntelligenceX;

var options = MarkdownRendererPresets.CreateStrict();
IntelligenceXMarkdownRenderer.ApplyTranscriptContract(options);
```

Use the transcript plugin when you want IX aliases, IX schema support, and the IX transcript reader/AST contract, but you do not want the broader compatibility-pack cleanup layer.

## Typed IX Fence Options

```csharp
using OfficeIMO.MarkdownRenderer.IntelligenceX;

var options = IntelligenceXMarkdownRenderer.ParseVisualFenceOptions(
    "ix-chart {#quarterly-summary .wide title=\"Quarterly Revenue\" pinned theme=\"amber\" maxItems=12}");
```

Use `ParseVisualFenceOptions(...)` when an IntelligenceX host/plugin wants a typed view of common IX fence metadata such as `Pinned`, `Theme`, `View`, `Variant`, `MaxItems`, `Title`, `ElementId`, and `Classes` while still building on the shared `OfficeIMO.Markdown` fence AST contract.

The package now carries that IX schema directly on `VisualsPlugin`, `TranscriptPlugin`, and `TranscriptCompatibilityPack`, so `ApplyVisuals(...)`, `ApplyTranscriptContract(...)`, `ApplyTranscriptCompatibility(...)`, and the `CreateTranscript*` helpers all register IX renderer aliases and IX option-schema validation, while the transcript-aware entrypoints also carry the IX transcript reader/AST contract as one package-level contract.
