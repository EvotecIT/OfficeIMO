# OfficeIMO.MarkdownRenderer.IntelligenceX - IntelligenceX renderer presets

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.MarkdownRenderer.IntelligenceX)](https://www.nuget.org/packages/OfficeIMO.MarkdownRenderer.IntelligenceX)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.MarkdownRenderer.IntelligenceX?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.MarkdownRenderer.IntelligenceX)

`OfficeIMO.MarkdownRenderer.IntelligenceX` is the first-party IntelligenceX feature pack for `OfficeIMO.MarkdownRenderer`. It keeps IntelligenceX transcript presets, visual aliases, and compatibility transforms outside the generic renderer package.

## Install

```powershell
dotnet add package OfficeIMO.MarkdownRenderer.IntelligenceX
```

## Quick start

```csharp
using OfficeIMO.MarkdownRenderer;
using OfficeIMO.MarkdownRenderer.IntelligenceX;

var options = IntelligenceXMarkdownRenderer.CreateTranscriptDesktopShell();
string html = MarkdownRenderer.RenderBodyHtml(markdownText, options);
```

## Examples

### Render transcript Markdown in the desktop shell profile

````csharp
using OfficeIMO.MarkdownRenderer;
using OfficeIMO.MarkdownRenderer.IntelligenceX;

string transcript = """
# Incident timeline

```ix-chart title="Events"
{"type":"bar","labels":["09:00","09:15"],"values":[3,8]}
```
""";

var options = IntelligenceXMarkdownRenderer.CreateTranscriptDesktopShell();
string shell = MarkdownRenderer.BuildShellHtml("Investigation", options);
string update = MarkdownRenderer.RenderUpdateScript(transcript, options);
````

### Parse with IntelligenceX compatibility transforms

```csharp
using OfficeIMO.MarkdownRenderer;
using OfficeIMO.MarkdownRenderer.IntelligenceX;

var options = IntelligenceXMarkdownRenderer.CreateTranscriptDesktopShell();
MarkdownRendererParseResult result = MarkdownRenderer.ParseDocumentResult(transcript, options);

foreach (var diagnostic in result.TransformDiagnostics) {
    Console.WriteLine($"{diagnostic.Source}: {diagnostic.TransformName}");
}
```

## What it adds

- IntelligenceX transcript presets and desktop-shell defaults.
- Transcript-oriented semantic visual aliases.
- Compatibility transforms for known IntelligenceX transcript shapes.
- Shared registration hooks for Markdown rendering and HTML round-trip flows.

## Boundaries

- Generic Markdown rendering stays in `OfficeIMO.MarkdownRenderer`.
- IntelligenceX-specific transcript behavior belongs here.
- Host application UI and storage behavior should stay in the IntelligenceX app, not this package.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
