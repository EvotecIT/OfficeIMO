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
