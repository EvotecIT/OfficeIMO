# OfficeIMO.MarkdownRenderer.SamplePlugin - renderer plug-in example

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.MarkdownRenderer.SamplePlugin)](https://www.nuget.org/packages/OfficeIMO.MarkdownRenderer.SamplePlugin)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.MarkdownRenderer.SamplePlugin?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.MarkdownRenderer.SamplePlugin)

`OfficeIMO.MarkdownRenderer.SamplePlugin` is a sample plug-in package for `OfficeIMO.MarkdownRenderer`. It shows how a third-party-style package can register visual features, HTML round-trip hints, and renderer options without changing the renderer core.

## Install

```powershell
dotnet add package OfficeIMO.MarkdownRenderer.SamplePlugin
```

## Use it from a renderer host

```csharp
using OfficeIMO.MarkdownRenderer;
using OfficeIMO.MarkdownRenderer.SamplePlugin;

var options = MarkdownRendererPresets.CreateStrict();
options.ApplyFeaturePack(SampleMarkdownRenderer.StatusPanelFeaturePack);

string html = MarkdownRenderer.RenderBodyHtml(markdownText, options);
```

## What it demonstrates

- Keeping host or product-specific visuals in a plug-in package.
- Registering renderer assets and Markdown document transforms.
- Carrying matching HTML round-trip hints for `OfficeIMO.Markdown.Html`.
- Preserving the generic renderer boundary.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
