# OfficeIMO.MarkdownRenderer - WebView-friendly Markdown rendering

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.MarkdownRenderer)](https://www.nuget.org/packages/OfficeIMO.MarkdownRenderer)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.MarkdownRenderer?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.MarkdownRenderer)

`OfficeIMO.MarkdownRenderer` is the host-oriented rendering package for `OfficeIMO.Markdown`. It builds full HTML shells, renders body fragments, parses renderer-owned Markdown documents, and produces update scripts for WebView2 and browser-based document/chat surfaces.

## Install

```powershell
dotnet add package OfficeIMO.MarkdownRenderer
```

For the first-party IntelligenceX preset pack:

```powershell
dotnet add package OfficeIMO.MarkdownRenderer.IntelligenceX
```

## Quick start

```csharp
using OfficeIMO.MarkdownRenderer;

var options = MarkdownRendererPresets.CreateStrict();

string shellHtml = MarkdownRenderer.BuildShellHtml("Markdown", options);
string bodyHtml = MarkdownRenderer.RenderBodyHtml("""
# Hello

This is rendered through OfficeIMO.MarkdownRenderer.
""", options);
```

## WebView update flow

```csharp
using OfficeIMO.MarkdownRenderer;

var options = MarkdownRendererPresets.CreateStrict();

webView.NavigateToString(MarkdownRenderer.BuildShellHtml("Markdown", options));
await webView.ExecuteScriptAsync(MarkdownRenderer.RenderUpdateScript(markdownText, options));
```

## What it does

- Builds a complete HTML shell for Markdown surfaces.
- Renders HTML body fragments for hosts that own their own shell.
- Produces incremental update scripts and streaming-friendly body output.
- Exposes strict, portable, minimal, relaxed, and transcript-oriented presets.
- Supports AST document transforms, pre-parse normalization, HTML post-processing, and plug-in/feature-pack registration.
- Keeps optional client-side features such as Mermaid, charts, math, Prism, and copy buttons in shell assets rather than managed runtime dependencies.

## Boundaries

- `OfficeIMO.MarkdownRenderer` hosts and renders Markdown output.
- The Markdown model and parser belong in `OfficeIMO.Markdown`.
- WPF/WebView2 control integration belongs in `OfficeIMO.MarkdownRenderer.Wpf`.
- IntelligenceX-specific presets and aliases belong in `OfficeIMO.MarkdownRenderer.IntelligenceX`.
- PDF output belongs in `OfficeIMO.Markdown.Pdf`.

## Related packages

- [OfficeIMO.Markdown](../OfficeIMO.Markdown/README.md)
- [OfficeIMO.Markdown.Html](../OfficeIMO.Markdown.Html/README.md)
- [OfficeIMO.MarkdownRenderer.Wpf](../OfficeIMO.MarkdownRenderer.Wpf/README.md)
- [OfficeIMO.MarkdownRenderer.IntelligenceX](../OfficeIMO.MarkdownRenderer.IntelligenceX/README.md)
- [OfficeIMO.MarkdownRenderer.SamplePlugin](../OfficeIMO.MarkdownRenderer.SamplePlugin/README.md)

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
