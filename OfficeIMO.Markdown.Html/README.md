# OfficeIMO.Markdown.Html - HTML to Markdown conversion

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Markdown.Html)](https://www.nuget.org/packages/OfficeIMO.Markdown.Html)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Markdown.Html?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Markdown.Html)

`OfficeIMO.Markdown.Html` converts HTML fragments or full HTML documents into Markdown text and `OfficeIMO.Markdown` document models.

## Install

```powershell
dotnet add package OfficeIMO.Markdown.Html
```

## Quick start

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;

string markdown = "<h1>Hello</h1><p>Body</p>".ToMarkdown();
MarkdownDoc document = "<h1>Hello</h1><p>Body</p>".LoadFromHtml();
```

## Options

```csharp
using OfficeIMO.Markdown.Html;

var options = new HtmlToMarkdownOptions {
    BaseUri = new Uri("https://example.com/docs/"),
    UseBodyContentsOnly = true,
    PreserveUnsupportedBlocks = true,
    PreserveUnsupportedInlineHtml = true
};

string markdown = "<p><a href=\"guide/start\">Docs</a></p>".ToMarkdown(options);
```

## What it maps

- Headings, paragraphs, ordered and unordered lists, block quotes, code blocks, horizontal rules, tables, images, figures, details/summary, definition lists, and raw HTML fallback blocks.
- Emphasis, strong emphasis, strike-through, code spans, links, images, line breaks, and selected inline HTML wrappers.
- Relative links and image targets when a base URI is supplied.
- Shared `data-omd-*` visual host elements back into semantic fenced blocks when host hints are registered.
- Custom block and inline converters for host or plug-in packages.

## Profiles

- Use the OfficeIMO profile when the downstream consumer is `OfficeIMO.Markdown` or `OfficeIMO.MarkdownRenderer`.
- Use the portable profile when output should remain friendly to generic Markdown engines.

```csharp
var options = HtmlToMarkdownOptions.CreatePortableProfile();
string portable = "<blockquote><p><strong>Example</strong></p></blockquote>".ToMarkdown(options);
```

## Boundaries

- This package owns HTML ingestion into Markdown.
- It does not render Markdown to a host shell; that belongs in `OfficeIMO.MarkdownRenderer`.
- It does not export to PDF; that belongs in `OfficeIMO.Markdown.Pdf` and `OfficeIMO.Pdf`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
