# OfficeIMO.Markdown — .NET Markdown Builder, Reader, and HTML Renderer

OfficeIMO.Markdown is a cross-platform .NET library for composing Markdown, parsing it back into a typed document model, and rendering it to HTML without external runtime dependencies.

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Markdown)](https://www.nuget.org/packages/OfficeIMO.Markdown)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Markdown?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Markdown)

- Targets: netstandard2.0, net472, net8.0, net10.0
- License: MIT
- NuGet: `OfficeIMO.Markdown`
- Dependencies: none

### AOT / Trimming Notes

- The core builder, reader, and renderer APIs avoid external runtime dependencies and are straightforward to trim.
- Reflection-based helpers remain available for convenience scenarios such as `FromAny(...)`.
- For NativeAOT or trimming-sensitive workloads, prefer typed overloads and explicit selectors.

```csharp
using OfficeIMO.Markdown;

var people = new[] {
    new Person("Alice", "Dev", 98.5),
    new Person("Bob", "Ops", 91.0)
};

var doc = MarkdownDoc.Create()
    .Table(t => t.FromSequenceAuto(people));

doc.FrontMatter(FrontMatterBlock.FromObject(new BuildInfo { Version = "1.0.0" }));

public sealed record Person(string Name, string Role, double Score);
public sealed class BuildInfo {
    public string Version { get; set; } = "";
}
```

## Install

```powershell
dotnet add package OfficeIMO.Markdown
```

## Hello, Markdown

```csharp
using OfficeIMO.Markdown;

var doc = MarkdownDoc
    .Create()
    .FrontMatter(new { title = "OfficeIMO.Markdown", tags = new[] { "docs", "reporting" } })
    .H1("OfficeIMO.Markdown")
    .P("Typed Markdown builder, reader, and HTML renderer for .NET.")
    .H2("Quick start")
    .Ul(ul => ul
        .Item("Build Markdown with a fluent API")
        .Item("Parse Markdown into typed blocks and inlines")
        .Item("Render HTML fragments or full documents"))
    .Code("powershell", "dotnet add package OfficeIMO.Markdown");

var markdown = doc.ToMarkdown();
var htmlFragment = doc.ToHtmlFragment();
var htmlDocument = doc.ToHtmlDocument();
```

## Common Tasks by Example

### Builder: headings, callouts, lists, and code

```csharp
var doc = MarkdownDoc.Create()
    .H1("Release Notes")
    .Callout("info", "Heads up", "This API is still evolving before 1.0.")
    .P("OfficeIMO.Markdown is suitable for document, reporting, and chat-style rendering flows.")
    .H2("Features")
    .Ul(ul => ul
        .Item("Typed AST/query model")
        .Item("TOC helpers")
        .Item("Front matter")
        .Item("HTML rendering"))
    .Code("csharp", "Console.WriteLine(\"Hello Markdown\");");
```

### Tables from typed objects or explicit selectors

```csharp
var results = new[] {
    new { Name = "Alice", Role = "Dev", Score = 98.5, Joined = "2024-01-10" },
    new { Name = "Bob", Role = "Ops", Score = 91.0, Joined = "2023-08-22" }
};

var doc = MarkdownDoc.Create()
    .H2("Team")
    .Table(t => t.FromSequence(results,
            ("Name", x => x.Name),
            ("Role", x => x.Role),
            ("Score", x => x.Score),
            ("Joined", x => x.Joined))
        .AlignNumericRight()
        .AlignDatesCenter());
```

### TOC helpers and scoped navigation

```csharp
var doc = MarkdownDoc.Create()
    .H1("Guide")
    .H2("Install")
    .H2("Usage")
    .H3("Tables")
    .H3("Lists")
    .TocAtTop("Contents", min: 2, max: 3)
    .H2("Appendix")
    .H3("Extra")
    .TocForPreviousHeading("Appendix Contents", min: 3, max: 3);
```

### Parse Markdown into a typed document model

```csharp
var parsed = MarkdownReader.Parse(File.ReadAllText("README.md"));

foreach (var heading in parsed.GetHeadingInfos()) {
    Console.WriteLine($"{heading.Level}: {heading.Text} -> {heading.Anchor}");
}

foreach (var block in parsed.DescendantsAndSelf()) {
    Console.WriteLine(block.GetType().Name);
}

if (parsed.HasDocumentHeader && parsed.TryGetFrontMatterValue<string>("title", out var title)) {
    Console.WriteLine("Title: " + title);
}
```

### Semantic fenced blocks and reader extensions

```csharp
var options = new MarkdownReaderOptions();
options.FencedBlockExtensions.Add(new MarkdownFencedBlockExtension(
    "Vendor charts",
    new[] { "vendor-chart" },
    context => new SemanticFencedBlock(MarkdownSemanticKinds.Chart, context.Language, context.Content, context.Caption)));

var parsed = MarkdownReader.Parse("""
```vendor-chart
{"type":"bar"}
```
""", options);
```

Use semantic fenced blocks when a fenced language represents a host contract or visual/document semantic rather than ordinary code.

### Post-parse document transforms

```csharp
var options = MarkdownReaderOptions.CreatePortableProfile();
options.DocumentTransforms.Add(
    new MarkdownJsonVisualCodeBlockTransform(MarkdownVisualFenceLanguageMode.GenericSemanticFence));

var parsed = MarkdownReader.Parse(markdown, options);
```

Use `DocumentTransforms` for AST-level cleanup that should happen after markdown is parseable but before writing, HTML rendering, or downstream export. Keep text repair in `InputNormalization` for genuinely pre-parse fixes only.

```csharp
var htmlOptions = HtmlToMarkdownOptions.CreatePortableProfile();
htmlOptions.DocumentTransforms.Add(new MarkdownInlineNormalizationTransform(
    new MarkdownInputNormalizationOptions {
        NormalizeTightParentheticalSpacing = true,
        NormalizeTightColonSpacing = true
    }));

var document = html.LoadFromHtml(htmlOptions);
```

Use `MarkdownInlineNormalizationTransform` when content is already parseable and you want AST-safe inline cleanup on an existing `MarkdownDoc`, including HTML-imported documents.

### Portable reader profile

```csharp
var portable = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreatePortableProfile());
```

Use the portable profile when portability-sensitive ingestion matters more than OfficeIMO-specific conveniences. It disables OfficeIMO-only callout/task-list parsing and turns off bare literal autolinks.
Treat that portable profile as the generic/portable contract boundary. OfficeIMO-only transcript behavior and host-specific extensions should stay opt-in on top of it rather than changing the portable defaults.

### Reader, writer, and HTML profile matrix

```csharp
var officeReader = MarkdownReaderOptions.CreateOfficeIMOProfile();
var commonMarkReader = MarkdownReaderOptions.CreateCommonMarkProfile();
var gfmReader = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();
var portableReader = MarkdownReaderOptions.CreatePortableProfile();

var officeWriter = MarkdownWriteOptions.CreateOfficeIMOProfile();
var portableWriter = MarkdownWriteOptions.CreatePortableProfile();

var portableHtml = new HtmlOptions { Kind = HtmlKind.Fragment };
MarkdownBlockRenderBuiltInExtensions.AddPortableHtmlFallbacks(portableHtml);
```

These are intentionally separate layers:

- `OfficeIMO`
  Full OfficeIMO behavior, including block parser extensions such as callouts, TOC placeholders, and footnotes.
- `CommonMark`
  Core markdown-style parsing without OfficeIMO-only or GFM-style block extensions.
- `GitHubFlavoredMarkdown`
  Core markdown plus GFM-oriented tables, task lists, and footnotes, but without OfficeIMO callouts or TOC placeholders.
- `Portable`
  Neutral parsing/output path for downstream engines that should not depend on OfficeIMO-only block syntax or HTML chrome.

The stack is now intentionally profile-driven:

- `MarkdownReaderOptions`
  Controls ingestion/parsing behavior.
- `MarkdownWriteOptions`
  Controls emitted markdown, including portable fallbacks for non-core blocks.
- `HtmlOptions`
  Controls HTML output, including portable fallbacks for callouts, TOC, and footnotes when requested.

### Opt back into specific extensions from a neutral profile

```csharp
var options = MarkdownReaderOptions.CreateCommonMarkProfile();
MarkdownReaderBuiltInExtensions.AddCallouts(options);

var parsed = MarkdownReader.Parse("""
> [!NOTE]
> This stays opt-in on top of the CommonMark-style profile.
""", options);
```

Use this pattern when a host wants a generic baseline but still needs a small, explicit set of OfficeIMO extensions.

### Control markdown output separately from parsing

```csharp
var parsed = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateOfficeIMOProfile());

var portableMarkdown = parsed.ToMarkdown(MarkdownWriteOptions.CreatePortableProfile());

var portableHtmlOptions = new HtmlOptions { Kind = HtmlKind.Fragment };
MarkdownBlockRenderBuiltInExtensions.AddPortableHtmlFallbacks(portableHtmlOptions);

var portableHtml = parsed.ToHtmlFragment(portableHtmlOptions);
```

In other words: ingest with one profile, then emit markdown or HTML with a different portability contract when needed.

### Input normalization for chat or model output

```csharp
var parsed = MarkdownReader.Parse(markdown, new MarkdownReaderOptions {
    InputNormalization = MarkdownInputNormalizationPresets.CreateIntelligenceXTranscript()
});

var strictTranscript = MarkdownReader.Parse(markdown, new MarkdownReaderOptions {
    InputNormalization = MarkdownInputNormalizationPresets.CreateIntelligenceXTranscriptStrict()
});

var docs = MarkdownReader.Parse(markdown, new MarkdownReaderOptions {
    InputNormalization = MarkdownInputNormalizationPresets.CreateDocsLoose()
});
```

Use `CreateIntelligenceXTranscript()` when the caller is intentionally ingesting IX-style transcript markdown. Use `CreateIntelligenceXTranscriptStrict()` when you need the broader repair profile for aggressively malformed transcript content.

### Streaming preview normalization for partial transcript deltas

```csharp
var preview = MarkdownStreamingPreviewNormalizer.NormalizeIntelligenceXTranscript(markdownDelta);
```

Use `NormalizeIntelligenceXTranscript(...)` when a host needs conservative cleanup for in-progress IX transcript output. This path keeps partial markdown reshaping minimal, but escalates known signal-flow and malformed-strong artifacts through the explicit `IntelligenceXTranscript` input-normalization contract.

### Explicit transcript preparation for export and DOCX hosts

```csharp
var body = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptBody(markdown);
var withoutMarkers = MarkdownTranscriptTransportMarkers.StripIntelligenceXCachedEvidenceTransportMarkers(markdown);
var export = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptForExport(withoutMarkers);
var docx = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptForDocx(markdown, preservesGroupedDefinitionLikeParagraphs: false);
var readerOptions = MarkdownTranscriptPreparation.CreateIntelligenceXTranscriptReaderOptions(
    preservesGroupedDefinitionLikeParagraphs: false,
    visualFenceLanguageMode: MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence);
```

Use `MarkdownTranscriptPreparation` when a host wants the explicit IX transcript prep contract as a visible composition point instead of manually chaining normalization, ordered-list repair, blank-line collapse, and DOCX definition-line compatibility helpers. Use `MarkdownTranscriptTransportMarkers.StripIntelligenceXCachedEvidenceTransportMarkers(...)` when the host is preparing IX transcript content for markdown export and needs that explicit transport cleanup before calling `PrepareIntelligenceXTranscriptForExport(...)`.
Those IX transcript helpers are thin named compositions over generic reader/input-normalization/document-transform building blocks. They should stay as host contracts, not become the primary implementation home for generic markdown behavior.

### HTML fragments and full documents

```csharp
var fragment = doc.ToHtmlFragment(new HtmlOptions {
    Style = HtmlStyle.GithubAuto
});

var fullPage = doc.ToHtmlDocument(new HtmlOptions {
    Title = "Release Notes",
    Style = HtmlStyle.Clean,
    CssDelivery = CssDelivery.Inline
});

doc.SaveHtml("release-notes.html", new HtmlOptions {
    Style = HtmlStyle.Word,
    CssDelivery = CssDelivery.ExternalFile
});
```

### HTML parts and asset manifests

```csharp
var parts = doc.ToHtmlParts(new HtmlOptions {
    EmitMode = AssetEmitMode.ManifestOnly,
    Prism = new PrismOptions { Enabled = true, Languages = { "csharp" } }
});

var merged = HtmlAssetMerger.Build(new[] { parts.Assets });
```

Use `ToHtmlParts(...)` and `HtmlAssetMerger.Build(...)` when a host wants to own the final HTML shell or deduplicate CSS/JS assets across multiple fragments.

## Feature Highlights

- Builder API: headings, paragraphs, links, images, tables, code fences, callouts, footnotes, front matter, TOC helpers
- Reader API: typed blocks/inlines, syntax-tree spans, traversal helpers, heading queries, list-item queries, front matter lookups
- Profiles: OfficeIMO, CommonMark-style, GFM-style, and portable reader behavior with explicit block parser extensions
- Writer profiles: default OfficeIMO markdown emission plus portable block fallbacks
- Semantic extensions: fenced block extension seam plus first-class semantic fenced-block nodes
- HTML rendering: fragment or full document, built-in themes, inline/external/link CSS delivery, CDN/offline asset handling, and portable output fallbacks
- Host integration: HTML parts/assets model for advanced embedding and shell assembly
- Table helpers: generate tables from objects or sequences, per-column alignment, renames, transforms, and formatters
- Chat/docs ingestion: named input-normalization presets for explicit IX transcript cleanup, strict chat repair, and docs cleanup workflows
- Deterministic output: stable markdown and HTML generation for snapshotting, diffs, and downstream export flows

## Detailed Feature Matrix

- Documents
  - ✅ Fluent builder and explicit block model
  - ✅ Typed traversal helpers (`TopLevelBlocks`, `DescendantsAndSelf`, `DescendantsOfType<T>()`)
  - ✅ Heading, front matter, and list-item query helpers
- Blocks
  - ✅ Headings, paragraphs, block quotes, callouts, fenced code, tables, lists, task lists, definition lists, front matter, footnotes, TOC placeholders
  - ✅ HTML comments, raw HTML blocks, horizontal rules, details/summary
- Inlines
  - ✅ Text, emphasis, strong, strike, highlight, code spans, links, images, reference links
  - ✅ Typed inline sequences and inline plain-text extraction
- Rendering
  - ✅ Markdown output
  - ✅ HTML fragment and full document output
  - ✅ Themes: Plain, Clean, GitHub Light/Dark/Auto, Chat Light/Dark/Auto, Word
  - ✅ TOC layouts: list, panel, sidebar left/right with sticky and ScrollSpy options
- Parsing
  - ✅ Typed reader with optional syntax tree
  - ✅ Portable profile for stricter, more neutral parsing defaults
  - ✅ Input normalization presets for transcript/chat/docs cleanup

## Supported Markdown

Blocks
- headings
- paragraphs and hard breaks
- fenced code blocks with optional language and captions
- images with optional size hints
- unordered, ordered, task, and definition lists
- GitHub-style pipe tables
- block quotes with lazy continuation
- callouts
- horizontal rules
- footnotes
- front matter
- TOC placeholders

Inlines
- text
- emphasis / strong / combined emphasis
- strike and highlight
- code spans
- inline and reference links
- inline and linked images

## HTML Rendering Notes

- `ToHtmlFragment(...)` is ideal for embedding into an existing page or WebView host.
- `ToHtmlDocument(...)` is ideal for standalone docs or generated reports.
- `HtmlStyle.Word` gives a document-like look with Word-style typography and table defaults.
- `ToHtmlParts(...)` exposes asset manifests for hosts that want to deduplicate or merge CSS/JS themselves.

## Benchmarks

`OfficeIMO.Markdown.Benchmarks` ships in-repo as the benchmark harness used for release-prep sanity checks.

- representative parse workloads
- representative HTML-render workloads
- default reader behavior vs portable profile

For release steps, see [../Docs/officeimo.markdown.release-checklist.md](../Docs/officeimo.markdown.release-checklist.md).

## Package Family

- `OfficeIMO.Markdown`: build, parse, query, and render Markdown/HTML
- `OfficeIMO.MarkdownRenderer`: host-oriented WebView/browser shell helpers built on top of `OfficeIMO.Markdown`
- `OfficeIMO.Word.Markdown`: Word conversion layer that uses the Markdown package family

## Dependencies & Versions

- No runtime NuGet dependencies for the core markdown library
- Targets: netstandard2.0, net472, net8.0, net10.0
- License: MIT

## Notes on Versioning

- Minor releases may add APIs, improve parser coverage, and broaden AST/query capabilities.
- Patch releases focus on correctness, compatibility, and rendering fixes.
- The current goal is a stable and intentional package baseline, not a frozen 1.0 contract yet.

## Notes

- Cross-platform: no COM automation, no Office requirement
- Deterministic output is a design goal for tests, docs, and downstream exports
- Public API and docs are being actively polished ahead of the next package line

