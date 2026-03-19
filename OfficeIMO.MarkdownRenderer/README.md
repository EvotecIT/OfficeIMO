# OfficeIMO.MarkdownRenderer — WebView-Friendly Markdown Rendering for .NET

OfficeIMO.MarkdownRenderer is a host-oriented companion package for `OfficeIMO.Markdown`. It builds full HTML shells, renders safe HTML fragments, and produces incremental update scripts for WebView2 and browser-based chat/document surfaces.

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.MarkdownRenderer)](https://www.nuget.org/packages/OfficeIMO.MarkdownRenderer)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.MarkdownRenderer?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.MarkdownRenderer)

- Targets: netstandard2.0, net472, net8.0, net10.0
- License: MIT
- NuGet: `OfficeIMO.MarkdownRenderer`
- Dependencies: `OfficeIMO.Markdown`, `OfficeIMO.Markdown.Html`, `System.Text.Json`

### AOT / Trimming Notes

- The renderer itself stays lightweight and mostly composes `OfficeIMO.Markdown`.
- Optional runtime features such as Mermaid, charts, and math are exposed through HTML shell assets rather than heavy managed dependencies.
- For minimal hosts, use the strict minimal presets to disable optional client-side renderers.

## Install

```powershell
dotnet add package OfficeIMO.MarkdownRenderer
```

For IntelligenceX-first hosts, add the first-party plugin pack as well:

```powershell
dotnet add package OfficeIMO.MarkdownRenderer.IntelligenceX
```

## Hello, Renderer

```csharp
using OfficeIMO.MarkdownRenderer;

var options = MarkdownRendererPresets.CreateStrict();

var shellHtml = MarkdownRenderer.BuildShellHtml("Markdown", options);
var bodyHtml = MarkdownRenderer.RenderBodyHtml("""
# Hello

This is rendered through OfficeIMO.MarkdownRenderer.
""", options);
var document = MarkdownRenderer.ParseDocument("# Hello", options);
var result = MarkdownRenderer.ParseDocumentResult("# Hello", options);
```

## Common Tasks by Example

### WebView2 shell + update flow

```csharp
using OfficeIMO.MarkdownRenderer;

var options = MarkdownRendererPresets.CreateStrict();

webView.NavigateToString(MarkdownRenderer.BuildShellHtml("Markdown", options));
await webView.ExecuteScriptAsync(MarkdownRenderer.RenderUpdateScript(markdownText, options));
```

### Streaming-friendly update path

```csharp
var options = MarkdownRendererPresets.CreateStrict();

webView.NavigateToString(MarkdownRenderer.BuildShellHtml("Markdown", options));

var bodyHtml = MarkdownRenderer.RenderBodyHtml(markdownText, options);
webView.CoreWebView2.PostWebMessageAsString(bodyHtml);
```

`RenderBodyHtml(...)` now treats `BaseHref` as render-local state. If the host reuses one `MarkdownRendererOptions` instance across multiple renders, the renderer restores the caller's original `HtmlOptions.BaseUri` after each call instead of leaking a previous render's base/origin policy into the next one.

### Optional chat bubble wrappers

```csharp
var options = MarkdownRendererPresets.CreateIntelligenceXTranscript();
var bubbleHtml = MarkdownRenderer.RenderChatBubbleBodyHtml(markdownText, ChatMessageRole.Assistant, options);

webView.CoreWebView2.PostWebMessageAsString(bubbleHtml);
```

### Explicit IntelligenceX transcript contract

```csharp
using OfficeIMO.MarkdownRenderer.IntelligenceX;

var options = IntelligenceXMarkdownRenderer.CreateTranscriptMinimal();
```

Use this as the preferred path for IX-style transcript rendering from a dedicated first-party package boundary.
`OfficeIMO.MarkdownRenderer` stays generic-first, while `OfficeIMO.MarkdownRenderer.IntelligenceX` exposes the IX-oriented entrypoints and keeps building on the same shared AST/document-transform contract underneath.
The core `MarkdownRendererPresets.CreateIntelligenceXTranscript*` APIs still exist and remain the underlying implementation for compatibility.

### IX desktop shell renderer contract

```csharp
using OfficeIMO.MarkdownRenderer.IntelligenceX;

var options = IntelligenceXMarkdownRenderer.CreateTranscriptDesktopShell();
```

Use `CreateTranscriptDesktopShell(...)` when the host wants the IntelligenceX desktop chat surface: minimal transcript shell defaults plus Mermaid, chart, and network visuals enabled.

### Applying transcript pre-processors without rendering

```csharp
var options = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
var normalized = MarkdownRendererPreProcessorPipeline.Apply(markdownText, options);
var diagnostics = new List<MarkdownRendererPreProcessorDiagnostic>();
var normalizedWithDiagnostics = MarkdownRendererPreProcessorPipeline.Apply(markdownText, options, diagnostics);
```

Use `MarkdownRendererPreProcessorPipeline.Apply(...)` when a host wants the explicit renderer-owned transcript preprocessor chain without rendering HTML yet. This keeps transcript preprocessor behavior sourced from OfficeIMO rather than re-iterated in app code.
That pipeline is intentionally limited to pre-parse normalization and legacy migration glue. Recoverable structure upgrades should happen later through the shared OfficeIMO AST/document-transform pipeline.
In the IntelligenceX compatibility path, the remaining recoverable cleanup now runs later as document transforms: cached-evidence marker removal, legacy JSON visual upgrades, and parseable legacy tool-slug headings. The preprocessor pipeline is no longer carrying the old legacy heading repair step.
Pass a diagnostics collection when the host wants to see whether escaped-newline normalization, shared input normalization, or custom pre-processors changed the input before parsing.

### Applying AST transforms before HTML rendering

```csharp
var options = MarkdownRendererPresets.CreateStrict();
options.DocumentTransforms.Add(new PromoteSummaryTailTransform());

var result = MarkdownRenderer.ParseDocumentResult(markdownText, options);
var document = result.Document;
var preprocessedMarkdown = result.PreprocessedMarkdown;
var originalSyntaxTree = result.SyntaxTree;
var finalSyntaxTree = result.FinalSyntaxTree;
var diagnostics = result.TransformDiagnostics;
var preProcessorDiagnostics = result.PreProcessorDiagnostics;
var bodyHtml = MarkdownRenderer.RenderBodyHtml(markdownText, options);
```

Use `options.DocumentTransforms` when the renderer host wants to rewrite the parsed `MarkdownDoc` before HTML generation.
Prefer this for structural/content fixes that should operate on the AST.
Reserve `MarkdownPreProcessors` for true pre-parse text repair and `HtmlPostProcessors` for the small set of changes that genuinely must happen on the emitted HTML string.
Use `MarkdownRenderer.ParseDocument(...)` when the host wants the final renderer-owned AST without emitting HTML yet.
Use `MarkdownRenderer.ParseDocumentResult(...)` when the host wants the final AST, the exact preprocessed markdown text that was parsed, the original pre-transform syntax tree, the final post-transform syntax tree, and both pre-parse and transform diagnostics in one typed result.

### Strict vs portable presets

```csharp
var strict = MarkdownRendererPresets.CreateStrict();
var strictPortable = MarkdownRendererPresets.CreateStrictPortable();
var minimal = MarkdownRendererPresets.CreateStrictMinimal();
var minimalPortable = MarkdownRendererPresets.CreateStrictMinimalPortable();
var relaxed = MarkdownRendererPresets.CreateRelaxed();

var ixTranscript = MarkdownRendererPresets.CreateIntelligenceXTranscript();
var ixTranscriptPortable = MarkdownRendererPresets.CreateIntelligenceXTranscriptPortable();
var ixTranscriptMinimal = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
var ixTranscriptMinimalPortable = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimalPortable();
var ixTranscriptDesktopShell = MarkdownRendererPresets.CreateIntelligenceXTranscriptDesktopShell();
var ixTranscriptRelaxed = MarkdownRendererPresets.CreateIntelligenceXTranscriptRelaxed();
```

- `CreateStrict(...)`: neutral untrusted-content defaults for generic markdown hosting
- `CreateStrictPortable(...)`: same neutral defaults, but uses the portable markdown reader profile and portable HTML output fallbacks
- `CreateStrictMinimal(...)`: neutral strict preset with Mermaid, charts, math, Prism, and copy buttons disabled
- `CreateStrictMinimalPortable(...)`: combines minimal shell behavior with the portable reader profile and portable HTML output fallbacks
- `CreateRelaxed(...)`: trusted-content preset that allows HTML parsing and sanitizes raw HTML conservatively
- `ApplyChatPresentation(...)`: composes chat presentation/chrome onto any existing preset without changing its security profile
- `CreateIntelligenceXTranscript(...)`: explicit IX transcript rendering preset with IX visual aliases and the shared transcript reader/document-transform contract
- `CreateIntelligenceXTranscriptPortable(...)`: explicit IX transcript preset using the portable reader profile
- `CreateIntelligenceXTranscriptMinimal(...)`: explicit minimal IX transcript preset for script-light chat shells
- `CreateIntelligenceXTranscriptMinimalPortable(...)`: explicit minimal IX transcript preset using the portable reader profile
- `CreateIntelligenceXTranscriptDesktopShell(...)`: explicit IX desktop-shell preset with Mermaid, chart, and network visuals enabled
- `CreateIntelligenceXTranscriptRelaxed(...)`: relaxed IX transcript preset for trusted transcript content

Generic presets register neutral fenced block languages such as `chart`, `network`, `visnetwork`, and `dataview`.
The explicit IX transcript presets additionally apply the `MarkdownRendererIntelligenceXAdapter`, which registers the IntelligenceX-oriented aliases `ix-chart`, `ix-network`, and `ix-dataview`.

Portable presets now also degrade OfficeIMO-specific HTML chrome to simpler generic HTML:

- callouts render as plain `<blockquote>` elements
- TOC placeholders render as simple lists instead of `md-toc`/sidebar navigation chrome
- footnotes render without the OfficeIMO-specific `footnotes` wrapper/list structure

If you want those fallbacks on a non-portable preset, apply them explicitly:

```csharp
var options = MarkdownRendererPresets.CreateStrict();
MarkdownRendererPresets.ApplyPortableHtmlOutputProfile(options);
```

Portable is now a composed contract, not just a parser switch:

- reader profile: `MarkdownReaderOptions.CreatePortableProfile()`
- shell/render defaults: same strict/minimal safety defaults as the non-portable preset you started from
- HTML output: OfficeIMO-specific callout, TOC, and footnote chrome degrades to generic HTML automatically

That makes the portable presets a better fit for hosts that want predictable generic output from the same AST pipeline, including `OfficeIMO.Markdown.Html` and future cross-engine bridges.
Treat the portable preset family as the generic parity boundary. Any intentionally OfficeIMO-specific or IX-specific behavior should remain opt-in rather than bleeding into the portable defaults.

### IntelligenceX alias adapter

```csharp
var options = MarkdownRendererPresets.CreateStrict();
IntelligenceXMarkdownRenderer.ApplyVisuals(options);
```

Use this when you want the generic strict/relaxed presets but still need the IntelligenceX alias fence contract without taking the full transcript preset.

### Transcript contract as a plugin

```csharp
using OfficeIMO.MarkdownRenderer.IntelligenceX;

var options = MarkdownRendererPresets.CreateStrict();
IntelligenceXMarkdownRenderer.ApplyTranscriptContract(options);
```

Use this when the host wants IX visual aliases, IX fence-option schema support, and the IX transcript reader/AST contract, but does not want the broader compatibility-pack cleanup layer.

### Generic-first opt-in composition

```csharp
var options = MarkdownRendererPresets.CreateStrictPortable();
IntelligenceXMarkdownRenderer.ApplyVisuals(options);
MarkdownRendererPresets.ApplyChatPresentation(options, enableCopyButtons: true);
```

Use this pattern when the host needs a generic-first baseline plus only a small slice of IX behavior. If the host actually wants the full IX transcript contract, prefer `CreateIntelligenceXTranscript(...)` or `CreateIntelligenceXTranscriptMinimal(...)`.

### Host-level feature packs

`MarkdownRendererPlugin` remains the low-level fenced-block renderer contract.
For broader host behavior, `MarkdownRendererFeaturePack` now groups plugins, plugin-carried fence option schemas, reader/AST configuration, preprocessors, postprocessors, and renderer defaults into one idempotent unit.

```csharp
using OfficeIMO.MarkdownRenderer.IntelligenceX;

var options = MarkdownRendererPresets.CreateStrict();
options.ApplyFeaturePack(IntelligenceXMarkdownRenderer.TranscriptCompatibilityPack);
```

Use this when a host wants a reusable compatibility bundle without hard-coding a sequence of IX-specific calls.
This is the intended pattern for future first-party or third-party host packages that build on top of the generic renderer.
If the host only needs renderer aliases plus reader/AST transcript behavior, prefer a plugin-level contract such as `IntelligenceXMarkdownRenderer.TranscriptPlugin`; use feature packs when you also need host-level compatibility glue or broader shell defaults.

### Fence metadata in custom renderers

`MarkdownFencedCodeBlockMatch` now exposes both the primary language token and parsed fence metadata:

```csharp
var options = new MarkdownRendererOptions();
options.FencedCodeBlockRenderers.Add(new MarkdownFencedCodeBlockRenderer(
    "Notes",
    new[] { "ix-note" },
    (match, _) => $"<aside data-title=\"{System.Net.WebUtility.HtmlEncode(match.FenceInfo.Title)}\">{System.Net.WebUtility.HtmlEncode(match.RawContent)}</aside>"));
```

Use `match.Language` for renderer dispatch semantics and `match.InfoString` / `match.FenceInfo` when the host wants structured fence attributes such as `title="..."` or boolean flags.
`match.FenceInfo` also carries brace-style metadata such as `{#release-note .wide}` through `ElementId` and `Classes`, so plugins can style or anchor visual/document contracts without reparsing the raw info string.
For typed host metadata, `match.FenceInfo` exposes helpers such as `TryGetBooleanAttribute(...)`, `TryGetInt32Attribute(...)`, and alias-aware `GetAttribute(...)`.

`MarkdownRendererOptions` also supports fence option schemas for host/plugin contracts:

```csharp
var options = new MarkdownRendererOptions();
options.ApplyFenceOptionSchema(new MarkdownFenceOptionSchema(
    "vendor.visual-options",
    "Vendor Visual Options",
    new[] { "vendor-chart" },
    new[] {
        MarkdownFenceOptionDefinition.Boolean("pinned"),
        MarkdownFenceOptionDefinition.Int32("maxItems", aliases: new[] { "limit" })
    }));
```

Use `TryGetFenceOptionSchema(...)` / `TryParseFenceOptions(...)` when a plugin wants declared option aliases and validation rules instead of ad hoc string parsing.
Plugins can now carry these schemas directly, and feature packs apply plugin-contributed schemas automatically alongside renderer registrations and plugin-owned reader/AST configuration.
Plugins can also carry their own idempotent reader/AST configuration directly, which lets first-party packages expose a reusable parser contract without forcing every host to depend on a higher-level feature pack.
That reader-side contract is available directly on `MarkdownReaderOptions` too: call `readerOptions.ApplyPlugin(plugin)` or `readerOptions.ApplyFeaturePack(pack)` when a host wants plugin-owned source parsing/document transforms without going through renderer presets.
Renderer plugins can now also carry renderer-stage AST transforms through `RendererDocumentTransforms`, which flow into `MarkdownRendererOptions.DocumentTransforms` when the plugin or composed feature pack is applied.
Plugins and feature packs also expose HTML-ingestion contracts for `OfficeIMO.Markdown.Html`: custom HTML element block converters, custom inline element converters, post-conversion `DocumentTransforms`, and `VisualElementRoundTripHints`. When the host also references `OfficeIMO.Markdown.Html`, apply them with `htmlOptions.ApplyPlugin(plugin)` or `htmlOptions.ApplyFeaturePack(pack)` instead of copying converters, transforms, or hints manually.

### Offline assets

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.MarkdownRenderer;

var options = MarkdownRendererPresets.CreateStrict();
options.HtmlOptions.AssetMode = AssetMode.Offline;

options.Mermaid.ScriptUrl = @"C:\app\assets\mermaid.min.js";
options.Chart.ScriptUrl = @"C:\app\assets\chart.umd.min.js";
options.Math.CssUrl = @"C:\app\assets\katex.min.css";
options.Math.ScriptUrl = @"C:\app\assets\katex.min.js";
options.Math.AutoRenderScriptUrl = @"C:\app\assets\auto-render.min.js";
```

### Shell CSS overrides

```csharp
var options = MarkdownRendererPresets.CreateIntelligenceXTranscript();
options.ShellCss = """
.omd-chat-bubble { border-radius: 18px; }
.omd-chat-row.omd-role-user .omd-chat-bubble { background: rgba(0, 120, 212, .18); }
""";
```

## Feature Highlights

- Full shell HTML builder for WebView/browser hosts
- Body fragment renderer for incremental updates
- Neutral and chat-specific strict, portable, minimal, and relaxed presets
- Optional Mermaid, Chart.js, vis-network, and math shell integration
- Chat bubble helpers and copy-button UX
- Host-friendly message contract and data attributes for native integration
- Shared markdown normalization story through `OfficeIMO.Markdown`

## Host Contract Highlights

### Primary APIs

- `BuildShellHtml(...)`
- `RenderBodyHtml(...)`
- `BuildUpdateScript(...)`
- `RenderUpdateScript(...)`
- `RenderChatBubbleBodyHtml(...)`

### WebView2 message contract

Host to web:
- `PostWebMessageAsString(bodyHtml)`
- `PostWebMessageAsJson({ type: "omd.update", bodyHtml: "..." })`

Web to host:
- `{ type: "omd.copy", text: "..." }`

### Shared visual metadata contract

Built-in visual renderers emit shared `data-omd-*` attributes:

- `data-omd-visual-kind`
- `data-omd-fence-language`
- `data-omd-fence-info` when the source fence carried normalized metadata after the language token
- `data-omd-fence-id` when the source fence provided `#id`
- `data-omd-fence-classes` when the source fence provided extra `.class` metadata
- `data-omd-visual-title` when fence metadata provides a title
- `data-omd-visual-hash`
- `data-omd-visual-contract`
- `data-omd-config-format`
- `data-omd-config-encoding`
- `data-omd-config-b64`

That keeps host integrations stable even when new visual types are added later.
Chart, network, and dataview built-ins now all flow through the same shared metadata builder, so future visual types can reuse the same contract instead of hand-assembling attributes per renderer.
When the source fence includes brace-style metadata such as `{#chart-summary .wide .compact}` or plugin-defined flags such as `pinned`, the built-in visual hosts also carry the normalized tail through explicit `data-omd-fence-*` attributes so HTML-to-markdown recovery can rebuild the original semantic fence metadata without guessing from host styling.

You can also emit the same metadata contract directly from host code through `MarkdownVisualContract.CreatePayload(...)` and `MarkdownVisualContract.BuildElementHtml(...)` when you need custom visual blocks outside the built-in renderer list. Pass a parsed `MarkdownCodeFenceInfo` when you want the helper to honor source fence `#id` / `.class` metadata too.

## Security Defaults

The strict presets are biased toward untrusted chat-style content:

- raw HTML parsing disabled
- raw HTML blocks stripped at render time
- restricted URL schemes
- `file:` URLs blocked
- `data:` URLs blocked by default
- external HTTP images blocked unless explicitly allowed
- external links hardened with `noopener noreferrer nofollow ugc`

Use the relaxed preset only for trusted or controlled content.

## Built-In Styling and Hooks

- Shell root: `#omdRoot`
- Default content wrapper: `article.markdown-body`
- Bubble wrappers: `.omd-chat-row`, `.omd-chat-bubble`
- Bubble roles: `.omd-role-user`, `.omd-role-assistant`, `.omd-role-system`
- Helper classes: `.omd-image-blocked`, `.omd-chart`, `.omd-math`

Themes come from `OfficeIMO.Markdown` HTML styles, and the chat presets default to `HtmlStyle.ChatAuto`.

## Normalization and Reader Behavior

The renderer uses `OfficeIMO.Markdown` reader options and normalization behavior underneath.

- strict presets enable the chat-output normalization helpers
- portable presets switch the underlying reader to `MarkdownReaderOptions.CreatePortableProfile()`
- portable presets also apply the portable HTML output fallbacks from `OfficeIMO.Markdown`
- the portable wrappers do both automatically, so callers do not need to remember separate reader and HTML fallback setup
- if you need the same parsing behavior outside the renderer, use `MarkdownReaderOptions` and `MarkdownInputNormalizationPresets` directly in `OfficeIMO.Markdown`

## Built-In Visuals

Fenced code blocks can be converted into shell-native visuals:

- Mermaid
- Chart.js
- vis-network
- dataview tables
- math rendering

Generic `dataview` fences accept both array-row and object-row payloads. The built-in parser recognizes neutral aliases such as `headers`/`items`, `caption`/`description`, and `callId`, while `ix-dataview` continues to mirror the legacy `data-ix-*` attributes for IntelligenceX hosts.

These features are optional and can be disabled entirely in the minimal presets.

## Dependencies & Versions

- `OfficeIMO.Markdown`
- `System.Text.Json` 8.x
- Targets: netstandard2.0, net472, net8.0, net10.0
- License: MIT

## Package Family

- `OfficeIMO.Markdown`: markdown builder, typed reader/AST, and HTML rendering
- `OfficeIMO.MarkdownRenderer`: host/WebView rendering shell and incremental update helpers
- `OfficeIMO.Word.Markdown`: Word conversion layer that can sit above the markdown family

## Notes on Versioning

- Minor releases may add presets, host hooks, and optional visual integrations.
- Patch releases focus on host compatibility, rendering correctness, and shell behavior hardening.

## Notes

- Designed for chat surfaces, docs viewers, and embedded reporting experiences
- WebView2 is a first-class scenario, but the generated HTML works in standard browser hosts too
- Keep the renderer preset and the host trust model aligned; use strict defaults unless you intentionally need relaxed HTML handling


