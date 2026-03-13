# OfficeIMO.MarkdownRenderer — WebView-Friendly Markdown Rendering for .NET

OfficeIMO.MarkdownRenderer is a host-oriented companion package for `OfficeIMO.Markdown`. It builds full HTML shells, renders safe HTML fragments, and produces incremental update scripts for WebView2 and browser-based chat/document surfaces.

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.MarkdownRenderer)](https://www.nuget.org/packages/OfficeIMO.MarkdownRenderer)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.MarkdownRenderer?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.MarkdownRenderer)

- Targets: netstandard2.0, net472, net8.0, net10.0
- License: MIT
- NuGet: `OfficeIMO.MarkdownRenderer`
- Dependencies: `OfficeIMO.Markdown`, `System.Text.Json`

### AOT / Trimming Notes

- The renderer itself stays lightweight and mostly composes `OfficeIMO.Markdown`.
- Optional runtime features such as Mermaid, charts, and math are exposed through HTML shell assets rather than heavy managed dependencies.
- For minimal hosts, use the strict minimal presets to disable optional client-side renderers.

## Install

```powershell
dotnet add package OfficeIMO.MarkdownRenderer
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

### Optional chat bubble wrappers

```csharp
var options = MarkdownRendererPresets.CreateChatStrict();
var bubbleHtml = MarkdownRenderer.RenderChatBubbleBodyHtml(markdownText, ChatMessageRole.Assistant, options);

webView.CoreWebView2.PostWebMessageAsString(bubbleHtml);
```

### Generic-first chat composition

```csharp
var options = MarkdownRendererPresets.CreateStrictMinimal();
MarkdownRendererPresets.ApplyChatPresentation(options, enableCopyButtons: false);
MarkdownRendererIntelligenceXAdapter.Apply(options);
```

Use this as the preferred composition path for downstream hosts. The `CreateChatStrict*` helpers remain available as compatibility wrappers.

### Strict vs portable presets

```csharp
var strict = MarkdownRendererPresets.CreateStrict();
var strictPortable = MarkdownRendererPresets.CreateStrictPortable();
var minimal = MarkdownRendererPresets.CreateStrictMinimal();
var minimalPortable = MarkdownRendererPresets.CreateStrictMinimalPortable();
var relaxed = MarkdownRendererPresets.CreateRelaxed();

var chatStrict = MarkdownRendererPresets.CreateChatStrict();
var chatMinimal = MarkdownRendererPresets.CreateChatStrictMinimal();
```

- `CreateStrict(...)`: neutral untrusted-content defaults for generic markdown hosting
- `CreateStrictPortable(...)`: same neutral defaults, but uses the portable markdown reader profile and portable HTML output fallbacks
- `CreateStrictMinimal(...)`: neutral strict preset with Mermaid, charts, math, Prism, and copy buttons disabled
- `CreateStrictMinimalPortable(...)`: combines minimal shell behavior with the portable reader profile and portable HTML output fallbacks
- `CreateRelaxed(...)`: trusted-content preset that allows HTML parsing and sanitizes raw HTML conservatively
- `ApplyChatPresentation(...)`: composes chat presentation/chrome onto any existing preset without changing its security profile
- `CreateChatStrict(...)`: compatibility wrapper built on top of the strict preset family
- `CreateChatStrictPortable(...)`: portable chat wrapper
- `CreateChatStrictMinimal(...)`: minimal chat wrapper
- `CreateChatStrictMinimalPortable(...)`: minimal portable chat wrapper
- `CreateChatRelaxed(...)`: relaxed chat wrapper

Generic presets register neutral fenced block languages such as `chart`, `network`, `visnetwork`, and `dataview`.
Chat presets additionally apply the `MarkdownRendererIntelligenceXAdapter`, which registers the IntelligenceX-oriented aliases `ix-chart`, `ix-network`, and `ix-dataview`.

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

### IntelligenceX alias adapter

```csharp
var options = MarkdownRendererPresets.CreateStrict();
MarkdownRendererIntelligenceXAdapter.Apply(options);
```

Use this when you want the generic strict/relaxed presets but still need the IntelligenceX alias fence contract.

### Recommended profile composition

```csharp
var options = MarkdownRendererPresets.CreateStrictPortable();
MarkdownRendererIntelligenceXAdapter.Apply(options);
MarkdownRendererPresets.ApplyChatPresentation(options, enableCopyButtons: true);
```

Use this pattern when the host needs a generic-first baseline plus explicit IntelligenceX visual aliases and chat chrome. It keeps the OfficeIMO/IX-specific behavior opt-in instead of making it part of the default markdown contract.

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
var options = MarkdownRendererPresets.CreateChatStrict();
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
- `data-omd-visual-hash`
- `data-omd-visual-contract`
- `data-omd-config-format`
- `data-omd-config-encoding`
- `data-omd-config-b64`

That keeps host integrations stable even when new visual types are added later.
Chart, network, and dataview built-ins now all flow through the same shared metadata builder, so future visual types can reuse the same contract instead of hand-assembling attributes per renderer.

You can also emit the same metadata contract directly from host code through `MarkdownVisualContract.CreatePayload(...)` and `MarkdownVisualContract.BuildElementHtml(...)` when you need custom visual blocks outside the built-in renderer list.

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

