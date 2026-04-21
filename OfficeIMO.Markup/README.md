# OfficeIMO.Markup

OfficeIMO.Markup is a unified Markdown-inspired authoring layer for OfficeIMO. It keeps the common authoring path familiar while giving Word, Excel, and PowerPoint richer semantic constructs when plain Markdown is not enough.

The layer is intentionally split into stages:

- parse Markdown, front matter, and OfficeIMO directives
- produce a semantic AST that is independent from C# and PowerShell APIs
- validate the AST against `Presentation`, `Document`, or `Workbook`
- emit starter C# or PowerShell code from the AST
- export real Office files through profile-specific exporters

## Core Markdown

Headings, paragraphs, lists, fenced code, images, and pipe tables are parsed through `OfficeIMO.Markdown` and mapped into shared semantic AST nodes.

```markdown
# Quarterly Review

Highlights from the quarter.

- Revenue grew
- Churn improved

| Product | Revenue |
| --- | ---: |
| A | 120 |
```

## OfficeIMO Extensions

Office-specific constructs use two directive families:

- `@...` configures a major Office container such as `@slide`, `@section`, or `@sheet`.
- `::...` inserts an Office-aware block such as `::notes`, `::chart`, `::mermaid`, `::range`, or `::formula`.

Front matter configures the whole file.

```markdown
---
profile: presentation
title: Quarterly Review
theme: evotec-modern
---

# Quarterly Review

@slide {
  layout: title-and-content
  transition: fade
}

- Revenue grew
- Churn improved

::notes
Open with the top-line result.
```

Presentation layout directives are parsed as semantic layout nodes, not generic extension text. Shared placement attributes use `x`, `y`, `w`, and `h`; percentages are resolved by the target exporter, while plain numbers are treated as target-native lengths.

```markdown
::textbox x=6% y=8% w=70% h=10% style=hero-title
Pipeline: From Text to PPTX

::columns gap=4%

::column width=48%
## Flow
- Markdown
- AST
- PowerPoint

::card title="Fast path"
Start simple, then add layout only where the slide needs control.
```

The first style resolver is intentionally small and semantic. Built-in names such as `hero-title`, `lead`, `body`, `caption`, `card`, `callout`, and `accent` resolve from the document theme into preview and PowerPoint text formatting. Inline attributes can override the built-in values when needed:

```markdown
::textbox style=hero-title color=#FFFFFF font-size=34
Board-ready automation

::card title="Risk" style=callout fill=#E8F2FF border=#2563EB
Keep the Markdown path clean, then drop into code when the deck needs custom logic.
```

```markdown
---
profile: workbook
title: Revenue Workbook
---

@sheet {
  name: Revenue
}

::range address=A1
Product,2024,2025
A,100,120
B,80,92
C,60,77

::table name="RevenueTable" range=A1:C4 header=true

::formula cell=D2
=C2-B2

::chart type=column title="Revenue" source=A1:C4 cell=F2 width=480 height=320 category-title=Product value-title=Revenue value-format="#,##0" legend=right labels=true label-position=outside-end
```

Workbook targets can be sheet-qualified when that reads better than switching the current sheet: `Data!A1`, `Data!C2`, `Data!A1:C10`, and `cell=Dashboard!B3` are accepted for ranges, formulas, formatting, tables, and chart placement. Workbook charts can also target a dashboard sheet while reading data from a named table on another sheet. Chart attributes stay semantic: axis titles, axis number formats, legend position, data labels, label number formats, and gridlines map to native editable Excel chart features. The exporter keeps the workbook editable, freezes table header rows, hides gridlines for generated sheets, auto-fits worksheet columns, applies a small default palette to chart series, and now opts into the existing OfficeIMO Excel save-time preflight, defined-name repair, and Open XML validation path by default so generated workbooks are less likely to trigger Excel repair prompts. The CLI keeps those protections on by default for workbook export and exposes `--no-safe-preflight`, `--no-defined-name-repair`, and `--no-openxml-validation` as explicit escape hatches when you need to debug a problematic workbook payload.

Workbook formatting directives can also layer style attributes without dropping earlier ones, so combinations like number formats, fills, font color, bold, italic, underline, horizontal/vertical alignment, wrap settings, and simple all-sides borders remain stable when applied to the same cells. The C# and PowerShell emitters now translate the same semantic `::format` blocks into concrete worksheet formatting calls, which makes the escape hatch feel much closer to the final workbook export instead of falling back to placeholder comments. The VS Code workbook preview also applies those formatting directives onto a lightweight sheet grid, so authors can see styled cells, alignment, emphasis, borders, and wrapped content before exporting.

```markdown
::format target=Data!B2:C3 numberFormat="#,##0" fill=#D9EAD3 color=#112233 bold=true italic=true underline=true align=center valign=middle wrap=true border=thin border-color=#445566
```

```markdown
@sheet {
  name: Revenue
}

::range address=A1
Product,2024,2025
A,100,120
B,80,92
C,60,77

::table name="RevenueTable" range=A1:C4 header=true

@sheet {
  name: Dashboard
}

::chart type=column title="Revenue" source=Revenue!RevenueTable cell=Dashboard!B3 width=640 height=360 category-title=Product value-title=Revenue value-format="#,##0" legend=right labels=true label-position=outside-end label-format="#,##0" gridlines=true
```

Mermaid diagrams are represented as semantic diagram nodes and can later be rendered to images for Word or PowerPoint.

```markdown
::mermaid x=8% y=20% w=60% h=48% fit=contain
flowchart LR
  Markdown --> AST --> Code
```

PowerPoint export can render Mermaid diagrams to PNG when Mermaid CLI (`mmdc`) is available. Set `OFFICEIMO_MARKUP_MERMAID_CLI` or pass `--mermaid-renderer <path-to-mmdc>` to the CLI. The VS Code extension also includes `OfficeIMO Markup: Install Mermaid Renderer`, which installs Mermaid CLI into the extension's local storage and remembers the path for future exports. The CLI also checks the default extension install location under the user's temporary OfficeIMO Markup tools folder, so exports can keep working after the editor installs the renderer. Use `fit=contain` for aspect-safe diagrams or `fit=stretch` when a diagram should occupy the authored frame. Rendered Mermaid diagrams get a light visual panel by default so transparent PNGs land as designed slide objects; use `panel=false` to remove it. Image blocks can opt into the same frame with `panel=true` or `frame=true`. When no renderer is available, the exporter keeps a readable text fallback so the deck still opens cleanly.

Presentation export uses the PowerPoint designer helpers when markup stays semantic, and falls back to a branded editable canvas when the slide includes explicit placement. Simple title-and-content slides with two to four bullet points become summary cards, clean `layout: two-columns` slides with semantic `::columns` regions become balanced themed column panels, while positioned diagram/chart slides get a light theme wash, accent rail, visual panels, and chart panels instead of floating on a plain white slide. Presentation charts understand the same semantic chart attributes as workbook charts for axis titles, number formats, legend position, data labels, label formats, and gridline control.

Presentation backgrounds can stay semantic as well. `background: solid(#FFFFFF)` applies a slide background color, `background: gradient(primary, accent1)` or `background: gradient(#0F172A, #2563EB)` maps to a native linear gradient in PowerPoint export, `angle=45` can tune the gradient direction when needed, `background: image("./images/mesh.png")` resolves relative to the markup file during CLI/export flows, and `overlay=rgba(0,0,0,0.35)` adds a full-slide overlay that keeps text readable over image backgrounds. Theme-style aliases such as `primary`, `accent1`, `accent2`, `background`, `surface`, and `text1` can be used in background colors for both export and preview. When a slide declares its own background, the exporter and VS Code preview now suppress the branded fallback canvas chrome so authored backgrounds remain the visual source of truth.

Inline chart data uses a compact CSV shape. The first row defines headers, the first column defines categories, and the remaining columns become chart series.

```markdown
::chart type=column title="Quarterly Revenue" category-title=Quarter value-title=Amount value-format="#,##0" legend=right labels=true label-position=outside-end label-format="#,##0"
Quarter,Revenue,Costs
Q1,120,85
Q2,180,94
Q3,260,132
Q4,320,150
```

The VS Code preview renders quick single-series or grouped multi-series chart bars for authoring feedback and surfaces the same semantic chart metadata used by the real exporters, including source ranges, source kind, target cell, chart size, axis titles, number formats, legend position, data labels, label format, and gridline settings. Source-backed workbook charts show as native chart placeholders until export. Presentation previews also honor explicit slide backgrounds, including local background images, theme-aware gradients, simple gradients, `fit=contain` or `fit=stretch`, and overlay washes, so the preview stays much closer to the real slide composition.

Fenced `officeimo` blocks are still accepted as a compatibility syntax for tools that must stay strictly inside CommonMark fenced-code constructs.

## API Sketch

```csharp
using OfficeIMO.Markup;

var result = OfficeMarkupParser.Parse(markup, new OfficeMarkupParserOptions {
    Profile = OfficeMarkupProfile.Presentation
});

var csharp = new OfficeMarkupCSharpEmitter().Emit(result.Document);
var powershell = new OfficeMarkupPowerShellEmitter().Emit(result.Document);
```

Emitters are intended as an escape hatch rather than a second rendering pipeline. They preserve semantic hints such as layout, presentation sections, placement, transition details, notes, chart categories/series, formulas, ranges, sheet-qualified workbook targets, workbook formatting directives, and Mermaid/image handoff comments so authors can start in readable markup and then continue in C# or PowerShell when a document needs custom OfficeIMO automation. Presentation transitions now resolve common directional forms such as `push direction=left`, `push direction=up`, `warp direction=out`, `ferris direction=right`, `blinds direction=vertical`, and `comb direction=horizontal` into native `OfficeIMO.PowerPoint` transition enums, while common timing attributes such as `duration=0.6`, `speed=fast`, `advance-on-click=false`, and `advance-after=5` flow into real `PowerPointSlide` transition properties and are still preserved in emitted comments as structured effect, native-enum, direction, and timing hints.

## CLI Sketch

```powershell
dotnet run --project OfficeIMO.Markup.Cli -- parse OfficeIMO.Markup\Examples\presentation.omd --format json
dotnet run --project OfficeIMO.Markup.Cli -- validate OfficeIMO.Markup\Examples\presentation.omd
dotnet run --project OfficeIMO.Markup.Cli -- emit OfficeIMO.Markup\Examples\presentation.omd --target csharp --output presentation.cs
dotnet run --project OfficeIMO.Markup.Cli -- export OfficeIMO.Markup\Examples\presentation.omd --target pptx --output presentation.pptx
dotnet run --project OfficeIMO.Markup.Cli -- export OfficeIMO.Markup\Examples\presentation.omd --target pptx --output presentation.pptx --mermaid-renderer C:\Tools\mmdc.cmd
dotnet run --project OfficeIMO.Markup.Cli -- export OfficeIMO.Markup\Examples\workbook.omd --target xlsx --output workbook.xlsx
dotnet run --project OfficeIMO.Markup.Cli -- export OfficeIMO.Markup\Examples\document.omd --target docx --output document.docx
```

The first real-file exporters target PowerPoint presentations, Excel workbooks, and Word documents. PowerPoint export maps presentation-profile AST nodes into an editable `.pptx` using `OfficeIMO.PowerPoint`, including slides, real PowerPoint sections from slide `section:` metadata, titles, text, lists, tables, images, native linear-gradient backgrounds, relative background-image directives with optional overlays, simple placement directives, transitions with directional variants plus timing/speed/advance metadata where the core library supports them, speaker notes, styled native charts from inline CSV chart data, and optional Mermaid-to-image export when Mermaid CLI is available. Excel export maps workbook-profile AST nodes into an editable `.xlsx` using `OfficeIMO.Excel`, including sheets, sheet-qualified ranges, formulas, named tables, number formatting, dashboard charts from inline data, worksheet ranges, or sheet-qualified table sources such as `Revenue!RevenueTable`, plus safe workbook defaults for gridlines, frozen table headers, auto-fit columns, styled Markdown headings, and chart series colors. Word export maps document-profile AST nodes into an editable `.docx` using `OfficeIMO.Word`, including headings, paragraphs, lists, tables, images, sections, headers, footers, table of contents, page breaks, and native charts from inline CSV chart data.
