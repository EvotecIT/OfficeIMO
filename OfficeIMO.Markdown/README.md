OfficeIMO.Markdown
==================

Fluent and object‑model Markdown builder for .NET with CommonMark/GFM‑style output. Zero runtime dependencies, rich table/list helpers, HTML export (fragment or full document) with built‑in styles and Prism highlighting.

## Why OfficeIMO.Markdown

- Pure .NET, cross‑platform — no native renderers required
- Fluent API or explicit object model — compose documents predictably
- CommonMark/GFM‑style output that renders well on GitHub, wikis, and docs sites
- HTML export with clean themes (Clean, GitHub Light/Dark/Auto), CDN/offline assets, and Prism highlighting
- Designed for reporting — tables from sequences/objects, callouts, TOC, and front matter

### Design & Expectations

- Deterministic output — stable ordering and formatting make diffs/snapshots easy
- Extensible blocks — add custom blocks with small helpers instead of string concatenation
- Good defaults — header transforms (Pretty + acronyms), numeric/date alignment heuristics
- Performance — string builders and pooled buffers; no exceptions in hot loops (e.g., header lookups)

Badges

<!-- Replace OWNER/REPO and workflow names to match your repo; update NuGet id when published. -->
[![NuGet](https://img.shields.io/nuget/v/OfficeIMO.Markdown.svg)](https://www.nuget.org/packages/OfficeIMO.Markdown)
[![NuGet Downloads](https://img.shields.io/nuget/dt/OfficeIMO.Markdown.svg)](https://www.nuget.org/packages/OfficeIMO.Markdown)
[![Build](https://img.shields.io/github/actions/workflow/status/EvotecIT/OfficeIMO/ci.yml?branch=main)](https://github.com/EvotecIT/OfficeIMO/actions)
[![Coverage](https://img.shields.io/codecov/c/github/EvotecIT/OfficeIMO.svg)](https://app.codecov.io/gh/EvotecIT/OfficeIMO)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](../LICENSE)

Highlights
- No external dependencies for Markdown/HTML generation.
- Fluent API and explicit object model.
- Core blocks: headings, paragraphs, links, images (with captions), lists (UL/OL/task/definition), tables, code blocks, callouts, front matter (YAML).
- Tables from objects/sequences: include/exclude/order/rename/formatters, alignment presets, date/number heuristics.
- Table of Contents: at top, here, or scoped to sections; GitHub‑style anchors.
- HTML: fragment or full document; styles (Clean, GitHub Light/Dark/Auto), CSS delivery (inline/link/external file), Online/Offline asset modes.
- Prism highlighting: CDN link or offline inline; manifest for safe dedupe across fragments.

Install

```bash
dotnet add package OfficeIMO.Markdown
```

Example

```csharp
using OfficeIMO.Markdown;

var md = MarkdownDoc
    .Create()
    .FrontMatter(new { title = "DomainDetective", tags = new[] { "dns", "email", "security" } })
    .H1("DomainDetective")
    .P("All-in-one DNS, email, and TLS analyzer with rich reports.")
    .Callout("info", "Early access", "APIs may change before 1.0.")
    .H2("Install")
    .Code("bash", "dotnet tool install -g DomainDetective")
    .H2("Quick start")
    .Code("powershell",
        "Test-DDMailDomainClassification -DomainName 'evotec.pl','evotec.xyz' -ExportFormat Word")
    .H2("Features")
    .Ul(ul => ul
        .Item("SPF/DKIM/DMARC scoring")
        .Item("TLS/SSL tests and cipher hints")
        .Item("WHOIS, MX, PTR, DNSSEC, BIMI")
        .Item("Exports: Word, HTML, PDF, Markdown"))
    .H2("Links")
    .Ul(ul => ul
        .ItemLink("Docs", "https://evotec.xyz/hub/")
        .ItemLink("Issues", "https://github.com/EvotecIT/DomainDetective/issues"));

var markdown = md.ToMarkdown();
var htmlFrag = md.ToHtmlFragment();
var htmlDoc  = md.ToHtmlDocument();
```

Status: early preview. API may evolve before 1.0.

Advanced usage

```csharp
using OfficeIMO.Markdown;
using System.Globalization;

var people = new[] {
    new { Name = "Alice", Role = "Dev", Score = 98.5, Joined = "2024-01-10" },
    new { Name = "Bob", Role = "Ops", Score = 91.0, Joined = "2023-08-22" }
};

// 1) Control header casing + acronyms (user-provided)
var headerTx = HeaderTransforms.PrettyWithAcronyms(new[] { "ID", "DMARC", "SPF" });
MarkdownDoc.Create()
    .H2("FromAny with header transform")
    .Table(t => t.Columns(headerTx).FromAny(new { DmarcPolicy = "p=none", SpfAligned = true }))

    // 2) FromSequence with explicit columns + numeric/date alignment
    .H2("FromSequence with selectors")
    .Table(t => t.FromSequence(people,
            ("Name",   x => x.Name),
            ("Role",   x => x.Role),
            ("Score",  x => x.Score),
            ("Joined", x => x.Joined))
        .AlignNumericRight()     // right-align numeric columns
        .AlignDatesCenter())     // center-align date-like columns

    // 3) TOC at top + scoped TOC inside a section
    .H2("Usage")
    .H3("Tables")
    .H3("Lists")
    .TocAtTop("Contents", min: 2, max: 3)
    .H2("Appendix").H3("Extra")
    .TocForPreviousHeading("Appendix Contents", min: 3, max: 3);

// HTML rendering
var htmlFragment = md.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.GithubAuto });
var htmlDocument = md.ToHtmlDocument(new HtmlOptions { Title = "Report", Style = HtmlStyle.Clean, CssDelivery = CssDelivery.Inline });
md.SaveHtml("Report.html", new HtmlOptions { Style = HtmlStyle.Clean, CssDelivery = CssDelivery.ExternalFile }); // writes Report.html + Report.css

// CDN link online vs offline (download and inline)
var cdn = "https://cdn.jsdelivr.net/npm/github-markdown-css@5.5.1/github-markdown.min.css";
var htmlCdnOnline  = md.ToHtmlDocument(new HtmlOptions { CssDelivery = CssDelivery.LinkHref, CssHref = cdn, AssetMode = AssetMode.Online, BodyClass = "markdown-body" });
var htmlCdnOffline = md.ToHtmlDocument(new HtmlOptions { CssDelivery = CssDelivery.LinkHref, CssHref = cdn, AssetMode = AssetMode.Offline, BodyClass = "markdown-body" });
```

Header transforms and acronyms

```csharp
// Provide your own acronyms to uppercase in headers
var tx = HeaderTransforms.PrettyWithAcronyms(new[] { "ID", "DMARC", "SPF" });
MarkdownDoc.Create().Table(t => t.Columns(tx).FromAny(new { DmarcPolicy = "p=none", SpfAligned = true, Id = 25 }));

// Or inline via options
MarkdownDoc.Create().Table(t => t.FromAny(new { DmarcPolicy = "p=none" }, o => o.HeaderTransform = HeaderTransforms.PrettyWithAcronyms(new[] { "DMARC" })));
```

TOC helpers

```csharp
var md = MarkdownDoc.Create()
    .H1("Report")
    .H2("Intro").P("...")
    .H2("Usage").H3("Tables").H3("Lists")
    .TocAtTop("Contents", min: 2, max: 3) // top-level TOC
    .H2("Appendix").H3("Extra")
    .TocHere(o => { o.MinLevel = 3; o.MaxLevel = 3; }) // TOC at current position
    .TocForSection("Usage", "Usage Contents", min: 3, max: 3); // scoped to a named section
```

Column Options (FromAny)

| Option          | Type                              | Purpose |
|-----------------|-----------------------------------|---------|
| Include         | HashSet<string>                   | Only include these properties |
| Exclude         | HashSet<string>                   | Exclude these properties |
| Order           | List<string>                      | Order of headers (others follow) |
| HeaderRenames   | Dictionary<string,string>         | Map property name → header text |
| HeaderTransform | Func<string,string>               | Transform header when no rename specified |
| Formatters      | Dictionary<string, Func<object?,string>> | Per‑property cell formatting |
| Alignments      | List<ColumnAlignment>             | Per‑column alignment for header/cells |

Heuristics
- AlignNumericRight(threshold=0.8): right‑align columns that look numeric.
- AlignDatesCenter(threshold=0.6): center‑align columns that look like dates.

Doc‑level helpers
- TableAuto(build, alignNumeric=true, alignDates=true)
- TableFromAuto(data, configure?, alignNumeric=true, alignDates=true)

HTML options
- Kind: Fragment | Document
- Style: Plain | Clean | GithubLight | GithubDark | GithubAuto
- CssDelivery: Inline | ExternalFile | LinkHref | None
- AssetMode: Online (link) | Offline (download+inline)
- Title, BodyClass (default "markdown-body"), IncludeAnchorLinks, ThemeToggle
- EmitMode: Emit (default) | ManifestOnly for host-side asset merging
- Prism: Enabled, Theme (Prism/Okaidia/GithubDark/GithubAuto), Languages, Plugins, CdnBase

De‑duping assets
- ToHtmlParts returns Assets: a list of { Id, Kind (Css/Js), Href or Inline }.
- Tags we emit include data-asset-id for easy deduplication if you concatenate HTML.
- Set EmitMode = ManifestOnly to suppress emitting <link>/<script> tags and merge Assets yourself.

Targets
- netstandard2.0 (library)
- net8.0, net9.0

License

MIT — see LICENSE.

Contributing

Issues and PRs welcome. Please keep one‑class/enum per file, avoid external runtime deps, and add targeted tests for new features.
