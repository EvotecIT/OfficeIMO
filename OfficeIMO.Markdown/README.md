OfficeIMO.Markdown (preview)
================================

Fluent and object‑model Markdown builder for .NET with CommonMark/GFM‑style output.

Key points
- No external dependencies.
- Fluent API and explicit object model.
- Basic blocks: headings, paragraphs, links, images, lists, tables, code blocks, callouts.
- Front matter (YAML) support.

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
