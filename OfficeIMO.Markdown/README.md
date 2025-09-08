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

