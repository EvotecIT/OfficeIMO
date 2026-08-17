---
title: "Convert PDF content to semantic HTML"
description: "Project readable PDF content into reviewable HTML, retain structured warnings, and distinguish semantic output from a pixel-perfect page clone."
meta.workflow_id: "pdf-html"
meta.eyebrow: "PDF semantic import"
meta.source_format: "PDF"
meta.destination_format: "HTML"
meta.package: "OfficeIMO.Html.Pdf"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Html.Pdf"
meta.runtime: "Browser-local WebAssembly or .NET"
meta.primary_url: "/apps/officeimo-converter/?workspace=convert&route=pdf-html"
meta.primary_label: "Convert PDF to HTML in the browser"
meta.secondary_url: "/docs/pdf/conversion/"
meta.secondary_label: "Read the conversion guide"
meta.summary_title: "PDF to HTML summary"
meta.limit: "The output is semantic, reviewable HTML; complex fixed-position page geometry may be approximated or omitted with warnings."
meta.related_url: "/pdf/"
meta.related_label: "Browse PDF tools and imports"
meta.howto.name: "Project PDF logical content into HTML"
meta.howto.description: "Read supported page structures, create semantic HTML, and retain conversion diagnostics."
meta.howto.steps:
  - name: "Select"
    text: "Choose a readable PDF or load the built-in sample."
  - name: "Convert"
    text: "Project logical headings, paragraphs, lists, tables, links, and supported resources into HTML."
  - name: "Review"
    text: "Preview or download the HTML and inspect the companion warning report."
---

Semantic HTML is useful for review, publishing, search ingestion, accessibility remediation, and content pipelines that need meaningful structure rather than a screenshot of each page.

## Convert from .NET

```csharp
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Open("policy.pdf");
PdfHtmlConversionResult result = pdf.ToHtmlResult();

File.WriteAllText("policy.html", result.Value);
foreach (PdfConversionWarning warning in result.Report.Warnings) {
    Console.WriteLine($"{warning.Code}: {warning.Message}");
}
```

The browser displays the generated HTML for review and downloads the same content with a JSON report describing the source, output, conversion profile, timing, and warnings.

## Semantic output versus visual reproduction

This route does not wrap a page image in HTML or promise identical browser layout. PDF reading order, positioned text, unusual fonts, drawings, and scanned pages may require approximation or separate handling. Use visual PDF rendering when fixed appearance matters; use this conversion when meaningful, editable web content is the goal.
