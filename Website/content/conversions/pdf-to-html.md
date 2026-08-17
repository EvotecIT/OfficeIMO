---
title: "Convert PDF content to reviewable HTML"
description: "Choose semantic or positioned-review HTML, retain structured warnings, and distinguish review output from a pixel-perfect page clone."
meta.workflow_id: "pdf-html"
meta.eyebrow: "PDF review import"
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
meta.limit: "The browser uses positioned-review HTML; complex graphics, optional content, and scans without OCR may still be approximated or omitted with warnings."
meta.related_url: "/pdf/"
meta.related_label: "Browse PDF tools and imports"
meta.howto.name: "Project PDF content into reviewable HTML"
meta.howto.description: "Read supported page structures, preserve page-aware review geometry in the browser, and retain conversion diagnostics."
meta.howto.steps:
  - name: "Select"
    text: "Choose a readable PDF or load the built-in sample."
  - name: "Convert"
    text: "Project text, tables, links, forms, images, and supported page geometry into positioned-review HTML."
  - name: "Review"
    text: "Preview or download the HTML and inspect the companion warning report."
---

OfficeIMO exposes two explicit PDF-to-HTML profiles. Semantic output is useful for publishing, search ingestion, accessibility remediation, and content pipelines. Positioned-review output keeps page containers and supported geometry so reviewers can inspect how extracted content relates to the source page. Neither profile is a browser clone of the PDF renderer.

## Convert from .NET

```csharp
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Open("policy.pdf");
PdfHtmlConversionResult semantic = pdf.ToHtmlResult();
PdfHtmlConversionResult review = pdf.ToHtmlResult(new PdfHtmlSaveOptions {
    Profile = PdfHtmlProfile.PositionedReview,
    IncludeLinkAnnotations = true,
    IncludeFormWidgets = true
});

File.WriteAllText("policy-review.html", review.Value);
foreach (PdfConversionWarning warning in review.Report.Warnings) {
    Console.WriteLine($"{warning.Code}: {warning.Message}");
}
```

The .NET API defaults to the semantic profile. The browser workbench deliberately selects `PositionedReview`, displays that generated HTML, and downloads the same content with a JSON report describing the source, output, conversion profile, timing, and warnings.

## Semantic output versus visual reproduction

This route does not wrap a page image in HTML or promise identical browser layout. PDF reading order, unusual fonts, complex drawings, optional content, and scanned pages may require approximation or separate handling. Use visual PDF rendering when fixed appearance matters, positioned review when page-aware inspection matters, and semantic output when meaningful editable web content is the goal.
