# OfficeIMO.Adf

`OfficeIMO.Adf` provides a dependency-light Atlas Document Format model plus conversions through OfficeIMO's Markdown and HTML engines.

```csharp
AdfDocument document = AdfDocument.Parse(adfJson);
AdfConversionResult<string> markdown = AdfConverter.ToMarkdown(document);
AdfConversionResult<string> html = AdfConverter.ToHtml(document);

AdfConversionResult<AdfDocument> fromMarkdown = AdfConverter.FromMarkdown("# Status\n\nReady.");
```

Unknown ADF nodes, marks, attributes, and extension properties remain in the parsed model and survive JSON round trips. A projection to Markdown or HTML reports unsupported constructs through `AdfConversionReport` instead of silently claiming full fidelity.

Structural validation enforces list and task-list parent/child contracts. Markdown task markers inside ordinary bullet or ordered lists remain visible text and produce a fidelity warning rather than generating invalid ADF hierarchy.
