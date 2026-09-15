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

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 4 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Adf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
