# OfficeIMO.Adf

`OfficeIMO.Adf` provides a dependency-light Atlas Document Format model plus conversions through OfficeIMO's Markdown and HTML engines.

```csharp
using OfficeIMO.Adf;

AdfDocument document = AdfDocument.Parse(adfJson);
AdfConversionResult<string> markdown = AdfConverter.ToMarkdown(document);
AdfConversionResult<string> html = AdfConverter.ToHtml(document);

AdfConversionResult<AdfDocument> fromMarkdown = AdfConverter.FromMarkdown("# Status\n\nReady.");
```

Unknown ADF nodes, marks, attributes, and extension properties remain in the parsed model and survive JSON round trips. A projection to Markdown or HTML reports unsupported constructs through `AdfConversionReport` instead of silently claiming full fidelity.

## Validate the ADF contract

`Validate()` checks recognized structural relationships while warning about unknown node and mark types. Select `FullSchema` when the caller needs the complete pinned Atlassian schema contract:

```csharp
AdfValidationResult validation = document.Validate(new AdfValidationOptions {
    Profile = AdfValidationProfile.FullSchema,
    MaximumSchemaEvaluations = 100000000
}, cancellationToken);
```

The package embeds the unmodified ADF schema from `@atlaskit/adf-schema` 57.6.21. Its [provenance and license notices](THIRD-PARTY-NOTICES.md) identify the source and checksum. Validation checks required node/mark properties, attributes, child alternatives, and schema limits without another runtime package. A rule-budget exhaustion returns an invalid result with `ADF_SCHEMA_LIMIT`. Node graphs reject cycles and null entries and are bounded to 64 levels and one million nodes and marks.

Schema validity and acceptance by a specific Jira or Confluence API are separate contracts. Product settings and destination capabilities can impose further restrictions.

Use `AdfDestinationPolicy` to record a caller-qualified product/version's supported content nodes and marks:

```csharp
var destination = new AdfValidationOptions {
    Profile = AdfValidationProfile.FullSchema,
    DestinationPolicy = new AdfDestinationPolicy(
        "My integration's qualified destination 2026-10",
        allowedNodeTypes: new[] { "paragraph", "text", "hardBreak", "bulletList", "listItem" },
        allowedMarkTypes: new[] { "strong", "em", "link" })
};
AdfValidationResult supported = document.Validate(destination, cancellationToken);

var generated = AdfConverter.FromMarkdown("**Review**", new AdfConversionOptions {
    DestinationValidation = destination
});
if (generated.Report.HasErrors) {
    // Inspect generated.Report.Diagnostics before submitting generated.Value.
}
```

The policy copies its lists and compares names case-sensitively. Null lists leave that category unrestricted; empty lists permit none. Restrictions apply throughout nested content and report up to 1,000 destination errors as `ADF_DESTINATION_NODE` and `ADF_DESTINATION_MARK`. They add to schema checks without rewriting the document. Generated ADF remains inspectable when validation fails, and `Report.RequireNoLoss()` rejects those errors. The three-argument `FromHtml(html, htmlOptions, adfOptions)` overload accepts the same generation options.

Caller lists describe the destination capabilities the caller has qualified. They do not prove current Jira or Confluence API acceptance, attribute-specific product rules, permissions, or tenant configuration.

## Project visible content with fidelity evidence

Mentions, emoji, status/date nodes, cards, media, expand/panel/decision containers, nested tasks, and rich table cells have visible projections where the target model can carry them. The library uses local labels and metadata and does not fetch card, media, or mention resources.

Known nodes also report attributes, marks, and extension properties omitted or approximated by projection. Inspect `Report.Diagnostics`, or call `Report.RequireNoLoss()` to reject a lossy result. Retaining data in native JSON does not make a Markdown or HTML projection lossless.

Markdown-to-ADF conversion uses parent-aware output. Tasks and other blocks that are invalid in their destination context receive a visible fallback with a diagnostic. Task IDs are deterministic for identical input; callers can supply a unique-ID policy:

```csharp
var options = new AdfConversionOptions {
    LocalIdFactory = path => "document-42:" + path
};
AdfDocument tasks = AdfConverter.FromMarkdown("- [ ] Review", options).Value;
```

`ExtensionResolver` accepts a caller-owned projection for `extension`, `bodiedExtension`, and `inlineExtension` nodes. Returning null uses the ordinary fallback. Inline extension output must be empty or a single paragraph. The resolver's projection is reported separately, and the original native extension payload remains available for JSON writing.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 4 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Adf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
