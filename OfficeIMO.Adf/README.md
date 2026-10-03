# OfficeIMO.Adf

`OfficeIMO.Adf` provides an Atlassian Document Format model plus conversions through OfficeIMO's Markdown and HTML engines.

```csharp
AdfDocument document = AdfDocument.Parse(adfJson);
AdfConversionResult<string> markdown = AdfConverter.ToMarkdown(document);
AdfConversionResult<string> html = AdfConverter.ToHtml(document);

AdfConversionResult<AdfDocument> fromMarkdown = AdfConverter.FromMarkdown("# Status\n\nReady.");
```

Unknown ADF nodes, marks, attributes, and extension properties remain in the parsed model and survive JSON round trips. A projection to Markdown or HTML reports unsupported constructs through `AdfConversionReport` instead of silently claiming full fidelity.

Structural validation checks recognized parent/child relationships, required content, panel types, media/caption order, and node-specific mark placement. Unknown node and mark types remain warning-only so native JSON can preserve newer vendor content. Empty native `content`, `marks` and `attrs` properties survive JSON round trips. Validation is a bounded structural check; the [opt-in schema runner](../Build/StructuredFormatVerification/README.md) provides separate evidence against the pinned full Atlassian schema.

Adjacent styled text retains its text and marks through Markdown round trips. Empty inline HTML comments separate delimiter runs when required. Paragraph and list metadata, block marks, unsupported attributes and vendor identities produce omission diagnostics. Mentions and emoji retain their visible labels; cards retain their URL as a link when available. Nodes with no visible fallback report an omission. External media retains its image URL and alternate text, with diagnostics for ADF layout and identity properties.

Standalone Markdown images become external `mediaSingle` nodes. Inline images become linked alternate text and report the lost image semantics. Nested task lists retain their hierarchy and completion state; task identities are regenerated on import. Mixed or more complex Markdown task items that do not fit ADF task-list structure retain visible task markers and a fidelity warning.

`AdfProcessingOptions` bounds JSON input, model nodes/marks, nesting, text and output. `AdfConversionOptions` inherits those limits. Defaults allow 16 MiB input, 64 node levels, 100,000 nodes/marks, 16 Mi characters of text, 32 MiB JSON output and 32 Mi characters of Markdown/HTML output. Markdown object input is checked for depth, object count and cycles before conversion. Limit violations reject the operation without returning truncated content. Cyclic ADF content is rejected. Cancellation is observed at traversal and JSON-write boundaries; synchronous Markdown/HTML parsing is checked before and after its calls.

```csharp
var limits = new AdfProcessingOptions { MaxInputBytes = 2 * 1024 * 1024 };
AdfDocument bounded = AdfDocument.Parse(adfJson, limits);
string json = bounded.ToJson(limits);
```

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 4 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Adf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
