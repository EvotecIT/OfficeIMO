# OfficeIMO.Mhtml

`OfficeIMO.Mhtml` loads and saves MHT/MHTML web archives while preserving the root HTML document, MIME resources, base URI, and diagnostics.

```powershell
dotnet add package OfficeIMO.Mhtml
```

```csharp
using OfficeIMO.Html;
using OfficeIMO.Mhtml;

MhtmlDocument archive = MhtmlDocument.Load("snapshot.mhtml");
Console.WriteLine(archive.HtmlDocument.NormalizedHtml);

var renderOptions = new HtmlRenderOptions();
archive.ConfigureRenderOptions(renderOptions);
byte[] png = archive.HtmlDocument.ToPng(renderOptions);

archive.Save("copy.mht");
```

`ConfigureRenderOptions` resolves `cid:` and `Content-Location` references first. Its default policy is offline and never invokes a caller resolver for missing network resources. Remote retrieval is explicit, same-origin by default, redirect-bounded, and still subject to the shared HTML count, byte, timeout, and URL policies:

Chromium MHTML snapshots can include rendered web-component content in `template shadowmode` elements. Configured rendering projects that saved content and its named/default slots into a static render tree without changing the archive source. Ordinary templates remain inert. The result reports `HtmlRenderSerializedShadowRootApproximated`; shadow-scoped stylesheets are omitted with `HtmlRenderSerializedShadowStyleOmitted` so their rules cannot affect unrelated content. Use browser-backed output when exact shadow styling, layout, or live behavior is required.

```csharp
var renderOptions = new HtmlRenderOptions();
MhtmlRemoteResourcePolicy remote = MhtmlRemoteResourcePolicy.CreateSameOriginProfile(maximumRedirects: 2);
remote.ResourceFetcher = async (request, cancellationToken) => {
    // Fetch exactly one response with automatic redirects disabled.
    // Return MhtmlRemoteResourceResponse.Redirect(location) for a 3xx response.
    return await applicationOneHopFetcher(request, cancellationToken);
};
archive.ConfigureRenderOptions(renderOptions, remote);
```

The one-hop fetcher must not automatically follow redirects. OfficeIMO resolves each returned redirect location, enforces scheme, origin, and redirect-count policy, and only then invokes the fetcher for the next hop. This prevents disallowed redirect targets from being contacted before policy evaluation. Duplicate `Content-ID` and resolved `Content-Location` values are deterministic first-wins conditions reported through `MimeDiagnostics`; malformed MIME recovery and legacy charset diagnostics come from the shared bounded Email reader. Script execution remains unsupported.

## Concealed-content inspection and cleanup

`InspectContentSafety` applies the shared bounded HTML/CSS safety model to the root HTML part and resolves linked stylesheets only from embedded MIME resources. It never fetches network or file-system resources.

```csharp
using OfficeIMO.ContentSafety;
using OfficeIMO.Mhtml;

OfficeContentSafetyReport report = MhtmlDocument.InspectContentSafety("snapshot.mhtml");
OfficeContentSafetyFinding finding = report.Findings.Single(item => item.TextPreview.Contains("ignore previous", StringComparison.OrdinalIgnoreCase));

OfficeContentCleanupResult cleaned = MhtmlDocument.RemoveSelectedContent(
    "snapshot.mhtml",
    "snapshot-clean.mhtml",
    new OfficeContentCleanupSelection(new[] { finding.Id }));
```

Cleanup removes only the selected current findings, preserves embedded resources and unrelated nested-message payloads, writes a bounded deterministic archive, and reopens and reinspects the result. HTML roots and embedded stylesheets use the same web-compatible charset aliases, including a bounded HTML encoding declaration when a MIME charset is absent, while rewritten parts retain their declared encoding and original BOM. Embedded-resource URI scheme and host matching is case-insensitive, while path and query matching is case-sensitive. Stale outer and part-level payload length or digest headers are removed. Missing, malformed, ambiguously decoded, active integrity-qualified, external, or over-budget stylesheet dependencies fail closed, as do a `multipart/related` `start` value that is missing or selects a non-HTML root, conflicting preferred titled stylesheet sets, and documents with Content Security Policy declarations that the bounded cascade does not model. Non-CSS or inactive-media stylesheets remain inert. An empty selection returns the original bytes. Mutation of a signed or encrypted MIME wrapper is rejected. Body-covering DKIM and ARC transport signatures also block cleanup by default; callers must explicitly choose `RemoveInvalidatedSignatures` to remove the invalidated transport-signature chain, or `PreserveSignatureMarkup` when retaining known-stale signature headers is intentional.

MHTML intentionally connects the HTML engine to the Email MIME engine. Plain HTML and plain Email packages do not depend on this bridge.

Dependency footprint: `OfficeIMO.Core`, `OfficeIMO.Html`, and `OfficeIMO.Email`.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 1 | 0 | 0 | 0 | 0 | 0 |
| Read | 1 | 0 | 0 | 0 | 0 | 0 |
| Edit | 1 | 0 | 0 | 0 | 0 | 0 |
| Preserve | 0 | 1 | 0 | 0 | 0 | 0 |
| Inspect | 1 | 0 | 0 | 0 | 0 | 0 |
| Validate | 1 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Mhtml` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
