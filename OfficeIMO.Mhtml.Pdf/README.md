# OfficeIMO.Mhtml.Pdf

`OfficeIMO.Mhtml.Pdf` converts bounded MHT/MHTML archives to the first-party OfficeIMO PDF model, including embedded CID and Content-Location resources.

```powershell
dotnet add package OfficeIMO.Mhtml.Pdf
```

```csharp
using OfficeIMO.Html.Pdf;
using OfficeIMO.Mhtml;
using OfficeIMO.Pdf;

MhtmlDocument archive = MhtmlDocument.Load("quarterly-update.mhtml");
PdfDocumentConversionResult result = await archive.ToPdfDocumentResultAsync(
    new HtmlToPdfOptions());
await result.SaveAsync("quarterly-update.pdf");
```

The result combines MIME, HTML-rendering, and PDF diagnostics. Local-file and remote-network access remain governed by the HTML resource policy; embedded archive resources do not silently widen it.

Saved `template shadowmode` component content is projected into the static PDF, including slot-assigned text. The report diagnoses the approximation and any omitted shadow-scoped stylesheets; browser-backed printing remains the path for exact component styling or live behavior. Ordinary templates stay inert.

Conversion is offline by default. To allow missing archive resources, apply an explicit bounded MHTML policy to the same options before conversion. The application fetcher must return exactly one response with automatic redirects disabled so OfficeIMO can approve every redirect target before requesting it:

```csharp
var options = new HtmlToPdfOptions {
    ResourcePolicy = PdfResourcePolicy.CreateTrustedHost()
};
MhtmlRemoteResourcePolicy remote = MhtmlRemoteResourcePolicy.CreateSameOriginProfile(maximumRedirects: 2);
remote.ResourceFetcher = async (request, cancellationToken) => {
    // Return MhtmlRemoteResourceResponse.Redirect(location) for a 3xx response.
    return await applicationOneHopFetcher(request, cancellationToken);
};
archive.ConfigureRenderOptions(options, remote);

PdfDocumentConversionResult result = await archive.ToPdfDocumentResultAsync(options);
```

Malformed multipart recovery, legacy charset decoding, nested related parts, duplicate Content-ID and Content-Location selection, and archive diagnostics are owned by the bounded Email/MHTML layer. Layout and PDF paint reuse the managed HTML/CSS renderer, and scripts remain inert.

Plain HTML/PDF consumers do not receive the Email MIME engine unless they install this bridge.

Dependency footprint: `OfficeIMO.Core`, `OfficeIMO.Mhtml`, `OfficeIMO.Html.Pdf`, and `OfficeIMO.Pdf`.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Mhtml.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
