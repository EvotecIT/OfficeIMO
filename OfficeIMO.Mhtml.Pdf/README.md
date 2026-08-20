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
    new HtmlPdfSaveOptions());
await result.SaveAsync("quarterly-update.pdf");
```

The result combines MIME, HTML-rendering, and PDF diagnostics. Local-file and remote-network access remain governed by the HTML resource policy; embedded archive resources do not silently widen it.

Conversion is offline by default. To allow missing archive resources through a host resolver, apply an explicit bounded MHTML policy to the same options before conversion. The resolver must report its final URI and redirect count so cross-origin or excessive redirects fail closed:

```csharp
var options = new HtmlPdfSaveOptions {
    ResourcePolicy = PdfResourcePolicy.CreateTrustedHost(),
    ResourceResolver = applicationResolver
};
archive.ConfigureRenderOptions(
    options,
    MhtmlRemoteResourcePolicy.CreateSameOriginProfile(maximumRedirects: 2));

PdfDocumentConversionResult result = await archive.ToPdfDocumentResultAsync(options);
```

Malformed multipart recovery, legacy charset decoding, nested related parts, duplicate Content-ID and Content-Location selection, and archive diagnostics are owned by the bounded Email/MHTML layer. Layout and PDF paint reuse the managed HTML/CSS renderer, and scripts remain inert.

Plain HTML/PDF consumers do not receive the Email MIME engine unless they install this bridge.

Dependency footprint: `OfficeIMO.Core`, `OfficeIMO.Mhtml`, `OfficeIMO.Html.Pdf`, and `OfficeIMO.Pdf`.
