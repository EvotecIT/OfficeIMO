# OfficeIMO.Mhtml.Pdf

`OfficeIMO.Mhtml.Pdf` converts bounded MHT/MHTML archives to the first-party OfficeIMO PDF model, including embedded CID and Content-Location resources.

```powershell
dotnet add package OfficeIMO.Mhtml.Pdf
```

```csharp
using OfficeIMO.Html.Pdf;
using OfficeIMO.Mhtml;

MhtmlDocument archive = MhtmlDocument.Load("quarterly-update.mhtml");
PdfDocumentConversionResult result = await archive.ToPdfDocumentResultAsync(
    new HtmlPdfSaveOptions());
await result.SaveAsync("quarterly-update.pdf");
```

The result combines MIME, HTML-rendering, and PDF diagnostics. Local-file and remote-network access remain governed by the HTML resource policy; embedded archive resources do not silently widen it.

Plain HTML/PDF consumers do not receive the Email MIME engine unless they install this bridge.

Dependency footprint: `OfficeIMO.Core`, `OfficeIMO.Mhtml`, `OfficeIMO.Html.Pdf`, and `OfficeIMO.Pdf`.
