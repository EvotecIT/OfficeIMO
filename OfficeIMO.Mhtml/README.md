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

`ConfigureRenderOptions` resolves `cid:` and `Content-Location` references before falling back to a caller resolver. It does not enable network or local-file access beyond the configured HTML resource policy.

MHTML intentionally connects the HTML engine to the Email MIME engine. Plain HTML and plain Email packages do not depend on this bridge.

Dependency footprint: `OfficeIMO.Core`, `OfficeIMO.Html`, and `OfficeIMO.Email`.
