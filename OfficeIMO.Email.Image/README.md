# OfficeIMO.Email.Image

`OfficeIMO.Email.Image` renders `EmailDocument` bodies and inline MIME resources through the shared OfficeIMO HTML image pipeline.

```powershell
dotnet add package OfficeIMO.Email.Image
```

The public APIs remain in the `OfficeIMO.Email` namespace:

```csharp
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Email;
using OfficeIMO.Html;

EmailDocument message = EmailDocument.Load("message.eml");
var imageOptions1 = new EmailImageExportOptions();
imageOptions1.Mode = HtmlRenderMode.Paged;
imageOptions1.Format = OfficeImageExportFormat.Png;
OfficeConversionReport pages = message.SaveAsImages("message-pages", imageOptions1);
```

The renderer consumes `OfficeIMO.Email.Html` for the same body selection, untrusted sanitization, remote-resource policy, and CID/Content-Location resolution used by Reader. It can convert an RTF body when HTML is absent and safely encodes plain text as the final fallback. Rendering limits, page selection, and diagnostics come from the shared HTML renderer.

`EmailImageExportOptions` also limits inline MIME resources before rendering. `MaxInlineResourceCount` defaults to 128, while `MaxResourceBytes` and `MaxTotalInlineResourceBytes` bound individual and aggregate resource reads.

Keeping this capability separate prevents MIME-only applications from acquiring AngleSharp, CSS layout, and the RTF bridge.

Dependency footprint: `OfficeIMO.Core`, `OfficeIMO.Email`, `OfficeIMO.Email.Html`, and `OfficeIMO.Html`.

`SaveAsImages(folder, options)` and its async form return `OfficeConversionReport`. The options carry the output format, selection, deadlines, resource limits and file conflict policy. Use the `ExportImages` consumer overloads to process each image with its sequence metadata and diagnostics.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Export | 0 | 5 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Email.Image` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
