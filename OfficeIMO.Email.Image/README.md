# OfficeIMO.Email.Image

`OfficeIMO.Email.Image` renders `EmailDocument` bodies and inline MIME resources through the shared OfficeIMO HTML image pipeline.

```powershell
dotnet add package OfficeIMO.Email.Image
```

The public APIs remain in the `OfficeIMO.Email` namespace:

```csharp
using OfficeIMO.Email;

EmailDocument message = EmailDocument.Load("message.eml");
IReadOnlyList<OfficeImageExportResult> pages = message
    .ToImages(new EmailImageExportOptions())
    .Paged()
    .AsPng()
    .Save("message-pages");
```

The renderer consumes `OfficeIMO.Email.Html` for the same body selection, untrusted sanitization, remote-resource policy, and CID/Content-Location resolution used by Reader. It can convert an RTF body when HTML is absent and safely encodes plain text as the final fallback. Rendering limits, page selection, and diagnostics come from the shared HTML renderer.

`EmailImageExportOptions` also limits inline MIME resources before rendering. `MaxInlineResourceCount` defaults to 128, while `MaxResourceBytes` and `MaxTotalInlineResourceBytes` bound individual and aggregate resource reads.

Keeping this capability separate prevents MIME-only applications from acquiring AngleSharp, CSS layout, and the RTF bridge.

Dependency footprint: `OfficeIMO.Core`, `OfficeIMO.Email`, `OfficeIMO.Email.Html`, and `OfficeIMO.Html`.
