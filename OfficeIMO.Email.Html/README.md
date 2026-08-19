# OfficeIMO.Email.Html

`OfficeIMO.Email.Html` is the optional HTML bridge for `OfficeIMO.Email`. It selects an HTML, RTF, or plain-text body once, applies the shared untrusted HTML policy, indexes CID and content-location resources, and exposes a prepared `HtmlConversionDocument` to renderers and readers.

```powershell
dotnet add package OfficeIMO.Email.Html
```

```csharp
using OfficeIMO.Email;

EmailBodyProjectionResult projection = EmailBodyProjection.Create(message);
string safeHtml = projection.Html;
EmailBodyResource? logo = projection.ResolveResource("cid:logo@example.test");
```

Remote resources are blocked by default. Selecting `AllowByConsumerResolver` only retains eligible HTTP(S) references; this package never downloads them. Attachment content remains operation-scoped and is opened only through bounded resource reads.

`OfficeIMO.Email.Image` uses the prepared HTML and resource index for rendering. `OfficeIMO.Reader.Email` uses the same projection before producing safe text or Markdown, so those adapters do not choose bodies, sanitize markup, or resolve embedded resources independently.

The core `OfficeIMO.Email` package does not depend on HTML libraries. Install this bridge only when safe HTML, RTF fallback, resource resolution, rendering, or Markdown projection is needed. The dependency footprint is `OfficeIMO.Email`, `OfficeIMO.Html`, and `OfficeIMO.Html.Rtf`.
