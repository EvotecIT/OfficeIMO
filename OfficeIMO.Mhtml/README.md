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

MHTML intentionally connects the HTML engine to the Email MIME engine. Plain HTML and plain Email packages do not depend on this bridge.

Dependency footprint: `OfficeIMO.Core`, `OfficeIMO.Html`, and `OfficeIMO.Email`.
