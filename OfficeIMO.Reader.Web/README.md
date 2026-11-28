# OfficeIMO.Reader.Web

`OfficeIMO.Reader.Web` is an explicit, bounded HTTP transport for an existing `OfficeDocumentReader`. It downloads bytes, then passes them to the same configured handlers and processors used for local files.

## Install

```sh
dotnet add package OfficeIMO.Reader.Web
dotnet add package OfficeIMO.Reader.Html   # the handler the example uses
```

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Web;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddHtmlHandler()
    .Build();

OfficeDocumentWebReader webReader = reader.CreateWebReader(new ReaderWebOptions {
    AllowedHosts = new[] { "docs.example.com" },
    MaxResponseBytes = 16L * 1024L * 1024L
});

OfficeDocumentReadResult result = await webReader.ReadDocumentAsync(
    new Uri("https://docs.example.com/guide.html"));
```

The transport accepts absolute HTTP(S) GET targets only, rejects URI-embedded credentials, checks an optional exact host allowlist, caps response bytes at the strictest Web, Reader, or selected-handler limit, applies a request timeout, and bounds concurrent operations per web-reader instance. Query strings are omitted from result metadata unless explicitly enabled.

Loopback, private (including 100.64.0.0/10), link-local (including cloud metadata at 169.254.169.254), unique-local IPv6 and other non-public addresses are refused by default. The web reader created without an `HttpClient` resolves every host name, checks every address and refuses the request when any of them is non-public. On .NET 8 and later the check runs when the connection is made and the socket connects to the checked address, so a second DNS answer cannot redirect it. Older frameworks check the resolved addresses before each request, send plain HTTP to the checked address and rely on certificate validation to bind HTTPS connections to the host name. That reader follows redirects itself, at most ten, and checks each destination the same way; a redirect from HTTPS to HTTP is refused. While non-public addresses are refused it connects directly, without a proxy. Set `AllowPrivateNetworkAddresses = true` only for trusted internal sources.

Source IDs hash the complete final URI, including its query, while query metadata remains redacted by default.
Chunk IDs include the source ID so distinct downloads with the same filename remain distinct.
For signed URLs whose query changes without changing the document, set `ReaderWebOptions.SourceKeySelector`
to an application-owned canonical key, for example `uri => uri.GetLeftPart(UriPartial.Path)` when the path
alone identifies that document. Do not remove content-selecting query parameters from the canonical key.

Format selection stays with Reader. The logical source name comes from an explicit `sourceName`, the response `Content-Disposition` filename, or the final URI path, in that order. Supply `sourceName` when a download URL has no usable extension and content detection cannot identify the intended modular handler.

`CreateWebReader(httpClient, options)` uses a caller-owned `HttpClient` instead. That client's handler owns DNS resolution, connection establishment, proxies, authentication, certificates, decompression, retries and automatic redirects, so Reader Web can screen only IP literals and localhost names for it, plus each redirect that reaches it and the final URI the response reports. For untrusted URIs, use the reader without a client, or configure the handler to check resolved addresses at connection time and not to follow redirects itself.

This package is not registered by `OfficeIMO.Reader.All` and performs no implicit network access. It adds no HTTP SDK, HTML parser, browser, process, native binary, model, or cloud provider; `System.Net.Http` comes from the target framework.

