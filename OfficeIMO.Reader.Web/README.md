# OfficeIMO.Reader.Web

`OfficeIMO.Reader.Web` is an explicit, bounded HTTP transport for an existing `OfficeDocumentReader`. It downloads bytes with a caller-owned `HttpClient`, then passes those bytes to the same configured handlers and processors used for local files.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Web;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddHtmlHandler()
    .Build();

var webReader = reader.CreateWebReader(
    httpClient,
    new ReaderWebOptions {
        AllowedHosts = new[] { "docs.example.com" },
        MaxResponseBytes = 16L * 1024L * 1024L
    });

OfficeDocumentReadResult result = await webReader.ReadDocumentAsync(
    new Uri("https://docs.example.com/guide.html"));
```

The transport accepts absolute HTTP(S) GET targets only, rejects URI-embedded credentials, checks an optional exact host allowlist, blocks obvious loopback/private/non-routable literal targets by default, caps response bytes, applies a request timeout, and bounds concurrent operations per web-reader instance. Query strings are omitted from result metadata unless explicitly enabled.

Format selection stays with Reader. The logical source name comes from an explicit `sourceName`, the response `Content-Disposition` filename, or the final URI path, in that order. Supply `sourceName` when a download URL has no usable extension and content detection cannot identify the intended modular handler.

The injected `HttpClient` owns authentication, proxies, certificates, DNS resolution, automatic redirects, decompression, retries, and connection policy. Because an already-configured handler can follow redirects before returning a response, server hosts must enforce DNS and redirect destinations in that handler when SSRF isolation matters. Reader Web validates the requested URI and the final URI reported by the response, but does not claim to intercept intermediate redirects.

This package is not registered by `OfficeIMO.Reader.All` and performs no implicit network access. It adds no HTTP SDK, HTML parser, browser, process, native binary, model, or cloud provider; `System.Net.Http` comes from the target framework.
