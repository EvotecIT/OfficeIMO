# OfficeIMO.AI.IntelligenceX

This optional .NET 10 adapter connects `OfficeIMO.AI` to IntelligenceX Treatment. It requires `IntelligenceX` 0.1.1 and keeps provider/authentication code out of the document engine.

## ChatGPT

```csharp
using OfficeIMO.AI;
using OfficeIMO.AI.IntelligenceX;

using var executor = await IntelligenceXOfficeAiExecutor.ConnectAsync(
    new OfficeAiExecutionProfile {
        Id = "chatgpt-documents",
        Provider = "ChatGPT",
        Model = "gpt-5.5",
        SupportsImages = true,
        EnforcesJsonSchema = true
    }, new OfficeAiIntelligenceXOptions {
        PreferCurrentCodexSession = true
    });

var engine = new OfficeAiEngine(executor);
```

The native route uses IX's authentication support. `PreferCurrentCodexSession` explicitly chooses the existing local Codex login over an older IX credential when available. The adapter does not redirect the native endpoint or overwrite Codex's authentication file. A hosted operation still requires `OfficeAiRequest.AllowRemoteProcessing = true`.

The example model is an explicit profile setting. If it is unavailable for an account, select an available model and qualify it; the adapter does not silently choose a replacement.

## Compatible HTTP

```csharp
using var executor = await IntelligenceXOfficeAiExecutor.ConnectAsync(
    new OfficeAiExecutionProfile {
        Id = "local-documents",
        Provider = "Local runtime",
        Model = "your-installed-model",
        IsLocal = true,
        SupportsImages = false,
        EnforcesJsonSchema = false
    }, new OfficeAiIntelligenceXOptions {
        Transport = OfficeAiIntelligenceXTransport.CompatibleHttp,
        Endpoint = new Uri("http://127.0.0.1:11434/v1"),
        Streaming = false
    });
```

Use the model identifier and capabilities of the configured runtime. `EnforcesJsonSchema = false` enables prompted JSON with the same local validator and a diagnostic identifying that mode. A hosted compatible provider uses `IsLocal = false`, an HTTPS endpoint, and a caller-supplied `ApiKey` when required. Switching profiles does not change document operations or result schemas. Providers that do not implement this protocol need an executor or an IX transport for their native protocol.

Local profiles require a loopback endpoint. HTTP redirects are disabled; local connections also bypass system proxies. The operator must verify that the service listening on loopback performs inference locally and does not itself forward requests to a hosted service. Transport checks do not certify the server's deployment.

## Isolation and limits

Each Treatment request is ephemeral: it starts fresh and removes local SDK thread state when the request settles. No ambient tool packs, filesystem tools, image-generation tools, or provider model fallback are enabled. Native request/response payload tracing and usage telemetry are disabled. These controls do not assert that a hosted provider deletes its own records.

Inline images, prompt text, model output, and response wire bytes have separate bounds. The wire bound includes SSE overhead. The SDK rejects JSON nesting beyond 128 containers before parsing provider envelopes or candidate output. Cancellation is forwarded to IX, and the document engine suppresses late results. Finish or cancel active operations before disposing the executor; a provider that ignores cancellation may still be running until its task settles.

The [headless example](../Examples/OfficeIMO.AI.Example/README.md) demonstrates authentication selection, local/hosted profiles, PDF/image input, and artifact output. The [engine README](../OfficeIMO.AI/README.md) defines evidence checks, review requirements, and result states.
