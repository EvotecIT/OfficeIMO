using System;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Ocr;

/// <summary>
/// A validated, immutable identity and capability snapshot for one configured OCR engine instance.
/// </summary>
/// <remarks>
/// Create one execution for a logical document operation and reuse it for every candidate. This prevents
/// mutable provider properties from changing provenance or concurrency behavior partway through the operation.
/// </remarks>
public sealed class OcrEngineExecution {
    private readonly OcrEngineCapabilities _capabilities;

    internal OcrEngineExecution(IOcrEngine engine, string id, OcrEngineCapabilities capabilities) {
        Engine = engine;
        Id = id;
        _capabilities = capabilities;
    }

    /// <summary>Validated provider identifier captured when this execution was created.</summary>
    public string Id { get; }

    /// <summary>Independent copy of the capabilities captured when this execution was created.</summary>
    public OcrEngineCapabilities Capabilities => _capabilities.Clone();

    internal IOcrEngine Engine { get; }

    internal bool SupportsConcurrentRequests => _capabilities.SupportsConcurrentRequests;

    /// <summary>Recognizes one raster payload under the shared timeout and concurrency policy.</summary>
    public Task<OcrResult> RecognizeAsync(
        OcrRequest request,
        TimeSpan timeout,
        CancellationToken cancellationToken = default) =>
        OcrEngineRunner.RecognizeAsync(this, request, timeout, cancellationToken);

    /// <summary>Recognizes a raster payload while bounding the owned result snapshot.</summary>
    public Task<OcrResult> RecognizeAsync(
        OcrRequest request,
        TimeSpan timeout,
        OcrResultCaptureLimits captureLimits,
        CancellationToken cancellationToken) {
        if (captureLimits == null) throw new ArgumentNullException(nameof(captureLimits));
        return OcrEngineRunner.RecognizeAsync(this, request, timeout, cancellationToken, captureLimits);
    }

    /// <summary>Recognizes a bounded payload while keeping the caller attached to provider cleanup.</summary>
    /// <remarks>Cancellation and the deadline request that work stop. Completion waits for every started provider
    /// call and its cancellation callbacks to settle, including after an error. A provider that ignores cancellation
    /// can delay completion indefinitely. Use this route when the host must retain admission or resource ownership
    /// until actual cleanup; ordinary <see cref="RecognizeAsync(OcrRequest, TimeSpan, CancellationToken)"/> retains
    /// its prompt cancellation and timeout behavior.</remarks>
    public Task<OcrResult> RecognizeAttachedAsync(OcrRequest request, TimeSpan timeout,
        OcrResultCaptureLimits captureLimits, CancellationToken cancellationToken = default) {
        if (captureLimits == null) throw new ArgumentNullException(nameof(captureLimits));
        return OcrEngineRunner.RecognizeAsync(this, request, timeout, cancellationToken, captureLimits,
            awaitProviderSettlement: true);
    }
}
