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
}
