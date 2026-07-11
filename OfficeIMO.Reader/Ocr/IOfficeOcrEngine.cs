using System;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

/// <summary>
/// Recognizes text for one bounded OCR candidate without coupling the Reader core to a specific engine or service SDK.
/// </summary>
public interface IOfficeOcrEngine {
    /// <summary>Stable provider identifier used in diagnostics and enrichment metadata.</summary>
    string Id { get; }

    /// <summary>Capabilities exposed by this engine instance.</summary>
    OfficeOcrEngineCapabilities Capabilities { get; }

    /// <summary>Recognizes text and optional detailed spans for one candidate asset.</summary>
    ValueTask<OfficeOcrEngineResult> RecognizeAsync(OfficeOcrEngineRequest request, CancellationToken cancellationToken = default);
}

/// <summary>
/// Adapts a caller-owned OCR callback, including cloud SDK integrations, to <see cref="IOfficeOcrEngine"/>.
/// </summary>
public sealed class DelegateOfficeOcrEngine : IOfficeOcrEngine {
    private readonly Func<OfficeOcrEngineRequest, CancellationToken, ValueTask<OfficeOcrEngineResult>> _recognizeAsync;
    private readonly OfficeOcrEngineCapabilities _capabilities;

    /// <summary>Creates a callback-backed OCR engine.</summary>
    public DelegateOfficeOcrEngine(
        string id,
        Func<OfficeOcrEngineRequest, CancellationToken, ValueTask<OfficeOcrEngineResult>> recognizeAsync,
        OfficeOcrEngineCapabilities? capabilities = null) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("OCR engine id cannot be empty.", nameof(id));
        Id = id.Trim();
        _recognizeAsync = recognizeAsync ?? throw new ArgumentNullException(nameof(recognizeAsync));
        _capabilities = (capabilities ?? new OfficeOcrEngineCapabilities()).Clone();
    }

    /// <inheritdoc />
    public string Id { get; }

    /// <inheritdoc />
    public OfficeOcrEngineCapabilities Capabilities => _capabilities.Clone();

    /// <inheritdoc />
    public ValueTask<OfficeOcrEngineResult> RecognizeAsync(OfficeOcrEngineRequest request, CancellationToken cancellationToken = default) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        return _recognizeAsync(request, cancellationToken);
    }
}
