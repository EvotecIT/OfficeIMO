using System;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Ocr;

/// <summary>Recognizes text and layout from one bounded raster payload.</summary>
public interface IOcrEngine {
    /// <summary>
    /// Stable provider identifier used in diagnostics and provenance. The value must remain unchanged for the
    /// lifetime of this configured engine instance, and the untrimmed value must not exceed
    /// <see cref="OcrEngineRunner.MaximumEngineIdCharacters"/> characters.
    /// </summary>
    string Id { get; }

    /// <summary>Stable capabilities exposed by this configured engine instance.</summary>
    OcrEngineCapabilities Capabilities { get; }

    /// <summary>Recognizes text and optional detailed spans from one raster payload.</summary>
    Task<OcrResult> RecognizeAsync(OcrRequest request, CancellationToken cancellationToken = default);
}

/// <summary>Adapts a caller-owned callback, including cloud SDK integrations, to <see cref="IOcrEngine"/>.</summary>
public sealed class DelegateOcrEngine : IOcrEngine {
    private readonly Func<OcrRequest, CancellationToken, Task<OcrResult>> _recognizeAsync;
    private readonly OcrEngineCapabilities _capabilities;

    /// <summary>Creates a callback-backed OCR engine.</summary>
    public DelegateOcrEngine(
        string id,
        Func<OcrRequest, CancellationToken, Task<OcrResult>> recognizeAsync,
        OcrEngineCapabilities? capabilities = null) {
        if (string.IsNullOrEmpty(id)) throw new ArgumentException("OCR engine id cannot be empty.", nameof(id));
        if (id.Length > OcrEngineRunner.MaximumEngineIdCharacters) {
            throw new ArgumentException(
                "OCR engine id cannot exceed " + OcrEngineRunner.MaximumEngineIdCharacters + " characters.",
                nameof(id));
        }
        string normalizedId = id.Trim();
        if (normalizedId.Length == 0) throw new ArgumentException("OCR engine id cannot be empty.", nameof(id));
        Id = normalizedId;
        _recognizeAsync = recognizeAsync ?? throw new ArgumentNullException(nameof(recognizeAsync));
        _capabilities = (capabilities ?? new OcrEngineCapabilities()).Clone();
    }

    /// <inheritdoc />
    public string Id { get; }

    /// <inheritdoc />
    public OcrEngineCapabilities Capabilities => _capabilities.Clone();

    /// <inheritdoc />
    public Task<OcrResult> RecognizeAsync(OcrRequest request, CancellationToken cancellationToken = default) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        return _recognizeAsync(request, cancellationToken);
    }
}
