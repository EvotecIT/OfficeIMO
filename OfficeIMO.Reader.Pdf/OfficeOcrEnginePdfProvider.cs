using OfficeIMO.Pdf;
using System.Threading.Tasks;

namespace OfficeIMO.Reader.Pdf;

/// <summary>
/// Adapts a reusable <see cref="IOfficeOcrEngine"/> to the rendered-page OCR contract owned by
/// <see cref="OfficeIMO.Pdf"/>.
/// </summary>
public sealed class OfficeOcrEnginePdfProvider : IPdfOcrProvider {
    private readonly IOfficeOcrEngine _engine;
    private readonly OptionsSnapshot _options;

    /// <summary>Creates a PDF OCR provider over an existing Reader OCR engine.</summary>
    public OfficeOcrEnginePdfProvider(IOfficeOcrEngine engine, OfficeOcrEnginePdfProviderOptions? options = null) {
        _engine = engine ?? throw new ArgumentNullException(nameof(engine));
        _options = OptionsSnapshot.Create(options);
    }

    /// <summary>Recognizes one rendered PDF page and projects word or line spans into pixel geometry.</summary>
    public async Task<PdfOcrResponse> RecognizeAsync(PdfOcrRequest request, CancellationToken cancellationToken = default) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        cancellationToken.ThrowIfCancellationRequested();

        byte[] payload = (byte[])request.Png.Clone();
        string assetId = "pdf-page-" + request.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
        var location = new ReaderLocation {
            Path = _options.SourceName,
            Page = request.PageNumber,
            SourceBlockKind = "rendered-pdf-page",
            BlockAnchor = assetId
        };
        var asset = new OfficeDocumentAsset {
            Id = assetId,
            Kind = "image",
            MediaType = "image/png",
            Extension = ".png",
            FileName = assetId + ".png",
            Width = request.PixelWidth,
            Height = request.PixelHeight,
            LengthBytes = payload.LongLength,
            PayloadHash = OfficeDocumentAssetHash.ComputeSha256Hex(payload),
            PayloadBytes = payload,
            Location = location
        };
        var candidate = new OfficeDocumentOcrCandidate {
            Id = "ocr-" + assetId,
            Kind = "page",
            Reason = "Rendered PDF page supplied for OCR.",
            AssetId = assetId,
            Location = location
        };
        var source = new OfficeDocumentSource {
            Path = _options.SourceName,
            SourceId = _options.SourceId,
            LengthBytes = payload.LongLength
        };
        OfficeOcrEngineResult result = await _engine.RecognizeAsync(new OfficeOcrEngineRequest {
            Candidate = candidate,
            Asset = asset,
            Payload = (byte[])payload.Clone(),
            Language = _options.Language,
            Source = source,
            ProviderOptions = _options.ProviderOptions
        }, cancellationToken).ConfigureAwait(false) ?? new OfficeOcrEngineResult();

        var diagnostics = new List<string>();
        if (result.Diagnostics != null) {
            foreach (OfficeDocumentDiagnostic diagnostic in result.Diagnostics) {
                if (diagnostic == null) continue;
                diagnostics.Add(string.IsNullOrWhiteSpace(diagnostic.Code)
                    ? diagnostic.Message
                    : diagnostic.Code + ": " + diagnostic.Message);
            }
        }

        IReadOnlyList<OfficeOcrTextSpan> spans = SelectSpans(result.Spans);
        var words = new List<PdfOcrWord>(spans.Count);
        bool usedFallbackConfidence = false;
        for (int index = 0; index < spans.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeOcrTextSpan span = spans[index];
            if (string.IsNullOrWhiteSpace(span.Text) || span.Region == null) continue;
            if (span.PageNumber.HasValue && span.PageNumber.Value != request.PageNumber && span.PageNumber.Value != 1) continue;
            if (!TryConvertRegion(span.Region, span.CoordinateUnit, request, out double x, out double y, out double width, out double height)) {
                diagnostics.Add("ocr-span-geometry: A recognized span had unsupported or non-finite geometry.");
                continue;
            }

            double confidence;
            if (span.Confidence.HasValue) {
                confidence = span.Confidence.Value;
            } else if (result.Confidence.HasValue) {
                confidence = result.Confidence.Value;
            } else {
                confidence = _options.ConfidenceWhenUnavailable;
                usedFallbackConfidence = true;
            }
            words.Add(new PdfOcrWord(span.Text, x, y, width, height, confidence));
        }

        if (spans.Count == 0 && !string.IsNullOrWhiteSpace(result.Text)) {
            diagnostics.Add("ocr-span-geometry-missing: The OCR engine returned text without word or line geometry, so it could not be placed on the PDF page.");
        }
        if (usedFallbackConfidence) {
            diagnostics.Add("ocr-confidence-unavailable: The OCR engine did not report confidence; the configured fallback confidence was used.");
        }

        return new PdfOcrResponse(
            words,
            diagnostics,
            provider: string.IsNullOrWhiteSpace(result.Provider) ? _engine.Id : result.Provider,
            model: result.Model,
            language: result.Language ?? _options.Language);
    }

    private IReadOnlyList<OfficeOcrTextSpan> SelectSpans(IReadOnlyList<OfficeOcrTextSpan>? source) {
        OfficeOcrTextSpan[] spans = (source ?? Array.Empty<OfficeOcrTextSpan>())
            .Where(static span => span != null)
            .OrderBy(static span => span.Sequence)
            .ToArray();
        OfficeOcrTextSpan[] words = spans.Where(static span => span.Level == OfficeOcrTextSpanLevel.Word).ToArray();
        if (words.Length > 0 || !_options.UseLineSpansWhenWordsUnavailable) return words;
        return spans.Where(static span => span.Level == OfficeOcrTextSpanLevel.Line).ToArray();
    }

    private static bool TryConvertRegion(
        OfficeDocumentRegion region,
        OfficeOcrCoordinateUnit unit,
        PdfOcrRequest request,
        out double x,
        out double y,
        out double width,
        out double height) {
        x = region.X;
        y = region.Y;
        width = region.Width;
        height = region.Height;
        switch (unit) {
            case OfficeOcrCoordinateUnit.Pixels:
                break;
            case OfficeOcrCoordinateUnit.Points:
                x *= request.Scale;
                y *= request.Scale;
                width *= request.Scale;
                height *= request.Scale;
                break;
            case OfficeOcrCoordinateUnit.Normalized:
                x *= request.PixelWidth;
                y *= request.PixelHeight;
                width *= request.PixelWidth;
                height *= request.PixelHeight;
                break;
            default:
                return false;
        }
        return IsFinite(x) && IsFinite(y) && IsFinite(width) && IsFinite(height);
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private sealed class OptionsSnapshot {
        internal string? Language { get; private set; }
        internal string? SourceName { get; private set; }
        internal string? SourceId { get; private set; }
        internal bool UseLineSpansWhenWordsUnavailable { get; private set; }
        internal double ConfidenceWhenUnavailable { get; private set; }
        internal IReadOnlyDictionary<string, string> ProviderOptions { get; private set; } = new Dictionary<string, string>(StringComparer.Ordinal);

        internal static OptionsSnapshot Create(OfficeOcrEnginePdfProviderOptions? options) {
            OfficeOcrEnginePdfProviderOptions source = options ?? new OfficeOcrEnginePdfProviderOptions();
            if (source.ConfidenceWhenUnavailable < 0D || source.ConfidenceWhenUnavailable > 1D ||
                double.IsNaN(source.ConfidenceWhenUnavailable) || double.IsInfinity(source.ConfidenceWhenUnavailable)) {
                throw new ArgumentOutOfRangeException(nameof(options), "Fallback confidence must be finite and between zero and one.");
            }
            return new OptionsSnapshot {
                Language = string.IsNullOrWhiteSpace(source.Language) ? null : source.Language!.Trim(),
                SourceName = string.IsNullOrWhiteSpace(source.SourceName) ? null : source.SourceName,
                SourceId = string.IsNullOrWhiteSpace(source.SourceId) ? null : source.SourceId,
                UseLineSpansWhenWordsUnavailable = source.UseLineSpansWhenWordsUnavailable,
                ConfidenceWhenUnavailable = source.ConfidenceWhenUnavailable,
                ProviderOptions = source.ProviderOptions == null
                    ? new Dictionary<string, string>(StringComparer.Ordinal)
                    : source.ProviderOptions.ToDictionary(static pair => pair.Key, static pair => pair.Value, StringComparer.Ordinal)
            };
        }
    }
}

/// <summary>Configures projection from a Reader OCR engine into PDF page OCR.</summary>
public sealed class OfficeOcrEnginePdfProviderOptions {
    /// <summary>Requested language or provider-specific language expression.</summary>
    public string? Language { get; set; }

    /// <summary>Optional source path or logical name attached to engine requests.</summary>
    public string? SourceName { get; set; }

    /// <summary>Optional stable source identifier attached to engine requests.</summary>
    public string? SourceId { get; set; }

    /// <summary>Uses line spans when the engine does not return word spans.</summary>
    public bool UseLineSpansWhenWordsUnavailable { get; set; } = true;

    /// <summary>Confidence assigned when neither a span nor the overall result reports confidence.</summary>
    public double ConfidenceWhenUnavailable { get; set; } = 1D;

    /// <summary>Provider-specific scalar options forwarded to the OCR engine.</summary>
    public IReadOnlyDictionary<string, string> ProviderOptions { get; set; } = new Dictionary<string, string>(StringComparer.Ordinal);
}
