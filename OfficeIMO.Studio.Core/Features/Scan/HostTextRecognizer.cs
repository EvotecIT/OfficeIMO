using OfficeIMO.Ocr;

namespace OfficeIMO.Studio.Features.Scan;

/// <summary>
/// Text recognition that belongs to the host, such as Apple's Vision framework on device, as an OCR engine: each page
/// image goes to the host, and the words it returns go through the same searchable-text writer as every OCR provider.
/// </summary>
internal static class HostTextRecognizer {
    internal const string EngineId = "host-recognizer";
    internal const int MaximumWordsPerPage = 20_000;

    private static readonly OcrEngineCapabilities Capabilities = new() { SupportsWordSpans = true, SupportsConfidence = true };

    /// <summary>An engine that asks <paramref name="recognize"/> for each page's words, given its number and PNG image.</summary>
    internal static IOcrEngine Create(Func<int, byte[], CancellationToken, IReadOnlyList<PdfScanWord>> recognize) =>
        new DelegateOcrEngine(EngineId, (request, token) =>
            Task.FromResult(ToResult(recognize(request.PageNumber ?? 1, request.Payload, token))), Capabilities);

    /// <summary>An engine that returns words already recognized, for a page the host read before composing it.</summary>
    internal static IOcrEngine FromWords(IReadOnlyList<PdfScanWord> words) =>
        new DelegateOcrEngine(EngineId, (_, _) => Task.FromResult(ToResult(words)), Capabilities);

    /// <summary>The words with text and a box inside the image, from its top-left corner in fractions of its size.</summary>
    internal static IReadOnlyList<PdfScanWord> Validate(IReadOnlyList<PdfScanWord>? words) {
        if (words is null || words.Count == 0) return [];
        if (words.Count > MaximumWordsPerPage) throw new ArgumentOutOfRangeException(nameof(words), $"A page has at most {MaximumWordsPerPage} words.");
        return words.Where(static word => !string.IsNullOrWhiteSpace(word.Text) &&
            double.IsFinite(word.X) && double.IsFinite(word.Y) && double.IsFinite(word.Width) && double.IsFinite(word.Height) &&
            word.Width > 0D && word.Height > 0D && word.X >= 0D && word.Y >= 0D && word.X + word.Width <= 1.0001D && word.Y + word.Height <= 1.0001D)
            .ToArray();
    }

    private static OcrResult ToResult(IReadOnlyList<PdfScanWord> words) {
        IReadOnlyList<PdfScanWord> valid = Validate(words);
        return new OcrResult {
            Text = string.Join(" ", valid.Select(static word => word.Text)),
            Spans = valid.Select((word, index) => new OcrTextSpan {
                Sequence = index, Level = OcrTextSpanLevel.Word, Text = word.Text, Confidence = word.Confidence ?? 1D,
                Region = new OcrRegion { X = word.X, Y = word.Y, Width = word.Width, Height = word.Height },
                CoordinateUnit = OcrCoordinateUnit.Normalized
            }).ToArray()
        };
    }
}
