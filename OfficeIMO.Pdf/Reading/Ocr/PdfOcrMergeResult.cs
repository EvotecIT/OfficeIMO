namespace OfficeIMO.Pdf;

/// <summary>Normalized OCR merge result with both native and OCR-enriched logical documents.</summary>
public sealed class PdfOcrMergeResult {
    internal PdfOcrMergeResult(PdfDocumentReadResult nativeDocument, PdfDocumentReadResult enrichedDocument, IReadOnlyList<PdfOcrPageMergeResult> pages) {
        NativeDocument = nativeDocument;
        EnrichedDocument = enrichedDocument;
        Pages = pages;
    }
    /// <summary>Native parser logical model used for overlap decisions.</summary>
    public PdfDocumentReadResult NativeDocument { get; }
    /// <summary>
    /// Logical model containing accepted OCR text and conservative OCR table inference in addition to native content.
    /// Pass this model directly to the Word, Excel, PowerPoint, HTML, RTF, or OpenDocument reverse converters.
    /// </summary>
    public PdfDocumentReadResult EnrichedDocument { get; }
    /// <summary>Number of accepted OCR words across all requested pages.</summary>
    public int AcceptedWordCount => Pages.Sum(static page => page.Words.Count);
    /// <summary>
    /// True when at least one OCR word passed merge filtering. This remains merge evidence when
    /// <see cref="PdfOcrMergeOptions.BuildEnrichedLogicalDocument"/> disables logical-model projection.
    /// </summary>
    public bool HasAcceptedOcrContent => AcceptedWordCount > 0;
    /// <summary>OCR merge reports in requested page order.</summary>
    public IReadOnlyList<PdfOcrPageMergeResult> Pages { get; }
    /// <summary>Combined page text separated by blank lines.</summary>
    public string Text => string.Join(Environment.NewLine + Environment.NewLine, Pages.Select(static page => page.Text));
}

/// <summary>Accepted OCR words and evidence for one page.</summary>
public sealed class PdfOcrPageMergeResult {
    internal PdfOcrPageMergeResult(int pageNumber, IReadOnlyList<PdfRecognizedWord> words, int rejectedLowConfidenceCount, int rejectedNativeOverlapCount, IReadOnlyList<string> diagnostics, string text, string? provider = null, string? model = null, string? language = null) {
        PageNumber = pageNumber; Words = words; RejectedLowConfidenceCount = rejectedLowConfidenceCount; RejectedNativeOverlapCount = rejectedNativeOverlapCount; Diagnostics = diagnostics; Text = text;
        Provider = provider; Model = model; Language = language;
    }
    /// <summary>One-based page number.</summary>
    public int PageNumber { get; }
    /// <summary>Accepted normalized OCR words.</summary>
    public IReadOnlyList<PdfRecognizedWord> Words { get; }
    /// <summary>Words rejected below confidence threshold.</summary>
    public int RejectedLowConfidenceCount { get; }
    /// <summary>Words rejected because native PDF text already covers the region.</summary>
    public int RejectedNativeOverlapCount { get; }
    /// <summary>Provider and normalization diagnostics.</summary>
    public IReadOnlyList<string> Diagnostics { get; }
    /// <summary>Native and accepted OCR text in approximate visual order.</summary>
    public string Text { get; }
    /// <summary>OCR provider identifier reported for this page, when available.</summary>
    public string? Provider { get; }
    /// <summary>OCR model or trained-data identifier reported for this page, when available.</summary>
    public string? Model { get; }
    /// <summary>Detected or requested OCR language reported for this page, when available.</summary>
    public string? Language { get; }
}

/// <summary>OCR word normalized to top-left visual PDF-point coordinates after crop and page rotation.</summary>
public sealed class PdfRecognizedWord {
    internal PdfRecognizedWord(string text, double x, double y, double width, double height, double confidence, int providerSequence) {
        Text = text; X = x; Y = y; Width = width; Height = height; Confidence = confidence; ProviderSequence = providerSequence;
    }
    /// <summary>Recognized text.</summary>
    public string Text { get; }
    /// <summary>Left coordinate in visual PDF points.</summary>
    public double X { get; }
    /// <summary>Top coordinate in visual PDF points.</summary>
    public double Y { get; }
    /// <summary>Width in PDF points.</summary>
    public double Width { get; }
    /// <summary>Height in PDF points.</summary>
    public double Height { get; }
    /// <summary>Provider confidence.</summary>
    public double Confidence { get; }
    /// <summary>Original logical position in the provider response.</summary>
    internal int ProviderSequence { get; }
}
