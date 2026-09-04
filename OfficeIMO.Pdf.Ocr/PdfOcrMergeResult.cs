using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Ocr;

/// <summary>Normalized OCR merge result with native overlap evidence and the parsed document.</summary>
public sealed class PdfOcrMergeResult {
    internal PdfOcrMergeResult(PdfDocumentReadResult nativeDocument, PdfDocumentReadResult document, IReadOnlyList<PdfOcrPageMergeResult> pages) {
        NativeDocument = nativeDocument;
        Document = document;
        Pages = pages;
    }
    /// <summary>Native parser logical model used for overlap decisions.</summary>
    public PdfDocumentReadResult NativeDocument { get; }
    /// <summary>
    /// Canonically parsed logical model containing accepted OCR evidence in addition to native content.
    /// Pass this model directly to the Word, Excel, PowerPoint, HTML, RTF, or OpenDocument reverse converters.
    /// </summary>
    public PdfDocumentReadResult Document { get; }
    /// <summary>Number of accepted OCR words across all requested pages.</summary>
    public int AcceptedWordCount => Pages.Sum(static page => page.Words.Count);
    /// <summary>
    /// True when at least one OCR word passed merge filtering and was supplied to the parsed document.
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

    internal PdfOcrPageMergeResult WithCanonicalText(string text, int maximumCharacters) {
        Guard.NotNull(text, nameof(text));
        if (text.Length > maximumCharacters) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, maximumCharacters, text.Length);
        }
        return new PdfOcrPageMergeResult(
            PageNumber,
            Words,
            RejectedLowConfidenceCount,
            RejectedNativeOverlapCount,
            Diagnostics,
            text,
            Provider,
            Model,
            Language);
    }
}

/// <summary>OCR word normalized to top-left visual PDF-point coordinates after crop and page rotation.</summary>
public sealed class PdfRecognizedWord {
    internal PdfRecognizedWord(
        string text,
        double x,
        double y,
        double width,
        double height,
        double confidence,
        int providerSequence,
        string? blockId = null,
        string? paragraphId = null,
        string? lineId = null) {
        Text = text; X = x; Y = y; Width = width; Height = height; Confidence = confidence; ProviderSequence = providerSequence;
        BlockId = blockId; ParagraphId = paragraphId; LineId = lineId;
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
    /// <summary>Provider block identifier, when available.</summary>
    public string? BlockId { get; }
    /// <summary>Provider paragraph identifier, when available.</summary>
    public string? ParagraphId { get; }
    /// <summary>Provider line identifier, when available.</summary>
    public string? LineId { get; }
    /// <summary>Original logical position in the provider response.</summary>
    internal int ProviderSequence { get; }
}
