namespace OfficeIMO.Pdf;

/// <summary>Normalized OCR merge result alongside the unchanged native logical document.</summary>
public sealed class PdfOcrMergeResult {
    internal PdfOcrMergeResult(PdfLogicalDocument nativeDocument, IReadOnlyList<PdfOcrPageMergeResult> pages) {
        NativeDocument = nativeDocument;
        Pages = pages;
    }
    /// <summary>Native parser logical model used for overlap decisions.</summary>
    public PdfLogicalDocument NativeDocument { get; }
    /// <summary>OCR merge reports in requested page order.</summary>
    public IReadOnlyList<PdfOcrPageMergeResult> Pages { get; }
    /// <summary>Combined page text separated by blank lines.</summary>
    public string Text => string.Join(Environment.NewLine + Environment.NewLine, Pages.Select(static page => page.Text));
}

/// <summary>Accepted OCR words and evidence for one page.</summary>
public sealed class PdfOcrPageMergeResult {
    internal PdfOcrPageMergeResult(int pageNumber, IReadOnlyList<PdfRecognizedWord> words, int rejectedLowConfidenceCount, int rejectedNativeOverlapCount, IReadOnlyList<string> diagnostics, string text) {
        PageNumber = pageNumber; Words = words; RejectedLowConfidenceCount = rejectedLowConfidenceCount; RejectedNativeOverlapCount = rejectedNativeOverlapCount; Diagnostics = diagnostics; Text = text;
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
}

/// <summary>OCR word normalized to top-left PDF-point coordinates.</summary>
public sealed class PdfRecognizedWord {
    internal PdfRecognizedWord(string text, double x, double y, double width, double height, double confidence) {
        Text = text; X = x; Y = y; Width = width; Height = height; Confidence = confidence;
    }
    /// <summary>Recognized text.</summary>
    public string Text { get; }
    /// <summary>Left coordinate in PDF points.</summary>
    public double X { get; }
    /// <summary>Top coordinate in PDF points.</summary>
    public double Y { get; }
    /// <summary>Width in PDF points.</summary>
    public double Width { get; }
    /// <summary>Height in PDF points.</summary>
    public double Height { get; }
    /// <summary>Provider confidence.</summary>
    public double Confidence { get; }
}
