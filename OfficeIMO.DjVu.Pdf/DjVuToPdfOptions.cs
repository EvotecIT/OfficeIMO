using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;

namespace OfficeIMO.DjVu.Pdf;

/// <summary>Bounded scanned-page PDF conversion, retaining stored text and optionally recognizing missing text.</summary>
public sealed class DjVuToPdfOptions {
    /// <summary>One-based source pages in output order. Null selects all pages; duplicates are rejected.</summary>
    public IReadOnlyList<int>? PageNumbers { get; set; }
    /// <summary>Whether stored or explicitly recognized text is embedded as a searchable layer.</summary>
    public bool IncludeText { get; set; } = true;
    /// <summary>Complete-page raster settings. Native source resolution is the default.</summary>
    public DjVuRenderOptions RenderOptions { get; set; } = new DjVuRenderOptions();
    /// <summary>Existing PDF serialization and conformance settings, copied for each conversion.</summary>
    public PdfOptions? PdfOptions { get; set; }
    /// <summary>Explicit caller-owned engine for absent or empty stored text. Requires an asynchronous conversion entrypoint.</summary>
    public IOcrEngine? OcrEngine { get; set; }
    /// <summary>Existing OCR execution, timeout, and result limits. Present and corrupt stored text are never selected for OCR.</summary>
    public OfficeDocumentOcrExecutionOptions OcrOptions { get; set; } = new OfficeDocumentOcrExecutionOptions();
    /// <summary>Maximum selected pages.</summary>
    public int MaxPages { get; set; } = 512;
    /// <summary>Maximum encoded PNG bytes retained for one page.</summary>
    public long MaxImageBytesPerPage { get; set; } = 32L * 1024 * 1024;
    /// <summary>Maximum aggregate encoded PNG bytes retained for the PDF model.</summary>
    public long MaxTotalImageBytes { get; set; } = 128L * 1024 * 1024;
    /// <summary>Maximum searchable spans across selected pages.</summary>
    public int MaxTextSpans { get; set; } = 200_000;
    /// <summary>Maximum searchable UTF-16 text characters across selected pages.</summary>
    public long MaxTextCharacters { get; set; } = 8L * 1024 * 1024;
    /// <summary>Maximum serialized PDF bytes, enforced by OfficeIMO.Pdf before writing beyond the limit.</summary>
    public long MaxPdfBytes { get; set; } = 256L * 1024 * 1024;

    /// <summary>Creates a validated, independent operation snapshot. The caller-owned engine retains its identity.</summary>
    public DjVuToPdfOptions Clone() {
        var copy = (DjVuToPdfOptions)MemberwiseClone();
        copy.PageNumbers = PageNumbers == null ? null : Array.AsReadOnly(PageNumbers.ToArray());
        copy.RenderOptions = (RenderOptions ?? throw new ArgumentNullException(nameof(RenderOptions))).Clone();
        copy.PdfOptions = PdfOptions?.Clone() ?? new PdfOptions();
        copy.OcrOptions = (OcrOptions ?? throw new ArgumentNullException(nameof(OcrOptions))).Clone();
        if (copy.RenderOptions.Region.HasValue || !copy.RenderOptions.ApplyRotation)
            throw new ArgumentException("PDF conversion requires complete, display-oriented page images.", nameof(RenderOptions));
        if (MaxPages <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPages));
        if (MaxImageBytesPerPage <= 0 || MaxImageBytesPerPage > 128L * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(MaxImageBytesPerPage));
        if (MaxTotalImageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTotalImageBytes));
        if (MaxTextSpans <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTextSpans));
        if (MaxTextCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTextCharacters));
        if (MaxPdfBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPdfBytes));
        if (OcrEngine != null && !IncludeText) throw new ArgumentException("OCR requires IncludeText.", nameof(OcrEngine));
        copy.PdfOptions.MaxGeneratedPages = Math.Min(copy.PdfOptions.MaxGeneratedPages ?? MaxPages, MaxPages);
        copy.PdfOptions.MaxGeneratedOutputBytes = Math.Min(copy.PdfOptions.MaxGeneratedOutputBytes ?? MaxPdfBytes, MaxPdfBytes);
        return copy;
    }
}
