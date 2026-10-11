using OfficeIMO.Reader;

namespace OfficeIMO.DjVu.Pdf;

/// <summary>Origin of the PDF's searchable layer, distinct from the source page's stored text state.</summary>
public enum DjVuPdfTextSource {
    /// <summary>No searchable text was embedded.</summary>
    None,
    /// <summary>The original stored DjVu text was reused.</summary>
    StoredText,
    /// <summary>An explicitly supplied engine recognized a page with absent or empty stored text.</summary>
    Ocr
}

/// <summary>Physical geometry and text provenance for one exported scanned page.</summary>
public sealed class DjVuPdfPageReport {
    internal DjVuPdfPageReport(int sourcePage, int outputPage, DjVuTextStatus status, DjVuPdfTextSource textSource,
        int imageWidth, int imageHeight, double width, double height, long textCharacters, OfficeDocumentRecognitionEvidence? recognition) {
        SourcePageNumber = sourcePage; OutputPageNumber = outputPage; StoredTextStatus = status; TextSource = textSource;
        ImageWidth = imageWidth; ImageHeight = imageHeight; WidthPoints = width; HeightPoints = height;
        TextCharacters = textCharacters; Recognition = recognition;
    }
    /// <summary>One-based source page number.</summary>
    public int SourcePageNumber { get; }
    /// <summary>One-based output page number.</summary>
    public int OutputPageNumber { get; }
    /// <summary>State of the original stored text layer.</summary>
    public DjVuTextStatus StoredTextStatus { get; }
    /// <summary>Origin of the searchable text.</summary>
    public DjVuPdfTextSource TextSource { get; }
    /// <summary>Rendered image width in pixels.</summary>
    public int ImageWidth { get; }
    /// <summary>Rendered image height in pixels.</summary>
    public int ImageHeight { get; }
    /// <summary>Display-oriented page width in PDF points, derived from native source DPI.</summary>
    public double WidthPoints { get; }
    /// <summary>Display-oriented page height in PDF points, derived from native source DPI.</summary>
    public double HeightPoints { get; }
    /// <summary>UTF-16 text characters embedded on this page.</summary>
    public long TextCharacters { get; }
    /// <summary>Engine provenance supplied for new OCR. Stored text has no invented recognition provenance.</summary>
    public OfficeDocumentRecognitionEvidence? Recognition { get; }
}

/// <summary>DjVu-stage fidelity and provenance, carried in the PDF result's source conversion reports.</summary>
public sealed class DjVuPdfConversionReport : IOfficeConversionReport {
    internal DjVuPdfConversionReport(string hash, List<DjVuPdfPageReport> pages,
        List<OfficeConversionFidelityDiagnostic> diagnostics, OfficeDocumentOcrExecutionReport? ocrReport) {
        SourceSha256 = hash; Pages = pages.AsReadOnly(); FidelityDiagnostics = diagnostics.AsReadOnly(); OcrReport = ocrReport;
    }
    /// <summary>Hash of the exact primary DjVu source snapshot.</summary>
    public string SourceSha256 { get; }
    /// <summary>Exported pages in output order.</summary>
    public IReadOnlyList<DjVuPdfPageReport> Pages { get; }
    /// <summary>Existing execution counters when an explicit OCR engine ran.</summary>
    public OfficeDocumentOcrExecutionReport? OcrReport { get; }
    /// <inheritdoc />
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics { get; }
    /// <inheritdoc />
    public bool HasLoss => FidelityDiagnostics.Any(d => d.LossKind != OfficeConversionLossKind.None);
    /// <inheritdoc />
    public void RequireNoLoss() {
        if (HasLoss) throw new OfficeConversionException("DjVu PDF conversion reported fidelity loss. Inspect FidelityDiagnostics.", this);
    }
}
