using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.Workflows;

/// <summary>Creates a separate raster-only PDF from selected, prepared scan pages.</summary>
public sealed class OfficeScanCleanupOptions {
    /// <summary>Page selection, sampling density, region cropping, perspective, and tonal settings.</summary>
    public PdfOcrMergeOptions Preparation { get; set; } = new();
    /// <summary>Must be true to acknowledge that only rendered appearances are copied; native text, forms, links, signatures, and attachments are omitted.</summary>
    public bool AcknowledgeRasterOutput { get; set; }
    /// <summary>Optional SHA-256 hex digest from a reviewed source snapshot; a mismatch blocks output.</summary>
    public string? ExpectedSourceSha256 { get; set; }
    internal OfficeScanCleanupOptions Snapshot() {
        if (!AcknowledgeRasterOutput) throw new ArgumentException("Acknowledge the raster-only scan output before saving.");
        if (Preparation == null) throw new ArgumentNullException(nameof(Preparation));
        if (ExpectedSourceSha256 != null && (ExpectedSourceSha256.Length != 64 || ExpectedSourceSha256.Any(character => !Uri.IsHexDigit(character))))
            throw new ArgumentException("The reviewed source digest must be SHA-256 hex.", nameof(ExpectedSourceSha256));
        if (Preparation.DetectOrientation) throw new ArgumentException("Scan-copy preparation requires explicit rotation; provider orientation detection is available during OCR.");
        return new() { Preparation = Preparation.Clone(), AcknowledgeRasterOutput = true, ExpectedSourceSha256 = ExpectedSourceSha256 };
    }
}