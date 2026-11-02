using System;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;

namespace PSWriteOffice.Models.Pdf;

/// <summary>Composes native PDF reading policy with optional OCR-review settings.</summary>
/// <para>OCR settings apply to Read-OfficeDocument -As Ocr. The native policies remain available for direct engine calls.</para>
public sealed class PdfReadOptions {
    /// <summary>Complete native parsing, decryption, resource and semantic-read policy, also used by the OCR read stage.</summary>
    public PdfLoadOptions LoadOptions { get; set; } = new();
    /// <summary>Optional native OCR rendering, recognition and layout policy.</summary>
    public PdfOcrMergeOptions? OcrMergeOptions { get; set; }

    internal PdfReadOptions Snapshot() {
        var policy = (LoadOptions ?? throw new ArgumentException("LoadOptions cannot be null.")).Clone();
        var ocr = OcrMergeOptions?.Clone();
        if (ocr != null) ocr.ReadOptions = policy.Clone();
        return new PdfReadOptions { LoadOptions = policy, OcrMergeOptions = ocr };
    }
}
