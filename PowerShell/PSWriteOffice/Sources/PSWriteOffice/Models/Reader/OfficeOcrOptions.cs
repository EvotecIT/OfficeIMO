using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Reader;

namespace PSWriteOffice.Models.Reader;

/// <summary>Optional OCR provider and existing native policies used by document reads and OCR reviews.</summary>
public sealed class OfficeOcrOptions {
    /// <summary>Engine-neutral provider. When omitted, the configured local Tesseract session is used.</summary>
    public IOcrEngine? Provider { get; set; }
    /// <summary>Tesseract discovery and language-data provisioning policy.</summary>
    public TesseractOcrSessionOptions SessionOptions { get; set; } = new();
    /// <summary>Bounded recognition and enrichment policy for ordinary Reader views.</summary>
    public OfficeDocumentOcrExecutionOptions ExecutionOptions { get; set; } = new();
    /// <summary>Layout, confidence and resource policy for original image or PDF OCR reviews.</summary>
    public PdfOcrMergeOptions RecognitionOptions { get; set; } = new() { ReconstructLayout = true };
}
