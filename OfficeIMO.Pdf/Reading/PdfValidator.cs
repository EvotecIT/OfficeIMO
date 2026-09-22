namespace OfficeIMO.Pdf;

/// <summary>
/// Wrapper-friendly PDF validation helpers backed by the OfficeIMO.Pdf preflight engine.
/// </summary>
internal static class PdfValidator {
    /// <summary>
    /// Validates a PDF from a byte array without throwing for malformed PDF content.
    /// </summary>
    public static PdfValidationResult Validate(byte[] pdf, PdfLoadOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        return new PdfValidationResult(PdfInspector.Preflight(pdf, options));
    }

    /// <summary>
    /// Validates a PDF from a file path without throwing for malformed PDF content.
    /// </summary>
    public static PdfValidationResult Validate(string path, PdfLoadOptions? options = null) {
        return new PdfValidationResult(PdfInspector.Preflight(path, options));
    }

    /// <summary>
    /// Validates a PDF from the current position of a readable stream without throwing for malformed PDF content.
    /// </summary>
    public static PdfValidationResult Validate(Stream stream, PdfLoadOptions? options = null) {
        return new PdfValidationResult(PdfInspector.Preflight(stream, options));
    }
}
