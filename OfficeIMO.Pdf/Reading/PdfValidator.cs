namespace OfficeIMO.Pdf;

/// <summary>
/// Wrapper-friendly PDF validation helpers backed by the OfficeIMO.Pdf preflight engine.
/// </summary>
internal static class PdfValidator {
    /// <summary>
    /// Validates a PDF from a byte array without throwing for malformed PDF content.
    /// </summary>
    public static PdfValidationResult Validate(byte[] pdf, PdfReadOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        return new PdfValidationResult(PdfInspector.Preflight(pdf, options));
    }

    /// <summary>
    /// Validates a PDF from a file path without throwing for malformed PDF content.
    /// </summary>
    public static PdfValidationResult Validate(string path, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return Validate(File.ReadAllBytes(path), options);
    }

    /// <summary>
    /// Validates a PDF from the current position of a readable stream without throwing for malformed PDF content.
    /// </summary>
    public static PdfValidationResult Validate(Stream stream, PdfReadOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Validate(buffer.ToArray(), options);
    }
}
