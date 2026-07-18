namespace OfficeIMO.Pdf;

/// <summary>
/// Canonical text and structured-content projection helpers over <see cref="PdfReadDocument"/>.
/// </summary>
internal static partial class PdfTextExtractor {
    private static readonly char[] CsvQuoteChars = new[] { ',', '"', '\r', '\n' };
}
