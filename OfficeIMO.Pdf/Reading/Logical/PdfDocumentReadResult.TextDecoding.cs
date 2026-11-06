namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentReadResult {
    private IReadOnlyList<PdfFontInspectionDiagnostic>? _textDecodingDiagnostics;

    /// <summary>Warnings for undecodable native text runs omitted from this result's Text, Markdown and tables.
    /// Each diagnostic identifies the one-based source page and font at its content-stream path.
    /// Valid ActualText replacements and supported font encodings remain readable.</summary>
    public IReadOnlyList<PdfFontInspectionDiagnostic> TextDecodingDiagnostics =>
        _textDecodingDiagnostics ??= Pages.GroupBy(static page => page.PageNumber)
            .SelectMany(static group => group.First().Analysis.TextDecodingDiagnostics).ToList().AsReadOnly();
}
