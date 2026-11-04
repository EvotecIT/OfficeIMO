using OfficeIMO.Html;

namespace OfficeIMO.Word.Html;

/// <summary>
/// Native Word document plus structured diagnostics from one HTML import operation.
/// </summary>
internal sealed class HtmlToWordResult : HtmlConversionResult<WordDocument> {
    internal HtmlToWordResult(WordDocument document, IEnumerable<HtmlDiagnostic> diagnostics) : base(document) {
        AddDiagnostics(diagnostics);
    }
}
