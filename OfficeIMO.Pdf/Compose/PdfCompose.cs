namespace OfficeIMO.Pdf;

/// <summary>
/// Entry point for the fluent composition DSL. Use <see cref="Page"/> to configure a page and its content.
/// </summary>
public class PdfCompose {
    private readonly PdfDocument _doc;
    internal PdfCompose(PdfDocument doc) { _doc = doc; }
    /// <summary>Configures a page (size, margins, content, footer).</summary>
    public PdfCompose Page(System.Action<PdfPageCompose> configure) {
        _doc.AddComposedPage(configure);
        return this;
    }

    /// <summary>Configures a section-scoped flow with its own page setup and content.</summary>
    public PdfCompose Section(System.Action<PdfPageCompose> configure) {
        _doc.AddComposedPage(configure);
        return this;
    }
}

