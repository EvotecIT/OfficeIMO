namespace OfficeIMO.Pdf;

/// <summary>
/// Entry point for the fluent composition DSL. Add ordinary flow through <see cref="Content"/>,
/// or use <see cref="Page"/> and <see cref="Section"/> for scoped page settings and layout.
/// </summary>
public sealed class PdfCompose {
    private readonly PdfDocument _doc;
    internal PdfCompose(PdfDocument doc) { _doc = doc; }

    /// <summary>Adds content to the document's current flow without introducing a page or section boundary.</summary>
    public PdfCompose Content(System.Action<PdfItemCompose> build) {
        Guard.NotNull(build, nameof(build));
        build(new PdfItemCompose(_doc));
        return this;
    }

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

