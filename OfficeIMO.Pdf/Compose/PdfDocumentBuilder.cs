namespace OfficeIMO.Pdf;

using OfficeIMO.Drawing;

/// <summary>
/// Entry point for the fluent composition DSL. Add ordinary flow through <see cref="Content"/>,
/// or use <see cref="Page"/> and <see cref="Section"/> for scoped page settings and layout.
/// </summary>
public sealed class PdfDocumentBuilder {
    private readonly PdfDocument _doc;
    internal PdfDocumentBuilder(PdfDocument doc) { _doc = doc; }

    /// <summary>Adds content to the document's current flow without introducing a page or section boundary.</summary>
    public PdfDocumentBuilder Content(System.Action<PdfContentBuilder> build) {
        Guard.NotNull(build, nameof(build));
        build(new PdfContentBuilder(_doc));
        return this;
    }

    /// <summary>
    /// Updates document-wide rendering, catalog, security, compliance, and attachment settings.
    /// This is primarily useful to adapters that compose a document incrementally after creation.
    /// </summary>
    public PdfDocumentBuilder Settings(System.Action<PdfOptions> configure) {
        _doc.ConfigureSettings(configure);
        return this;
    }

    /// <summary>Applies a shared, explicit font, fallback, language, and shaping profile.</summary>
    public PdfDocumentBuilder Typography(
        OfficeRenderingProfile profile,
        OfficeRenderingProfileApplyMode mode = OfficeRenderingProfileApplyMode.Replace) {
        _doc.ConfigureTypography(profile, mode);
        return this;
    }

    /// <summary>Configures a page (size, margins, content, footer).</summary>
    public PdfDocumentBuilder Page(System.Action<PdfPageBuilder> configure) {
        _doc.AddComposedPage(configure);
        return this;
    }

    /// <summary>Configures a section-scoped flow with its own page setup and content.</summary>
    public PdfDocumentBuilder Section(System.Action<PdfPageBuilder> configure) {
        _doc.AddComposedPage(configure);
        return this;
    }
}

