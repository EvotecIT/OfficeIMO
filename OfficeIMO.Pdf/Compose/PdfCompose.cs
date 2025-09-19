namespace OfficeIMO.Pdf;

/// <summary>
/// Entry point for the fluent composition DSL. Use <see cref="Page"/> to configure a page and its content.
/// </summary>
public class PdfCompose {
    private readonly PdfDoc _doc;
    internal PdfCompose(PdfDoc doc) { _doc = doc; }
    /// <summary>Configures a page (size, margins, content, footer).</summary>
    public PdfCompose Page(System.Action<PdfPageCompose> configure) { var p = new PdfPageCompose(_doc); configure(p); return this; }
}

