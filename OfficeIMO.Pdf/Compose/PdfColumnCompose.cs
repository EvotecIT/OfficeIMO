namespace OfficeIMO.Pdf;

/// <summary>Column container used within <see cref="PdfContentCompose"/>.</summary>
public class PdfColumnCompose {
    private readonly PdfDoc _doc;
    internal PdfColumnCompose(PdfDoc doc) { _doc = doc; }
    /// <summary>Begins a new item builder in this column.</summary>
    public PdfItemCompose Item() => new PdfItemCompose(_doc);
}

