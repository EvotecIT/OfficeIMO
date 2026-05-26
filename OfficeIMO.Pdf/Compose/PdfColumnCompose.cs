namespace OfficeIMO.Pdf;

/// <summary>Column container used within <see cref="PdfContentCompose"/>.</summary>
public class PdfColumnCompose {
    private readonly PdfDoc _doc;
    internal PdfColumnCompose(PdfDoc doc) { _doc = doc; }
    /// <summary>Begins a new item builder in this column.</summary>
    public PdfItemCompose Item() => new PdfItemCompose(_doc);
    /// <summary>Adds one or more flow items to this column.</summary>
    public PdfColumnCompose Item(System.Action<PdfItemCompose> build) {
        Guard.NotNull(build, nameof(build));
        var item = new PdfItemCompose(_doc);
        build(item);
        return this;
    }
    /// <summary>Adds invisible vertical space to this column flow.</summary>
    public PdfColumnCompose Spacer(double height) {
        _doc.Spacer(height);
        return this;
    }
    /// <summary>Starts a new page from this column flow.</summary>
    public PdfColumnCompose PageBreak() {
        _doc.PageBreak();
        return this;
    }
}
