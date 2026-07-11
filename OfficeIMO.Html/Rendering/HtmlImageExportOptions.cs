namespace OfficeIMO.Html;

/// <summary>
/// Image-specific HTML rendering options following the shared OfficeIMO image export pattern.
/// </summary>
public sealed class HtmlImageExportOptions : HtmlRenderOptions {
    /// <summary>Zero-based page selected by single-image export methods in paged mode.</summary>
    public int PageIndex { get; set; }

    /// <summary>Creates an independent image export options snapshot.</summary>
    public override HtmlRenderOptions Clone() => CloneImage();

    /// <summary>Creates an independent image export options snapshot.</summary>
    public HtmlImageExportOptions CloneImage() {
        var clone = CopyTo(new HtmlImageExportOptions());
        clone.PageIndex = PageIndex;
        return clone;
    }
}
