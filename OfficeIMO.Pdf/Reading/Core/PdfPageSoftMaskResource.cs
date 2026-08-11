using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed class PdfPageSoftMaskResource {
    public PdfPageSoftMaskResource(PdfStream group, OfficeSoftMaskMode mode, OfficeColor backdropColor, PdfDictionary? parentResources = null) {
        Group = group;
        Mode = mode;
        BackdropColor = backdropColor;
        ParentResources = parentResources;
    }

    public PdfStream Group { get; }

    public OfficeSoftMaskMode Mode { get; }

    public OfficeColor BackdropColor { get; }

    internal PdfDictionary? ParentResources { get; }
}
