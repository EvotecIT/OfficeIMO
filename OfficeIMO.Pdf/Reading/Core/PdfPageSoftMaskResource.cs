using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed class PdfPageSoftMaskResource {
    public PdfPageSoftMaskResource(PdfStream group, OfficeSoftMaskMode mode, OfficeColor backdropColor, bool isIsolated, PdfDictionary? parentResources = null) {
        Group = group;
        Mode = mode;
        BackdropColor = backdropColor;
        IsIsolated = isIsolated;
        ParentResources = parentResources;
    }

    public PdfStream Group { get; }

    public OfficeSoftMaskMode Mode { get; }

    public OfficeColor BackdropColor { get; }

    public bool IsIsolated { get; }

    internal PdfDictionary? ParentResources { get; }
}
