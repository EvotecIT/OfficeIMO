using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed class PdfPageSoftMaskResource {
    public PdfPageSoftMaskResource(PdfStream group, OfficeSoftMaskMode mode, OfficeColor backdropColor, bool isIsolated, bool hasExplicitGroupColorSpace, PdfDictionary? parentResources = null) {
        Group = group;
        Mode = mode;
        BackdropColor = backdropColor;
        IsIsolated = isIsolated;
        HasExplicitGroupColorSpace = hasExplicitGroupColorSpace;
        ParentResources = parentResources;
    }

    public PdfStream Group { get; }

    public OfficeSoftMaskMode Mode { get; }

    public OfficeColor BackdropColor { get; }

    public bool IsIsolated { get; }

    public bool HasExplicitGroupColorSpace { get; }

    internal PdfDictionary? ParentResources { get; }
}
