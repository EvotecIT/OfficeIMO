using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed class PdfPageSoftMaskResource {
    public PdfPageSoftMaskResource(PdfStream group, OfficeSoftMaskMode mode, OfficeColor backdropColor) {
        Group = group;
        Mode = mode;
        BackdropColor = backdropColor;
    }

    public PdfStream Group { get; }

    public OfficeSoftMaskMode Mode { get; }

    public OfficeColor BackdropColor { get; }
}
