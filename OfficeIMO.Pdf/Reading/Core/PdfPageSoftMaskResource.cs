using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed class PdfPageSoftMaskResource {
    public PdfPageSoftMaskResource(PdfStream group, OfficeSoftMaskMode mode, OfficeColor backdropColor, bool hasByteExactBackdrop, bool isIsolated, bool hasExplicitGroupColorSpace, PdfDictionary? parentResources = null, byte[]? transferTable = null) {
        Group = group;
        Mode = mode;
        BackdropColor = backdropColor;
        HasByteExactBackdrop = hasByteExactBackdrop;
        IsIsolated = isIsolated;
        HasExplicitGroupColorSpace = hasExplicitGroupColorSpace;
        ParentResources = parentResources;
        TransferTable = transferTable;
    }

    /// <summary>The /TR transfer function sampled at 256 mask values, or null for the identity.</summary>
    public byte[]? TransferTable { get; }

    public PdfStream Group { get; }

    public OfficeSoftMaskMode Mode { get; }

    public OfficeColor BackdropColor { get; }

    public bool HasByteExactBackdrop { get; }

    public bool IsIsolated { get; }

    public bool HasExplicitGroupColorSpace { get; }

    internal PdfDictionary? ParentResources { get; }
}
