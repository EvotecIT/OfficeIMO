namespace OfficeIMO.Pdf;

/// <summary>Two-up duplex booklet layout in print order, with blank pages padded to a multiple of four.</summary>
public sealed class PdfBookletOptions {
    /// <summary>Creates a booklet layout using the physical sheet size in PDF points.</summary>
    public PdfBookletOptions(PageSize sheetSize) => SheetSize = sheetSize;

    /// <summary>Physical sheet dimensions in PDF points.</summary>
    public PageSize SheetSize { get; }
    /// <summary>Outer sheet margin in PDF points.</summary>
    public double Margin { get; set; } = 18D;
    /// <summary>Space between the two page cells in PDF points.</summary>
    public double Gutter { get; set; } = 9D;
    /// <summary>Swaps left and right cells on every sheet side for right-to-left reading.</summary>
    public bool RightToLeft { get; set; }
    /// <summary>Maximum physical duplex sheets accepted by one operation.</summary>
    public int MaxPhysicalSheets { get; set; } = 100;
    /// <summary>Maximum source pages imported by one operation.</summary>
    public int MaxSourcePages { get; set; } = 500;
    /// <summary>Maximum bytes retained for imposed page content and the finished PDF.</summary>
    public long MaxOutputBytes { get; set; } = 256L * 1024L * 1024L;
    /// <summary>Allows omission of source annotations, forms, structure tags, attachments, output intents, navigation, catalog and page features, metadata, and encryption. See <see cref="PdfImpositionSourceFeatureLoss"/> for the reported losses. Signatures require a separate policy.</summary>
    public bool AllowSourceFeatureLoss { get; set; }
    /// <summary>Explicit handling for a signed source. The default rejects it.</summary>
    public PdfImpositionSignaturePolicy SignaturePolicy { get; set; } = PdfImpositionSignaturePolicy.Reject;

    internal PdfNUpOptions ToNUpOptions() {
        if (MaxPhysicalSheets < 1 || MaxPhysicalSheets > int.MaxValue / 2) throw new ArgumentOutOfRangeException(nameof(MaxPhysicalSheets));
        return new PdfNUpOptions(SheetSize, 2, 1) {
            Margin = Margin,
            HorizontalGutter = Gutter,
            VerticalGutter = 0D,
            MaxSheets = MaxPhysicalSheets * 2,
            MaxSourcePages = MaxSourcePages,
            MaxOutputBytes = MaxOutputBytes,
            AllowSourceFeatureLoss = AllowSourceFeatureLoss,
            SignaturePolicy = SignaturePolicy
        };
    }
}
