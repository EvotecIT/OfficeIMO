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
    /// <summary>Allows a visual-only result when source annotations, forms, signatures, or tags cannot be retained.</summary>
    public bool AllowVisualOnlySourceFeatures { get; set; }

    internal PdfNUpOptions ToNUpOptions() {
        if (MaxPhysicalSheets < 1 || MaxPhysicalSheets > int.MaxValue / 2) throw new ArgumentOutOfRangeException(nameof(MaxPhysicalSheets));
        return new PdfNUpOptions(SheetSize, 2, 1) {
            Margin = Margin,
            HorizontalGutter = Gutter,
            VerticalGutter = 0D,
            MaxSheets = MaxPhysicalSheets * 2,
            MaxSourcePages = MaxSourcePages,
            AllowVisualOnlySourceFeatures = AllowVisualOnlySourceFeatures
        };
    }
}
