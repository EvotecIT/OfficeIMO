namespace OfficeIMO.Pdf;

/// <summary>
/// Defines print-production page boxes as insets from the generated page MediaBox.
/// </summary>
/// <remarks>
/// Bleed insets must not exceed trim insets, so the resolved BleedBox contains the TrimBox.
/// Use <see cref="FullBleed"/> when the MediaBox, BleedBox, and TrimBox are identical.
/// </remarks>
public sealed class PdfPrintProductionPageBoxes {
    /// <summary>Creates print-production page boxes from MediaBox-relative insets.</summary>
    public PdfPrintProductionPageBoxes(PageMargins trimInsets, PageMargins bleedInsets) {
        ValidateContainment(trimInsets, bleedInsets);
        TrimInsets = trimInsets;
        BleedInsets = bleedInsets;
    }

    /// <summary>Insets used to resolve the TrimBox from the MediaBox.</summary>
    public PageMargins TrimInsets { get; }

    /// <summary>Insets used to resolve the BleedBox from the MediaBox.</summary>
    public PageMargins BleedInsets { get; }

    /// <summary>Creates coincident MediaBox, BleedBox, and TrimBox boundaries.</summary>
    public static PdfPrintProductionPageBoxes FullBleed { get; } =
        new(PageMargins.Uniform(0D), PageMargins.Uniform(0D));

    internal PdfResolvedPrintProductionPageBoxes Resolve(double pageWidth, double pageHeight) {
        Guard.Positive(pageWidth, nameof(pageWidth));
        Guard.Positive(pageHeight, nameof(pageHeight));
        if (TrimInsets.Left + TrimInsets.Right >= pageWidth ||
            TrimInsets.Top + TrimInsets.Bottom >= pageHeight) {
            throw new InvalidOperationException("PDF print-production trim insets must leave a positive TrimBox within the generated MediaBox.");
        }

        return new PdfResolvedPrintProductionPageBoxes(
            TrimInsets.Left,
            TrimInsets.Bottom,
            pageWidth - TrimInsets.Right,
            pageHeight - TrimInsets.Top,
            BleedInsets.Left,
            BleedInsets.Bottom,
            pageWidth - BleedInsets.Right,
            pageHeight - BleedInsets.Top);
    }

    private static void ValidateContainment(PageMargins trimInsets, PageMargins bleedInsets) {
        if (bleedInsets.Left > trimInsets.Left ||
            bleedInsets.Top > trimInsets.Top ||
            bleedInsets.Right > trimInsets.Right ||
            bleedInsets.Bottom > trimInsets.Bottom) {
            throw new ArgumentException("PDF print-production bleed insets cannot exceed trim insets because the BleedBox must contain the TrimBox.", nameof(bleedInsets));
        }
    }
}

internal readonly struct PdfResolvedPrintProductionPageBoxes {
    internal PdfResolvedPrintProductionPageBoxes(
        double trimLeft,
        double trimBottom,
        double trimRight,
        double trimTop,
        double bleedLeft,
        double bleedBottom,
        double bleedRight,
        double bleedTop) {
        TrimLeft = trimLeft;
        TrimBottom = trimBottom;
        TrimRight = trimRight;
        TrimTop = trimTop;
        BleedLeft = bleedLeft;
        BleedBottom = bleedBottom;
        BleedRight = bleedRight;
        BleedTop = bleedTop;
    }

    internal double TrimLeft { get; }
    internal double TrimBottom { get; }
    internal double TrimRight { get; }
    internal double TrimTop { get; }
    internal double BleedLeft { get; }
    internal double BleedBottom { get; }
    internal double BleedRight { get; }
    internal double BleedTop { get; }
}
