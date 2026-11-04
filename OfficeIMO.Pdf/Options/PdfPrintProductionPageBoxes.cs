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
    public PdfPrintProductionPageBoxes(PdfPageMargins trimInsets, PdfPageMargins bleedInsets) {
        ValidateContainment(trimInsets, bleedInsets);
        TrimInsets = trimInsets;
        BleedInsets = bleedInsets;
    }

    /// <summary>Insets used to resolve the TrimBox from the MediaBox.</summary>
    public PdfPageMargins TrimInsets { get; }

    /// <summary>Insets used to resolve the BleedBox from the MediaBox.</summary>
    public PdfPageMargins BleedInsets { get; }

    /// <summary>Creates coincident MediaBox, BleedBox, and TrimBox boundaries.</summary>
    public static PdfPrintProductionPageBoxes FullBleed { get; } =
        new(PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(0D)), PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(0D)));

    internal PdfResolvedPrintProductionPageBoxes Resolve(double pageWidth, double pageHeight) {
        Guard.Positive(pageWidth, nameof(pageWidth));
        Guard.Positive(pageHeight, nameof(pageHeight));
        if (TrimInsets.Left.ToPoints() + TrimInsets.Right.ToPoints() >= pageWidth ||
            TrimInsets.Top.ToPoints() + TrimInsets.Bottom.ToPoints() >= pageHeight) {
            throw new InvalidOperationException("PDF print-production trim insets must leave a positive TrimBox within the generated MediaBox.");
        }

        return new PdfResolvedPrintProductionPageBoxes(
            TrimInsets.Left.ToPoints(),
            TrimInsets.Bottom.ToPoints(),
            pageWidth - TrimInsets.Right.ToPoints(),
            pageHeight - TrimInsets.Top.ToPoints(),
            BleedInsets.Left.ToPoints(),
            BleedInsets.Bottom.ToPoints(),
            pageWidth - BleedInsets.Right.ToPoints(),
            pageHeight - BleedInsets.Top.ToPoints());
    }

    private static void ValidateContainment(PdfPageMargins trimInsets, PdfPageMargins bleedInsets) {
        if (bleedInsets.Left.ToPoints() > trimInsets.Left.ToPoints() ||
            bleedInsets.Top.ToPoints() > trimInsets.Top.ToPoints() ||
            bleedInsets.Right.ToPoints() > trimInsets.Right.ToPoints() ||
            bleedInsets.Bottom.ToPoints() > trimInsets.Bottom.ToPoints()) {
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
