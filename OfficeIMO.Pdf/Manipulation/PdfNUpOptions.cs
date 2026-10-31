namespace OfficeIMO.Pdf;

/// <summary>Sheet geometry and preservation policy for vector PDF N-up output.</summary>
public sealed class PdfNUpOptions {
    /// <summary>Creates an N-up sheet layout.</summary>
    public PdfNUpOptions(PdfPageSize sheetSize, int columns, int rows) {
        SheetSize = sheetSize;
        Columns = columns;
        Rows = rows;
    }

    /// <summary>Output sheet dimensions in PDF points.</summary>
    public PdfPageSize SheetSize { get; }
    /// <summary>Number of columns, from one through eight.</summary>
    public int Columns { get; }
    /// <summary>Number of rows, from one through eight.</summary>
    public int Rows { get; }
    /// <summary>Outer sheet margin in PDF points.</summary>
    public OfficeLength Margin { get; set; } = OfficeIMO.OfficeLength.Points(18D);
    /// <summary>Horizontal gap between cells in PDF points.</summary>
    public OfficeLength HorizontalGutter { get; set; } = OfficeIMO.OfficeLength.Points(9D);
    /// <summary>Vertical gap between cells in PDF points.</summary>
    public OfficeLength VerticalGutter { get; set; } = OfficeIMO.OfficeLength.Points(9D);
    /// <summary>Maximum output sheets accepted by one operation.</summary>
    public int MaxSheets { get; set; } = 100;
    /// <summary>Maximum source pages imported by one operation.</summary>
    public int MaxSourcePages { get; set; } = 500;
    /// <summary>Maximum bytes retained for imposed page content and the finished PDF.</summary>
    public long MaxOutputBytes { get; set; } = 256L * 1024L * 1024L;
    /// <summary>Allows omission of source annotations, forms, structure tags, attachments, output intents, navigation, catalog and page features, metadata, and encryption. See <see cref="PdfImpositionSourceFeatureLoss"/> for the reported losses. Signatures require a separate policy.</summary>
    public bool AllowSourceFeatureLoss { get; set; }
    /// <summary>Explicit handling for a signed source. The default rejects it.</summary>
    public PdfImpositionSignaturePolicy SignaturePolicy { get; set; } = PdfImpositionSignaturePolicy.Reject;

    internal (double CellWidth, double CellHeight) Validate() {
        if (Columns < 1 || Columns > 8) throw new ArgumentOutOfRangeException(nameof(Columns));
        if (Rows < 1 || Rows > 8) throw new ArgumentOutOfRangeException(nameof(Rows));
        if (MaxSheets < 1) throw new ArgumentOutOfRangeException(nameof(MaxSheets));
        if (MaxSourcePages < 1) throw new ArgumentOutOfRangeException(nameof(MaxSourcePages));
        if (MaxOutputBytes < 1L || MaxOutputBytes > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(MaxOutputBytes));
        if (SignaturePolicy != PdfImpositionSignaturePolicy.Reject && SignaturePolicy != PdfImpositionSignaturePolicy.CreateUnsignedDerivative)
            throw new ArgumentOutOfRangeException(nameof(SignaturePolicy));
        if (SheetSize.Width.ToPoints() <= 0D || SheetSize.Height.ToPoints() <= 0D ||
            double.IsNaN(SheetSize.Width.ToPoints()) || double.IsNaN(SheetSize.Height.ToPoints()) ||
            double.IsInfinity(SheetSize.Width.ToPoints()) || double.IsInfinity(SheetSize.Height.ToPoints())) throw new ArgumentOutOfRangeException(nameof(SheetSize));
        ValidateNonNegative(Margin.ToPoints(), nameof(Margin));
        ValidateNonNegative(HorizontalGutter.ToPoints(), nameof(HorizontalGutter));
        ValidateNonNegative(VerticalGutter.ToPoints(), nameof(VerticalGutter));
        double width = (SheetSize.Width.ToPoints() - 2D * Margin.ToPoints() - (Columns - 1) * HorizontalGutter.ToPoints()) / Columns;
        double height = (SheetSize.Height.ToPoints() - 2D * Margin.ToPoints() - (Rows - 1) * VerticalGutter.ToPoints()) / Rows;
        if (width <= 0D || height <= 0D) throw new ArgumentException("Sheet margins and gutters leave no drawable cell area.");
        return (width, height);
    }

    private static void ValidateNonNegative(double value, string name) {
        if (value < 0D || double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(name);
    }
}
