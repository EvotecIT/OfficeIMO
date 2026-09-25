namespace OfficeIMO.Pdf;

/// <summary>Sheet geometry and preservation policy for vector PDF N-up output.</summary>
public sealed class PdfNUpOptions {
    /// <summary>Creates an N-up sheet layout.</summary>
    public PdfNUpOptions(PageSize sheetSize, int columns, int rows) {
        SheetSize = sheetSize;
        Columns = columns;
        Rows = rows;
    }

    /// <summary>Output sheet dimensions in PDF points.</summary>
    public PageSize SheetSize { get; }
    /// <summary>Number of columns, from one through eight.</summary>
    public int Columns { get; }
    /// <summary>Number of rows, from one through eight.</summary>
    public int Rows { get; }
    /// <summary>Outer sheet margin in PDF points.</summary>
    public double Margin { get; set; } = 18D;
    /// <summary>Horizontal gap between cells in PDF points.</summary>
    public double HorizontalGutter { get; set; } = 9D;
    /// <summary>Vertical gap between cells in PDF points.</summary>
    public double VerticalGutter { get; set; } = 9D;
    /// <summary>Maximum output sheets accepted by one operation.</summary>
    public int MaxSheets { get; set; } = 100;
    /// <summary>Maximum source pages imported by one operation.</summary>
    public int MaxSourcePages { get; set; } = 500;
    /// <summary>Maximum bytes retained for imposed page content and the finished PDF.</summary>
    public long MaxOutputBytes { get; set; } = 256L * 1024L * 1024L;
    /// <summary>Allows source annotations, form widgets, and structure tags to be omitted; their appearances can also be lost.</summary>
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
        if (SheetSize.Width <= 0D || SheetSize.Height <= 0D ||
            double.IsNaN(SheetSize.Width) || double.IsNaN(SheetSize.Height) ||
            double.IsInfinity(SheetSize.Width) || double.IsInfinity(SheetSize.Height)) throw new ArgumentOutOfRangeException(nameof(SheetSize));
        ValidateNonNegative(Margin, nameof(Margin));
        ValidateNonNegative(HorizontalGutter, nameof(HorizontalGutter));
        ValidateNonNegative(VerticalGutter, nameof(VerticalGutter));
        double width = (SheetSize.Width - 2D * Margin - (Columns - 1) * HorizontalGutter) / Columns;
        double height = (SheetSize.Height - 2D * Margin - (Rows - 1) * VerticalGutter) / Rows;
        if (width <= 0D || height <= 0D) throw new ArgumentException("Sheet margins and gutters leave no drawable cell area.");
        return (width, height);
    }

    private static void ValidateNonNegative(double value, string name) {
        if (value < 0D || double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(name);
    }
}
