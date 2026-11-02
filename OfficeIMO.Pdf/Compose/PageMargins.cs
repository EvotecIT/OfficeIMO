namespace OfficeIMO.Pdf;

/// <summary>Represents physical page margins.</summary>
public readonly struct PdfPageMargins {
    /// <summary>Left margin.</summary>
    public OfficeLength Left { get; }
    /// <summary>Top margin.</summary>
    public OfficeLength Top { get; }
    /// <summary>Right margin.</summary>
    public OfficeLength Right { get; }
    /// <summary>Bottom margin.</summary>
    public OfficeLength Bottom { get; }

    /// <summary>Creates page margins with explicit units.</summary>
    public PdfPageMargins(OfficeLength left, OfficeLength top, OfficeLength right, OfficeLength bottom) {
        if (left.ToEmus() < 0) throw new ArgumentOutOfRangeException(nameof(left));
        if (top.ToEmus() < 0) throw new ArgumentOutOfRangeException(nameof(top));
        if (right.ToEmus() < 0) throw new ArgumentOutOfRangeException(nameof(right));
        if (bottom.ToEmus() < 0) throw new ArgumentOutOfRangeException(nameof(bottom));
        Left = left;
        Top = top;
        Right = right;
        Bottom = bottom;
    }

    /// <summary>Creates uniform margins.</summary>
    public static PdfPageMargins Uniform(OfficeLength length) => new PdfPageMargins(length, length, length, length);

    /// <summary>Normal Word-compatible one-inch margins.</summary>
    public static PdfPageMargins Normal => Uniform(OfficeLength.Inches(1));
    /// <summary>Narrow Word-compatible half-inch margins.</summary>
    public static PdfPageMargins Narrow => Uniform(OfficeLength.Inches(0.5));
    /// <summary>Moderate Word-compatible margins.</summary>
    public static PdfPageMargins Moderate => new PdfPageMargins(OfficeLength.Points(54), OfficeLength.Inches(1), OfficeLength.Points(54), OfficeLength.Inches(1));
    /// <summary>Wide Word-compatible margins.</summary>
    public static PdfPageMargins Wide => new PdfPageMargins(OfficeLength.Inches(2), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1));
    /// <summary>Mirrored Word-compatible margin preset values.</summary>
    public static PdfPageMargins Mirrored => new PdfPageMargins(OfficeLength.Points(90), OfficeLength.Inches(1), OfficeLength.Inches(1), OfficeLength.Inches(1));
    /// <summary>Word 2003 compatible default margin preset values.</summary>
    public static PdfPageMargins Office2003Default => new PdfPageMargins(OfficeLength.Points(90), OfficeLength.Inches(1), OfficeLength.Points(90), OfficeLength.Inches(1));
}
