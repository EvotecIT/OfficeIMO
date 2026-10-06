namespace OfficeIMO.Pdf;

/// <summary>Immutable width and following gutter for a sequential flow column.</summary>
public sealed class PdfFlowColumn {
    /// <summary>Creates a fixed, percentage or weighted column with an optional gutter in points.</summary>
    /// <param name="width">Column sizing within the content width after gutters are reserved. Content-sized automatic widths are not supported.</param>
    /// <param name="gapAfter">Gutter after this column in points; null uses the containing layout's gap. The final gutter is unused.</param>
    public PdfFlowColumn(PdfColumnWidth width, double? gapAfter = null) {
        width.Validate(nameof(width));
        if (width.Unit == PdfColumnWidthUnit.Auto)
            throw new ArgumentException("Sequential columns require fixed, percentage or relative widths.", nameof(width));
        if (gapAfter is double gap && (gap < 0D || double.IsNaN(gap) || double.IsInfinity(gap)))
            throw new ArgumentOutOfRangeException(nameof(gapAfter), "Column gutters must be finite and non-negative.");
        Width = width;
        GapAfter = gapAfter;
    }

    /// <summary>Column sizing within the available width after gutters.</summary>
    public PdfColumnWidth Width { get; }
    /// <summary>Following gutter in points, or null for the layout's default.</summary>
    public double? GapAfter { get; }
}
