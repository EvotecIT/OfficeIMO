namespace OfficeIMO.Pdf;

internal sealed class TableBlock : IPdfBlock {
    public System.Collections.Generic.List<string[]> Rows { get; } = new();
    public PdfAlign Align { get; }
    public PdfTableStyle? Style { get; }
    // Optional per-cell link URIs: key = (rowIndex, colIndex)
    public System.Collections.Generic.Dictionary<(int Row, int Col), string> Links { get; } = new();
    public TableBlock(System.Collections.Generic.IEnumerable<string[]> rows, PdfAlign align, PdfTableStyle? style) {
        Guard.NotNull(rows, nameof(rows));
        Align = align; Style = style;
        // Validate and normalize rows
        foreach (var r in rows) {
            if (r is null) throw new System.ArgumentException("Table rows cannot contain null entries.", nameof(rows));
            Rows.Add(r);
        }
        if (Rows.Count > 0) {
            // Ensure header row doesn't contain null cells to avoid rendering artifacts
            var header = Rows[0];
            for (int i = 0; i < header.Length; i++) if (header[i] is null) header[i] = string.Empty;
        }
    }
}

/// <summary>
/// Describes visual and layout options for table rendering.
/// Attach an instance to a table block or use the presets in <see cref="TableStyles"/>.
/// </summary>
public sealed class PdfTableStyle {
    /// <summary>
    /// Color of the table borders and cell grid lines. Set to <c>null</c> to hide borders.
    /// </summary>
    public PdfColor? BorderColor { get; set; } = new PdfColor(0.8, 0.8, 0.8);

    /// <summary>
    /// Stroke width, in points, for table borders and cell grid lines.
    /// </summary>
    public double BorderWidth { get; set; } = 0.5;

    /// <summary>
    /// Background fill color for the header row. Set to <c>null</c> for no fill.
    /// </summary>
    public PdfColor? HeaderFill { get; set; } = new PdfColor(0.95, 0.95, 0.95);

    /// <summary>
    /// Optional alternating row fill color (applied to every other body row). Set to <c>null</c> to disable.
    /// </summary>
    public PdfColor? RowStripeFill { get; set; } = new PdfColor(0.98, 0.98, 0.98);

    /// <summary>
    /// Text color for body rows. When <c>null</c> the writer’s default text color is used.
    /// </summary>
    public PdfColor? TextColor { get; set; }

    /// <summary>
    /// Text color for header cells. When <c>null</c> the writer’s default text color is used.
    /// </summary>
    public PdfColor? HeaderTextColor { get; set; }

    /// <summary>
    /// Horizontal padding inside each cell, in points.
    /// </summary>
    public double CellPaddingX { get; set; } = 4;

    /// <summary>
    /// Vertical padding inside each cell, in points.
    /// </summary>
    public double CellPaddingY { get; set; } = 2;

    /// <summary>
    /// Vertical baseline adjustment, in points. Positive moves text down, negative up.
    /// Use this to fine-tune how text sits within row rectangles for a given font/viewer.
    /// </summary>
    public double RowBaselineOffset { get; set; }

    /// <summary>
    /// Optional per-column alignment. When <c>null</c> or missing entries, columns fall back to <see cref="PdfColumnAlign.Left"/>.
    /// </summary>
    public System.Collections.Generic.List<PdfColumnAlign>? Alignments { get; set; }

    /// <summary>
    /// When true, cell values that look numeric are right-aligned automatically.
    /// </summary>
    public bool RightAlignNumeric { get; set; }
}
