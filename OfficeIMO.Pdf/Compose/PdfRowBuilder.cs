namespace OfficeIMO.Pdf;

using System;
/// <summary>Builds a horizontal row with explicit column sizing.</summary>
public sealed class PdfRowBuilder {
    private readonly PdfDocument _doc;
    private readonly RowBlock _row = new RowBlock();

    internal PdfRowBuilder(PdfDocument doc) { _doc = doc; }

    /// <summary>Sets the horizontal gutter, in points, between columns in this row.</summary>
    public PdfRowBuilder Gap(double points) {
        _row.SetGap(points);
        return this;
    }

    /// <summary>Applies row/column layout rhythm for this row.</summary>
    public PdfRowBuilder Style(PdfRowStyle style) {
        _row.SetStyle(style);
        return this;
    }

    /// <summary>Draws a vertical separator line between columns in this row.</summary>
    public PdfRowBuilder ColumnSeparator(PdfColor color, double width = 0.5D) {
        var style = _row.Style ?? new PdfRowStyle();
        style.ColumnSeparatorColor = color;
        style.ColumnSeparatorWidth = width;
        _row.SetStyle(style);
        return this;
    }

    /// <summary>Adds a column with an explicit sizing strategy. Static elements, semantic groups, and components retain their content.</summary>
    /// <remarks>Page boundaries, nested rows, automatic multi-column layouts, and page-dependent deferred content belong outside a row.</remarks>
    public PdfRowBuilder Column(PdfColumnWidth width, System.Action<PdfContentBuilder> build) {
        Guard.NotNull(build, nameof(build));
        width.Validate(nameof(width));
        System.Collections.Generic.IReadOnlyList<IPdfBlock> blocks = _doc.BuildFlowBlocks(build);
        ValidateColumnBlocks(blocks);
        var col = new RowColumn(width);
        foreach (IPdfBlock block in blocks) {
            col.AddBlock(block);
        }
        _row.AddColumn(col);
        return this;
    }

    /// <summary>Adds a column that receives a weighted share of remaining width.</summary>
    public PdfRowBuilder RelativeColumn(System.Action<PdfContentBuilder> build, double weight = 1D) =>
        Column(PdfColumnWidth.Relative(weight), build);

    /// <summary>Adds a column with an exact width in points.</summary>
    public PdfRowBuilder FixedColumn(double points, System.Action<PdfContentBuilder> build) =>
        Column(PdfColumnWidth.Fixed(points), build);

    /// <summary>Adds a content-sized column, optionally constrained in points.</summary>
    public PdfRowBuilder AutoColumn(System.Action<PdfContentBuilder> build, double minimum = 0D, double? maximum = null) =>
        Column(PdfColumnWidth.Auto(minimum, maximum), build);

    /// <summary>Adds a column that consumes a percentage of available row width.</summary>
    public PdfRowBuilder PercentColumn(double percent, System.Action<PdfContentBuilder> build) =>
        Column(PdfColumnWidth.Percent(percent), build);

    internal void Commit() {
        if (_row.Columns.Count == 0)
            throw new InvalidOperationException("Rows require at least one column.");

        double percentageTotal = 0D;
        foreach (RowColumn column in _row.Columns) {
            if (column.Width.Unit == PdfColumnWidthUnit.Percent) {
                percentageTotal += column.Width.Value;
            }
        }

        if (percentageTotal > 100.0001D) {
            throw new InvalidOperationException("Percentage row columns cannot exceed 100% of the available width.");
        }

        _doc.AddRow(_row);
    }

    private static void ValidateColumnBlocks(System.Collections.Generic.IReadOnlyList<IPdfBlock> blocks) {
        foreach (IPdfBlock block in blocks) {
            if (PdfFlowNestingRules.IsColumnFlowSupported(block)) {
                continue;
            }

            throw new NotSupportedException(
                "Row columns support static content, elements, semantic groups, and components. " +
                "Place page boundaries, nested rows, automatic multi-column layouts, and page-dependent content outside the row.");
        }
    }
}
