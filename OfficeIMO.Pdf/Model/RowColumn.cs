namespace OfficeIMO.Pdf;

internal sealed class RowColumn {
    public double WidthPercent { get; internal set; }
    public System.Collections.Generic.List<IPdfBlock> Blocks { get; } = new();
    public RowColumn(double widthPercent) { WidthPercent = widthPercent; }
}

