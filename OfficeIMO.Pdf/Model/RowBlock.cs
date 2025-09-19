namespace OfficeIMO.Pdf;

internal sealed class RowBlock : IPdfBlock {
    public System.Collections.Generic.List<RowColumn> Columns { get; } = new();
}

internal sealed class RowColumn {
    public double WidthPercent { get; }
    public System.Collections.Generic.List<IPdfBlock> Blocks { get; } = new();
    public RowColumn(double widthPercent) { WidthPercent = widthPercent; }
}

