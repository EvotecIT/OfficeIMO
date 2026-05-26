namespace OfficeIMO.Pdf;

internal sealed class RowBlock : IPdfBlock {
    private readonly System.Collections.Generic.List<RowColumn> _columns = new();
    private readonly System.Collections.ObjectModel.ReadOnlyCollection<RowColumn> _columnsView;
    private double? _gap;
    private PdfRowStyle? _style;

    public System.Collections.Generic.IReadOnlyList<RowColumn> Columns => _columnsView;
    public double Gap => _gap ?? _style?.Gap ?? PdfRowStyle.DefaultGap;
    public PdfRowStyle? Style => _style?.Clone();
    internal double? GapOverride => _gap;
    internal PdfRowStyle? StyleSnapshot => _style;

    public RowBlock() {
        _columnsView = new System.Collections.ObjectModel.ReadOnlyCollection<RowColumn>(_columns);
    }

    internal void AddColumn(RowColumn column) {
        Guard.NotNull(column, nameof(column));
        _columns.Add(column);
    }

    internal void SetGap(double gap) {
        Guard.NonNegative(gap, nameof(gap));
        _gap = gap;
    }

    internal void SetStyle(PdfRowStyle style) {
        Guard.NotNull(style, nameof(style));
        _style = style.Clone();
    }
}
