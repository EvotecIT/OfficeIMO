namespace OfficeIMO.Pdf;

internal sealed class RowBlock : IPdfBlock {
    public System.Collections.Generic.List<RowColumn> Columns { get; } = new();
}
