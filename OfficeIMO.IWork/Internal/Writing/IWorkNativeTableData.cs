namespace OfficeIMO.IWork.Internal;

/// <summary>Shared immutable scalar table snapshot consumed by the native table graph owner.</summary>
internal sealed class IWorkNativeTableData {
    internal IWorkNativeTableData(IWorkTableData data, float left, float top, float width, float height,
        string name, string fontName, float fontSize, IWorkColor color) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        IWorkCreationValidation.ValidateFinite(left, nameof(left));
        IWorkCreationValidation.ValidateFinite(top, nameof(top));
        IWorkCreationValidation.ValidatePositive(width, nameof(width));
        IWorkCreationValidation.ValidatePositive(height, nameof(height));
        IWorkCreationValidation.ValidateFontName(fontName);
        IWorkCreationValidation.ValidatePositive(fontSize, nameof(fontSize));
        IWorkCreationValidation.ValidateName(name, nameof(name));
        Name = IWorkCreationValidation.NormalizeText(name);
        RowCount = data.RowCount; ColumnCount = data.ColumnCount;
        Cells = Array.AsReadOnly(data.Snapshot());
        Geometry = new IWorkGeometry(left, top, width, height, 0);
        FontName = fontName; FontSizePoints = fontSize; Color = color;
    }

    internal string Name { get; }
    internal int RowCount { get; }
    internal int ColumnCount { get; }
    internal IReadOnlyList<IWorkTableCell> Cells { get; }
    internal IWorkGeometry Geometry { get; }
    internal string FontName { get; }
    internal float FontSizePoints { get; }
    internal IWorkColor Color { get; }

    internal IWorkProtoWriter ModelFields(int maximumBytes, System.Threading.CancellationToken cancellationToken) {
        IWorkProtoWriter P() => new(maximumBytes, cancellationToken);
        var data = P().UInt(1, (uint)RowCount).UInt(2, (uint)ColumnCount).String(3, Name)
            .Float(4, (float)Geometry.LeftPoints).Float(5, (float)Geometry.TopPoints)
            .Float(6, (float)Geometry.WidthPoints).Float(7, (float)Geometry.HeightPoints)
            .String(8, FontName).Float(9, FontSizePoints).String(10, Color.RgbHex);
        foreach (IWorkTableCell cell in Cells) {
            var value = P().UInt(1, (uint)cell.Row).UInt(2, (uint)cell.Column).UInt(3, (uint)cell.Kind);
            if (cell.Kind == IWorkCellKind.Text) value.String(4, (string)cell.Value!);
            else if (cell.Kind == IWorkCellKind.Number) value.Double(5, (double)cell.Value!);
            else if (cell.Kind == IWorkCellKind.Boolean) value.Bool(6, (bool)cell.Value!);
            data.Message(11, value);
        }
        return data;
    }
}
