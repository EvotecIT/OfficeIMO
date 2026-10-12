using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

/// <summary>Collects newly authored scalar tables on one Numbers sheet.</summary>
public sealed class IWorkNumbersSheetBuilder {
    private readonly List<IWorkNumbersTable> _tables = new();
    internal IWorkNumbersSheetBuilder(string name) { Name = name; Tables = _tables.AsReadOnly(); }

    /// <summary>Gets the sheet name.</summary>
    public string Name { get; }
    /// <summary>Gets tables in insertion order.</summary>
    public IReadOnlyList<IWorkNumbersTable> Tables { get; }

    /// <summary>Adds a snapshot of scalar data with equally sized rows and columns.</summary>
    /// <remarks>Positions and dimensions are in points. Tables on one sheet require distinct names compared without case.
    /// Cells share the specified font/color, white fill, black solid grid, wrapping and four-point padding.</remarks>
    public IWorkNumbersTable AddTable(IWorkTableData data, string name = "Table 1", float leftPoints = 0,
        float topPoints = 0, float columnWidthPoints = 100, float rowHeightPoints = 24,
        string fontName = "Arial", float fontSizePoints = 12, string color = "000000") {
        if (data == null) throw new ArgumentNullException(nameof(data));
        IWorkCreationValidation.ValidatePositive(columnWidthPoints, nameof(columnWidthPoints));
        IWorkCreationValidation.ValidatePositive(rowHeightPoints, nameof(rowHeightPoints));
        if (_tables.Any(table => string.Equals(table.Name, name, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException("Table names on a sheet must be unique without regard to case.", nameof(name));
        var snapshot = new IWorkNativeTableData(data, leftPoints, topPoints,
            columnWidthPoints * data.ColumnCount, rowHeightPoints * data.RowCount,
            name, fontName, fontSizePoints, IWorkCreationValidation.ParseColor(color));
        var table = new IWorkNumbersTable(snapshot); _tables.Add(table); return table;
    }
}
