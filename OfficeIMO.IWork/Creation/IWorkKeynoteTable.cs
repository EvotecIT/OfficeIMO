namespace OfficeIMO.IWork;

/// <summary>An immutable positioned scalar table in a newly created Keynote presentation.</summary>
/// <remarks>Rows and columns are equally sized. All cells share a font, color, white fill and four-point padding.
/// Text wrapping is enabled; row resizing, formulas, merges and selected cell formatting are not applied.</remarks>
public sealed class IWorkKeynoteTable : IWorkKeynoteCreationDrawable {
    internal IWorkKeynoteTable(IWorkTableData data, float left, float top, float width, float height,
        string name, string fontName, float fontSize, IWorkColor color) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        IWorkCreationValidation.ValidateFinite(left, nameof(left));
        IWorkCreationValidation.ValidateFinite(top, nameof(top));
        IWorkCreationValidation.ValidatePositive(width, nameof(width));
        IWorkCreationValidation.ValidatePositive(height, nameof(height));
        IWorkCreationValidation.ValidateFontName(fontName);
        IWorkCreationValidation.ValidatePositive(fontSize, nameof(fontSize));
        if (name == null) throw new ArgumentNullException(nameof(name));
        if (string.IsNullOrWhiteSpace(name) || name.Length > 128 || name.Any(char.IsControl))
            throw new ArgumentException("A table name of at most 128 characters without controls is required.", nameof(name));
        Name = IWorkCreationValidation.NormalizeText(name);
        RowCount = data.RowCount; ColumnCount = data.ColumnCount;
        Cells = Array.AsReadOnly(data.Snapshot());
        Geometry = new IWorkGeometry(left, top, width, height, 0);
        FontName = fontName; FontSizePoints = fontSize; Color = color;
    }

    /// <summary>Gets the table name retained in the native model. The visible title is disabled.</summary>
    public string Name { get; }
    /// <summary>Gets the row count, including unset positions.</summary>
    public int RowCount { get; }
    /// <summary>Gets the column count, including unset positions.</summary>
    public int ColumnCount { get; }
    /// <summary>Gets set cells in row-major order. Unset positions remain empty.</summary>
    public IReadOnlyList<IWorkTableCell> Cells { get; }
    /// <summary>Gets the unrotated table frame in points.</summary>
    public IWorkGeometry Geometry { get; }
    /// <summary>Gets the font shared by all cells.</summary>
    public string FontName { get; }
    /// <summary>Gets the font size shared by all cells.</summary>
    public float FontSizePoints { get; }
    /// <summary>Gets the opaque sRGB text color shared by all cells.</summary>
    public IWorkColor Color { get; }
}
