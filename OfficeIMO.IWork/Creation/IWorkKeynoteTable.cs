using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

/// <summary>An immutable positioned scalar table in a newly created Keynote presentation.</summary>
/// <remarks>Rows and columns are equally sized. All cells share a font, color, white fill and four-point padding.
/// Text wrapping is enabled; row resizing, formulas, merges and selected cell formatting are not applied.</remarks>
public sealed class IWorkKeynoteTable : IWorkKeynoteCreationDrawable {
    internal IWorkKeynoteTable(IWorkTableData data, float left, float top, float width, float height,
        string name, string fontName, float fontSize, IWorkColor color) {
        Data = new IWorkNativeTableData(data, left, top, width, height, name, fontName, fontSize, color);
    }

    internal IWorkNativeTableData Data { get; }

    /// <summary>Gets the table name retained in the native model. The visible title is disabled.</summary>
    public string Name => Data.Name;
    /// <summary>Gets the row count, including unset positions.</summary>
    public int RowCount => Data.RowCount;
    /// <summary>Gets the column count, including unset positions.</summary>
    public int ColumnCount => Data.ColumnCount;
    /// <summary>Gets set cells in row-major order. Unset positions remain empty.</summary>
    public IReadOnlyList<IWorkTableCell> Cells => Data.Cells;
    /// <summary>Gets the unrotated table frame in points.</summary>
    public IWorkGeometry Geometry => Data.Geometry;
    /// <summary>Gets the font shared by all cells.</summary>
    public string FontName => Data.FontName;
    /// <summary>Gets the font size shared by all cells.</summary>
    public float FontSizePoints => Data.FontSizePoints;
    /// <summary>Gets the opaque sRGB text color shared by all cells.</summary>
    public IWorkColor Color => Data.Color;
}
