using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

/// <summary>An immutable inline scalar table in a newly created Pages body document.</summary>
public sealed class IWorkPagesTable : IWorkPagesBodyBlock {
    internal IWorkPagesTable(IWorkNativeTableData data) => Data = data;
    internal IWorkNativeTableData Data { get; }

    /// <summary>Gets the native table name. Its visible title is disabled.</summary>
    public string Name => Data.Name;
    /// <summary>Gets the number of rows, including unset positions.</summary>
    public int RowCount => Data.RowCount;
    /// <summary>Gets the number of columns, including unset positions.</summary>
    public int ColumnCount => Data.ColumnCount;
    /// <summary>Gets set cells in row-major order; unset positions remain empty.</summary>
    public IReadOnlyList<IWorkTableCell> Cells => Data.Cells;
    /// <summary>Gets the zero-offset frame in points. Its position in the body follows insertion order.</summary>
    public IWorkGeometry Geometry => Data.Geometry;
    /// <summary>Gets the font shared by all cells.</summary>
    public string FontName => Data.FontName;
    /// <summary>Gets the font size shared by all cells.</summary>
    public float FontSizePoints => Data.FontSizePoints;
    /// <summary>Gets the opaque sRGB text color shared by all cells.</summary>
    public IWorkColor Color => Data.Color;
}
