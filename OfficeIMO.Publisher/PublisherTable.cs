using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher;

/// <summary>A recovered native table grid and its placement on a document or master page.</summary>
public sealed class PublisherTable {
    internal PublisherTable(uint id, uint? storyId, uint pageId, double x, double y, double width, double height,
        OfficeTransform pageTransform, IReadOnlyList<double> columns, IReadOnlyList<double> rows,
        IReadOnlyList<PublisherTableCell> cells, bool hasTextMapping) {
        Id = id; StoryId = storyId; PageId = pageId; X = x; Y = y; Width = width; Height = height;
        PageTransform = pageTransform; HasTextMapping = hasTextMapping;
        ColumnWidths = Array.AsReadOnly(columns.ToArray()); RowHeights = Array.AsReadOnly(rows.ToArray());
        Cells = Array.AsReadOnly(cells.ToArray());
    }
    /// <summary>Native publication object identifier.</summary>
    public uint Id { get; }
    /// <summary>Declared source story identifier, or null when no story is declared.</summary>
    public uint? StoryId { get; }
    /// <summary>Owning document or master page identifier.</summary>
    public uint PageId { get; }
    /// <summary>Left edge in page-local points before object and enclosing-group rotation or reflection.</summary>
    public double X { get; }
    /// <summary>Top edge in page-local points before object and enclosing-group rotation or reflection.</summary>
    public double Y { get; }
    /// <summary>Native frame width in points. The recovered grid tracks retain their own dimensions.</summary>
    public double Width { get; }
    /// <summary>Native frame height in points. The recovered grid tracks retain their own dimensions.</summary>
    public double Height { get; }
    /// <summary>Maps table-local points into page-local points, including object and enclosing-group rotation/reflection.</summary>
    public OfficeTransform PageTransform { get; }
    /// <summary>Native column widths in points, in grid order.</summary>
    public IReadOnlyList<double> ColumnWidths { get; }
    /// <summary>Native row heights in points, in grid order.</summary>
    public IReadOnlyList<double> RowHeights { get; }
    /// <summary>Number of native grid columns.</summary>
    public int ColumnCount => ColumnWidths.Count;
    /// <summary>Number of native grid rows.</summary>
    public int RowCount => RowHeights.Count;
    /// <summary>Recovered cells in native definition order. A merged cell appears once with its span.</summary>
    public IReadOnlyList<PublisherTableCell> Cells { get; }
    /// <summary>Whether native text boundaries resolve to every recovered cell. False retains the grid with empty cell text and an omission report.</summary>
    public bool HasTextMapping { get; }
}
