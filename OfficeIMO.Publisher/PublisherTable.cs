using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher;

/// <summary>An immutable snapshot of a current table grid and its placement on a document or master page.</summary>
public sealed class PublisherTable {
    internal PublisherTable(uint id, uint? storyId, uint pageId, double x, double y, double width, double height,
        OfficeTransform pageTransform, IReadOnlyList<double> columns, IReadOnlyList<double> rows,
        IReadOnlyList<PublisherTableCell> cells, bool hasTextMapping, bool hasPlacement = true,
        string? cellTextEditLimitation = null, string? boundsEditLimitation = null) {
        Id = id; StoryId = storyId; PageId = pageId; X = x; Y = y; Width = width; Height = height;
        PageTransform = pageTransform; HasTextMapping = hasTextMapping; HasPlacement = hasPlacement;
        ColumnWidths = Array.AsReadOnly(columns.ToArray()); RowHeights = Array.AsReadOnly(rows.ToArray());
        Cells = Array.AsReadOnly(cells.ToArray());
        CellTextEditLimitation = cellTextEditLimitation; BoundsEditLimitation = boundsEditLimitation;
    }
    /// <summary>Stable current publication object identifier.</summary>
    public uint Id { get; }
    /// <summary>Current complete story identifier, or null when no story is declared.</summary>
    public uint? StoryId { get; }
    /// <summary>Owning document or master page identifier.</summary>
    public uint PageId { get; }
    /// <summary>Left edge in page-local points before object and enclosing-group rotation or reflection.</summary>
    public double X { get; }
    /// <summary>Top edge in page-local points before object and enclosing-group rotation or reflection.</summary>
    public double Y { get; }
    /// <summary>Current frame width in points. Imported grid tracks can retain dimensions differing from the frame.</summary>
    public double Width { get; }
    /// <summary>Current frame height in points. Imported grid tracks can retain dimensions differing from the frame.</summary>
    public double Height { get; }
    /// <summary>Whether the recovered native frame has usable page placement. False retains grid and cell data without paint geometry.</summary>
    public bool HasPlacement { get; }
    /// <summary>Maps table-local points into page-local points, including object and enclosing-group rotation/reflection.</summary>
    public OfficeTransform PageTransform { get; }
    /// <summary>Current column widths in points, in grid order. Imported tracks use saved layout endpoints when available.</summary>
    public IReadOnlyList<double> ColumnWidths { get; }
    /// <summary>Current row heights in points, in grid order. Imported tracks use saved layout endpoints when available; no native reflow is performed.</summary>
    public IReadOnlyList<double> RowHeights { get; }
    /// <summary>Number of current grid columns.</summary>
    public int ColumnCount => ColumnWidths.Count;
    /// <summary>Number of current grid rows.</summary>
    public int RowCount => RowHeights.Count;
    /// <summary>Current cells in imported definition order or authored row-major order. A merged cell appears once with its span.</summary>
    public IReadOnlyList<PublisherTableCell> Cells { get; }
    /// <summary>Whether native text boundaries resolve to every recovered cell. False retains the grid with empty cell text and an omission report.</summary>
    public bool HasTextMapping { get; }
    /// <summary>Whether anchor-cell paragraph replacement has a qualified, complete and unambiguous story partition.</summary>
    public bool CanEditCellText => CellTextEditLimitation == null;
    /// <summary>Why cell paragraph replacement is unavailable, or null when supported.</summary>
    public string? CellTextEditLimitation { get; }
    /// <summary>Whether the current table can be moved and proportionally resized.</summary>
    public bool CanEditBounds => BoundsEditLimitation == null;
    /// <summary>Why table placement edits are unavailable, or null when supported.</summary>
    public string? BoundsEditLimitation { get; }
}
