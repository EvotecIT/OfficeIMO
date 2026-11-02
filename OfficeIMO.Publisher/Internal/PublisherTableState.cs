using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Current table placement and grid. Cell content belongs to the complete story's ordered partition in PublisherTextState.</summary>
internal sealed class PublisherTableState {
    private readonly OfficeShape? _paint;
    internal PublisherTableState(uint id, uint? storyId, uint pageId, double x, double y, double width, double height,
        OfficeTransform transform, bool hasPlacement, bool hidden, IReadOnlyList<double> columns, IReadOnlyList<double> rows,
        IReadOnlyList<PublisherTableCellSlot> cells, OfficeTextPadding padding, OfficeShape? paint,
        string? boundsLimitation, bool isAuthored = false, IReadOnlyList<OfficeColor?>? cellFills = null,
        IReadOnlyList<PublisherTableBorderSegment>? borders = null) {
        Id = id; StoryId = storyId; PageId = pageId; X = x; Y = y; Width = width; Height = height;
        PageTransform = transform; HasPlacement = hasPlacement; Hidden = hidden;
        Columns = Array.AsReadOnly(columns.ToArray()); Rows = Array.AsReadOnly(rows.ToArray());
        Cells = Array.AsReadOnly(cells.ToArray()); Padding = padding; _paint = paint?.Clone();
        BoundsEditLimitation = boundsLimitation; IsAuthored = isAuthored;
        CellFills = Array.AsReadOnly((cellFills ?? new OfficeColor?[cells.Count]).ToArray());
        BorderSegments = Array.AsReadOnly((borders ?? Array.Empty<PublisherTableBorderSegment>()).ToArray());
        if (CellFills.Count != cells.Count) throw new ArgumentException("Cell fill count must match cell ownership.", nameof(cellFills));
    }
    internal uint Id { get; }
    internal uint? StoryId { get; }
    internal uint PageId { get; }
    internal double X { get; }
    internal double Y { get; }
    internal double Width { get; }
    internal double Height { get; }
    internal OfficeTransform PageTransform { get; }
    internal bool HasPlacement { get; }
    internal bool Hidden { get; }
    internal IReadOnlyList<double> Columns { get; }
    internal IReadOnlyList<double> Rows { get; }
    internal IReadOnlyList<PublisherTableCellSlot> Cells { get; }
    internal OfficeTextPadding Padding { get; }
    internal bool HasSemanticPaint => _paint != null;
    internal string? BoundsEditLimitation { get; }
    internal bool IsAuthored { get; }
    internal IReadOnlyList<OfficeColor?> CellFills { get; }
    internal IReadOnlyList<PublisherTableBorderSegment> BorderSegments { get; }
    internal OfficeShape? CreatePaint() => _paint?.Clone();

    internal PublisherTableState WithDecoration(IReadOnlyList<OfficeColor?> fills, IReadOnlyList<PublisherTableBorderSegment> borders) =>
        new(Id, StoryId, PageId, X, Y, Width, Height, PageTransform, HasPlacement, Hidden,
            Columns, Rows, Cells, Padding, _paint, BoundsEditLimitation, IsAuthored, fills, borders);

    internal PublisherTableState WithCells(IReadOnlyList<PublisherTableCellSlot> cells) {
        var fills = Cells.Select((cell, index) => (cell, index)).ToDictionary(item => (item.cell.Row, item.cell.Column), item => CellFills[item.index]);
        return new(Id, StoryId, PageId, X, Y, Width, Height, PageTransform, HasPlacement, Hidden,
            Columns, Rows, cells, Padding, _paint, BoundsEditLimitation, IsAuthored,
            cells.Select(cell => fills[(cell.Row, cell.Column)]).ToArray(), BorderSegments);
    }

    internal bool HasTextMapping(PublisherTextState text) => StoryId.HasValue && text.Stories.ContainsKey(StoryId.Value)
        && text.TableStories.TryGetValue(StoryId.Value, out PublisherTableStory? partition) && partition.Segments.Count == Cells.Count;

    internal PublisherTable Snapshot(PublisherTextState text, string? textLimitation) {
        bool mapped = HasTextMapping(text);
        IReadOnlyList<PublisherTextStory>? segments = mapped ? text.TableStories[StoryId!.Value].Segments : null;
        double[] xs = Offsets(Columns), ys = Offsets(Rows);
        var cells = new List<PublisherTableCell>(Cells.Count);
        for (int i = 0; i < Cells.Count; i++) {
            PublisherTableCellSlot cell = Cells[i];
            cells.Add(new PublisherTableCell(cell.Row, cell.Column, cell.RowSpan, cell.ColumnSpan, xs[cell.Column], ys[cell.Row],
                xs[cell.Column + cell.ColumnSpan] - xs[cell.Column], ys[cell.Row + cell.RowSpan] - ys[cell.Row],
                segments == null ? Array.Empty<OfficeRichTextParagraph>() : segments[i].Paragraphs, CellFills[i]));
        }
        return new PublisherTable(Id, StoryId, PageId, X, Y, Width, Height, PageTransform, Columns, Rows, cells,
            mapped, HasPlacement, textLimitation, BoundsEditLimitation, BorderSegments);
    }

    internal PublisherTableState WithBounds(double x, double y, double width, double height) {
        double[] columns = Columns.Select(value => value / Width * width).ToArray();
        double[] rows = Rows.Select(value => value / Height * height).ToArray();
        ValidateTracks(columns, nameof(width)); ValidateTracks(rows, nameof(height));
        OfficeShape? paint = CreatePaint();
        if (paint != null) { paint.Width = width; paint.Height = height; }
        return new PublisherTableState(Id, StoryId, PageId, x, y, width, height, OfficeTransform.Translate(x, y), true,
            Hidden, columns, rows, Cells, Padding, paint, BoundsEditLimitation, IsAuthored, CellFills, BorderSegments);
    }

    internal static double ValidateTracks(IReadOnlyList<double> tracks, string name) {
        if (tracks.Count == 0) throw new ArgumentException("A table requires at least one track on each axis.", name);
        double total = 0;
        foreach (double value in tracks) {
            if (!Finite(value) || value <= 0) throw new ArgumentOutOfRangeException(name, "Table tracks must be positive and finite.");
            total += value;
            if (!Finite(total)) throw new ArgumentOutOfRangeException(name, "Table track sums must be finite.");
        }
        return total;
    }
    private static double[] Offsets(IReadOnlyList<double> sizes) {
        var offsets = new double[sizes.Count + 1];
        for (int i = 0; i < sizes.Count; i++) offsets[i + 1] = offsets[i] + sizes[i];
        return offsets;
    }
    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}

/// <summary>A cell's immutable grid anchor and merged ownership, in definition order.</summary>
internal readonly record struct PublisherTableCellSlot(int Row, int Column, int RowSpan, int ColumnSpan);
