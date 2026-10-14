namespace OfficeIMO.IWork;

public sealed partial class IWorkTable {
    internal IReadOnlyList<Internal.IWorkCellLayout?> CellLayoutDefaults { get; }

    internal IEnumerable<IWorkCellPadding> CellPaddings =>
        CellLayoutDefaults.Select(layout => layout?.Padding).Concat(Cells.Select(cell => cell.Padding))
            .Where(padding => padding != null).Cast<IWorkCellPadding>();

    internal bool HasDefaultVerticalAlignment => CellLayoutDefaults.Any(layout => layout?.VerticalAlignment != null);

    /// <summary>Gets supported effective padding at a one-based position, including unstored empty cells.</summary>
    /// <remarks>A selected padding message replaces all four sides; omitted sides in that message are zero.
    /// Unresolved selected padding suppresses region defaults. Header and footer rows precede header columns.</remarks>
    public IWorkCellPadding? GetPadding(int row, int column) {
        IWorkTableCell? cell = GetCell(row, column);
        if (cell?.HasUnresolvedPadding == true) return null;
        return cell?.Padding ?? GetDefaultCellLayout(row, column)?.Padding;
    }

    /// <summary>Gets supported effective top, middle or bottom alignment at a one-based position, including unstored empty cells.</summary>
    /// <remarks>Unresolved selected alignment suppresses region defaults. Header and footer rows precede header columns.</remarks>
    public IWorkCellVerticalAlignment? GetVerticalAlignment(int row, int column) {
        IWorkTableCell? cell = GetCell(row, column);
        if (cell?.HasUnresolvedVerticalAlignment == true) return null;
        return cell?.VerticalAlignment ?? GetDefaultCellLayout(row, column)?.VerticalAlignment;
    }

    private Internal.IWorkCellLayout? GetDefaultCellLayout(int row, int column) {
        if (row <= HeaderRowCount) return CellLayoutDefaults[1];
        if (row > RowCount - FooterRowCount) return CellLayoutDefaults[3];
        if (column <= HeaderColumnCount) return CellLayoutDefaults[2];
        return CellLayoutDefaults[0];
    }
}
