namespace OfficeIMO.PowerPoint;

public partial class PowerPointSlide {
    private static void ApplyTableAlignments(PowerPointTable table, OfficeTableOptions? options) {
        if (options == null) return;
        for (int column = 0; column < table.ColumnCount && column < options.Alignments.Count; column++) {
            if (options.Alignments[column] is not { } alignment) continue;
            var native = alignment switch {
                OfficeAlignment.Left or OfficeAlignment.Start => PowerPointTextAlignment.Left,
                OfficeAlignment.Center => PowerPointTextAlignment.Center,
                OfficeAlignment.Right or OfficeAlignment.End => PowerPointTextAlignment.Right,
                _ => PowerPointTextAlignment.Justified
            };
            for (int row = 0; row < table.RowCount; row++) table.GetCell(row, column).HorizontalAlignment = native;
        }
    }
}
